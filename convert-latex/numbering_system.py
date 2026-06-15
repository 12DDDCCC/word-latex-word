#!/usr/bin/env python3
"""章节编号约束系统

编号模式:
  'simple'    — 纯数字编号: Figure 1, Table 2, (3)     (manuscript/非final模式)
  'sectioned' — 章节编号:   Figure 2.1, Table 3.2, (2-1) (final模式)
"""

import re
from pathlib import Path

# 参考文献关键词（用于检测参考文献段落起始位置）
REF_KEYWORDS = ['references', 'reference', '参考文献', 'bibliography', '参考文献列表']


def detect_template_numbering_mode(template_dir, layout_spec=None, doc_options=None):
    """检测模板的编号模式

    考虑 documentclass 选项（manuscript/final），
    copernicus.cls 中 \@addtoreset 在 \\if@stage@final 条件块内，
    manuscript 模式下不生效。

    Args:
        template_dir: 模板目录路径
        layout_spec: 排版规格（含 numbering 字段）
        doc_options: documentclass 选项列表，如 ['acp', 'manuscript']

    Returns:
        str: 'simple' 或 'sectioned'
    """
    doc_options = doc_options or []

    # 1. 优先从 layout_spec 的 numbering 字段获取
    if layout_spec:
        numbering = layout_spec.get('numbering', {})
        # 检查无条件重置（不受条件块约束）
        # 注意：equation 可能在 appendix 块内重置，不影响正文编号
        # 优先检查 figure/table 的重置
        unconditional = numbering.get('_unconditional_resets', [])
        for counter, parent in unconditional:
            if parent == 'section' and counter in ('figure', 'table'):
                return 'sectioned'
        # 检查条件重置 — 仅在 final 选项下生效
        conditional = numbering.get('_conditional_resets', [])
        for counter, parent, cond_name in conditional:
            if parent == 'section' and counter in ('figure', 'table'):
                if 'final' in cond_name.lower() and 'final' in doc_options:
                    return 'sectioned'
        # 从 numbering_format 字段判断（需结合 doc_options）
        is_final = 'final' in doc_options
        for key in ('figure_format', 'table_format', 'equation_format'):
            fmt = numbering.get(key, '')
            if fmt and 'thesection' in fmt and is_final:
                return 'sectioned'

    # 2. 从.cls文件检测
    cls_files = list(Path(template_dir).glob('*.cls'))
    for cls_file in cls_files:
        try:
            content = cls_file.read_text(encoding='utf-8', errors='ignore')
        except Exception:
            continue
        # 检测 \if@stage@final 条件块内的 \@addtoreset
        cond_reset = re.search(
            r'\\if@stage@final\s*.*?\\@addtoreset\s*\{(?:figure|table|equation)\}\s*\{section\}',
            content, re.DOTALL)
        if cond_reset:
            if 'final' in doc_options:
                return 'sectioned'
            else:
                return 'simple'
        # 无条件 \@addtoreset
        for counter in ('figure', 'table', 'equation'):
            if re.search(r'\\@addtoreset\{' + counter + r'\}\{section\}', content):
                return 'sectioned'
        # \thefigure 定义中包含 \thesection
        for name in ('figure', 'table', 'equation'):
            m = re.search(r'\\def\\the' + name + r'\s*\{([^}]+)\}', content)
            if m and 'thesection' in m.group(1):
                return 'sectioned'

    return 'simple'


def apply_numbering_system(tex_content, ref_map=None, template_numbering='simple',
                           source_numbering='simple', clean_caption_prefix_in_tex=None):
    """编号系统统一入口：\ref替换 + caption清理 + 编号模式转换

    Args:
        tex_content: LaTeX文本
        ref_map: 编号→label映射
        template_numbering: 模板编号模式 ('simple'/'sectioned')
        source_numbering: 源文档编号模式 ('simple'/'sectioned')
        clean_caption_prefix_in_tex: 清理caption前缀的函数（从shared导入）
    """
    # 1. 替换硬编码编号为 \ref/\eqref
    if ref_map:
        print(f"[编号系统] ref_map条目数: {len(ref_map)}")
        fig_items = {k: v for k, v in ref_map.items() if v.startswith('fig')}
        print(f"[编号系统] fig映射: {fig_items}")
        tex_content = convert_numbering_references(tex_content, ref_map)

    # 2. 清理caption中重复的编号前缀
    if clean_caption_prefix_in_tex:
        tex_content = clean_caption_prefix_in_tex(tex_content)

    # 3. 当源文档sectioned但模板要求simple时，转换残留的sectioned编号
    if source_numbering == 'sectioned' and template_numbering == 'simple':
        tex_content = renumber_sectioned_to_simple(tex_content, ref_map=ref_map)
        print("[编号转换] sectioned→simple: 图2.1→图1, 表3.1→表1, (2-1)→(1)")

    return tex_content


def detect_source_numbering_mode(paragraphs):
    """检测源Word文档中的编号模式

    从段落文本中提取图/表/公式编号格式，判断是纯数字还是章节编号。
    只要存在图/表编号含小数点(如 图2.1)即判定为sectioned，
    因为simple编号不会出现小数点。

    Returns:
        str: 'simple' 或 'sectioned'
    """
    sectioned_count = 0
    simple_count = 0

    for p in paragraphs:
        text = p.get('text', '').strip()
        if not text:
            continue
        # 图编号: "Figure 2.1" / "图2.1" → sectioned
        if re.search(r'(?:Figure|Fig\.?|图)\s*\d+\.\d+', text, re.IGNORECASE):
            sectioned_count += 1
        elif re.search(r'(?:Figure|Fig\.?|图)\s*\d+(?!\.\d)', text, re.IGNORECASE):
            simple_count += 1
        # 表编号: "Table 3.1" → sectioned
        if re.search(r'(?:Table|表)\s*\d+\.\d+', text, re.IGNORECASE):
            sectioned_count += 1
        elif re.search(r'(?:Table|表)\s*\d+(?!\.\d)', text, re.IGNORECASE):
            simple_count += 1
        # 公式编号: "(2-1)" / "(2.1)" → sectioned
        if re.search(r'\(\d+[-\.]\d+\)', text):
            sectioned_count += 1

    # sectioned编号(图2.1)是明确信号，simple编号(图3)可能是sectioned的简写引用
    # 所以sectioned只要>0就优先判定
    if sectioned_count > 0:
        return 'sectioned'
    return 'simple'


def convert_numbering_references(text, ref_map=None):
    """将正文中的硬编码图/表/公式编号替换为LaTeX的\\ref{}/\\eqref{}命令

    只替换ref_map中已建立映射的编号，未映射的保留原样。
    对于中文图/表引用，保留前缀（图/表），只替换编号部分。
    使用正则确保完整匹配数字边界，避免 图2 匹配到 图2.1 的前半部分。

    ref_map: dict, 源文档编号→label名映射，例如:
        {'图2.1': 'fig1', 'Figure 2.1': 'fig1', '(3)': 'eq2-3', '表2': 'tab1'}
    """
    if not ref_map:
        return text

    # 按编号长度降序排列，避免短编号先匹配导致长编号被截断
    sorted_refs = sorted(ref_map.items(), key=lambda x: len(x[0]), reverse=True)

    for src_num, label in sorted_refs:
        if src_num.startswith('(') and src_num.endswith(')'):
            # 公式编号: (3) → \eqref{eq2-3}
            pattern = r'(?<![\\\w\{])' + re.escape(src_num) + r'(?![\w\}])'
            text = re.sub(pattern, lambda _m: f'\\eqref{{{label}}}', text)
        elif src_num.startswith('图') or src_num.startswith('表'):
            # 中文图/表引用: 图2.1 → 图\ref{fig1}（保留前缀）
            # 用正则确保数字后面不是点号+数字（避免 图2 匹配 图2.1）
            num_part = re.sub(r'^[图表]\s*', '', src_num)
            # 构造正则：匹配 图2.1 但不匹配 图2.1x (如果后面还有数字)
            pattern = re.escape(src_num[:1]) + r'\s*' + re.escape(num_part) + r'(?!\.\d)'
            text = re.sub(pattern, f'{src_num[:1]}\\\\ref{{{label}}}', text)
        elif src_num.startswith('Figure') or src_num.startswith('Fig'):
            # 英文图引用: Figure 2.1 → Figure~\ref{fig1}
            prefix = src_num.split()[0]
            num_part = src_num[len(prefix):].strip()
            pattern = re.escape(prefix) + r'\s*' + re.escape(num_part) + r'(?!\.\d)'
            text = re.sub(pattern, f'{prefix}~\\\\ref{{{label}}}', text)
        elif src_num.startswith('Table') or src_num.startswith('Talbe'):
            # 英文表引用: Table 2 → Table~\ref{tab1}
            prefix = src_num.split()[0]
            num_part = src_num[len(prefix):].strip()
            pattern = re.escape(prefix) + r'\s*' + re.escape(num_part) + r'(?!\.\d)'
            text = re.sub(pattern, f'{prefix}~\\\\ref{{{label}}}', text)

    return text


def renumber_sectioned_to_simple(text, ref_map=None):
    """将tex文本中残留的sectioned编号转换为simple编号

    当模板要求simple编号但源文档使用sectioned编号时，在\ref替换之后调用。
    处理\ref{}替换未覆盖的sectioned编号（如图例段落、正文漏网引用）。

    利用ref_map中的已知映射(图2.1→fig1意味着图2.1→图1)作为基准，
    未映射的sectioned编号按出现顺序追加分配simple编号。

    转换规则:
    - 中文: 图2.1→图1, 表3.1→表1 (保留子图字母: 图4.3d→图5d)
    - 英文: Figure 2.1→Figure 1, Table 3.1→Table 1
    - 公式: (2-1)→(1), (2.1)→(1)
    """
    # 从ref_map提取已知映射: 图2.1→1, 图4.3→5 等
    fig_map = {}    # sectioned_num -> simple_num (e.g. "2.1" -> "1")
    tab_map = {}
    eq_map = {}
    fig_max = 0
    tab_max = 0
    eq_max = 0

    if ref_map:
        for src, label in ref_map.items():
            # Figure: 图2.1 -> fig1, Figure 2.1 -> fig2-1
            m = re.match(r'(?:图|Figure|Fig\.?\s*)(\d+\.\d+)', src, re.IGNORECASE)
            if m and label.startswith('fig'):
                num_str = m.group(1)  # "2.1"
                lm = re.match(r'fig(\d+)', label)
                if lm:
                    fig_map[num_str] = lm.group(1)
                    fig_max = max(fig_max, int(lm.group(1)))
                continue
            # Table: 表3.1 -> tab1
            m = re.match(r'(?:表|Table)\s*(\d+\.\d+)', src, re.IGNORECASE)
            if m and label.startswith('tab'):
                num_str = m.group(1)
                lm = re.match(r'tab(\d+)', label)
                if lm:
                    tab_map[num_str] = lm.group(1)
                    tab_max = max(tab_max, int(lm.group(1)))
                continue
            # Equation: (2-1) -> eq2-1
            m = re.match(r'\((\d+[-\.]\d+)\)', src)
            if m and label.startswith('eq'):
                num_str = m.group(1)
                lm = re.match(r'eq(\d+[-]?\d*)', label)
                if lm:
                    eq_map[num_str] = lm.group(1)
                    first_num = re.match(r'\d+', lm.group(1))
                    if first_num:
                        eq_max = max(eq_max, int(first_num.group()))

    # 图编号: 中文 "图2.1" / "图2.1d"
    def _replace_fig_cn(m):
        nonlocal fig_max
        prefix = m.group(1)  # "图" or "图 "
        sec_num = m.group(2)  # "2.1"
        sub_letter = m.group(3) or ''  # "d" (子图字母)
        if sec_num in fig_map:
            return f'{prefix}{fig_map[sec_num]}{sub_letter}'
        fig_max += 1
        fig_map[sec_num] = str(fig_max)
        return f'{prefix}{fig_map[sec_num]}{sub_letter}'
    text = re.sub(r'(图\s*)(\d+\.\d+)([a-zA-Z]?)', _replace_fig_cn, text)

    # 英文图: Figure 2.1 / Fig. 2.1
    def _replace_fig_en(m):
        nonlocal fig_max
        prefix = m.group(1)
        sec_num = m.group(2)
        sub_letter = m.group(3) or ''
        if sec_num in fig_map:
            return f'{prefix}{fig_map[sec_num]}{sub_letter}'
        fig_max += 1
        fig_map[sec_num] = str(fig_max)
        return f'{prefix}{fig_map[sec_num]}{sub_letter}'
    text = re.sub(r'(Figure\s+|Fig\.?\s+)(\d+\.\d+)([a-zA-Z]?)', _replace_fig_en, text, flags=re.IGNORECASE)

    # 表编号: 中文 "表3.1"
    def _replace_tab_cn(m):
        nonlocal tab_max
        prefix = m.group(1)
        sec_num = m.group(2)
        if sec_num in tab_map:
            return f'{prefix}{tab_map[sec_num]}'
        tab_max += 1
        tab_map[sec_num] = str(tab_max)
        return f'{prefix}{tab_map[sec_num]}'
    text = re.sub(r'(表\s*)(\d+\.\d+)', _replace_tab_cn, text)

    # 表编号: 英文 "Table 3.1"
    def _replace_tab_en(m):
        nonlocal tab_max
        prefix = m.group(1)
        sec_num = m.group(2)
        if sec_num in tab_map:
            return f'{prefix}{tab_map[sec_num]}'
        tab_max += 1
        tab_map[sec_num] = str(tab_max)
        return f'{prefix}{tab_map[sec_num]}'
    text = re.sub(r'(Table\s+)(\d+\.\d+)', _replace_tab_en, text, flags=re.IGNORECASE)

    # 公式编号: (2-1) → (1), (2.1) → (1)
    def _replace_eq(m):
        nonlocal eq_max
        sec_num = m.group(1)
        if sec_num in eq_map:
            return f'({eq_map[sec_num]})'
        eq_max += 1
        eq_map[sec_num] = str(eq_max)
        return f'({eq_map[sec_num]})'
    text = re.sub(r'(?<![\\\w\{])\((\d+[-\.]\d+)\)(?![\w\}])', _replace_eq, text)

    return text


def generate_equation_label(eq_counter, current_section, numbering_mode):
    """根据编号模式生成公式label

    Args:
        eq_counter: 全局公式计数器
        current_section: 当前section编号 (int)
        numbering_mode: 'simple' 或 'sectioned'

    Returns:
        str: LaTeX label, 如 '\\label{eq1}' 或 '\\label{eq2-1}'
    """
    if numbering_mode == 'sectioned' and current_section:
        return f'\\label{{eq{current_section}-{eq_counter}}}'
    else:
        return f'\\label{{eq{eq_counter}}}'
