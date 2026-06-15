#!/usr/bin/env python3
r"""Pandoc预处理模块 — 将LaTeX源文件预处理为Pandoc可识别的格式

处理期刊自定义命令、TikZ表格占位、display公式占位、引用替换等。
"""
import os, re, sys
from pathlib import Path

# 导入共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import (
    citation_marker,
    match_balanced_braces as _match_balanced_braces,
)

# 导入BBL解析器
from _bbl_parser import _parse_bbl


def _author_to_latex(match):
    r"""将\Author[email]{FirstName}{Surname}转为\author{FirstName Surname}"""
    whole = match.group(0)
    names = re.findall(r'\{([^{}]*)\}', whole)
    if len(names) >= 2:
        first, last = names[-2], names[-1]
        return f'\\author{{{first} {last}}}'
    return f'\\author{{}}'


def _replace_statement_cmd(content, cmd_name, title):
    r"""替换带平衡大括号的statement命令, 如\dataavailability{...}→\section{Data Availability}...

    Strips any leading title-like prefix from inner content to avoid
    duplicate headings (e.g. "Data availability. xxx" → "xxx").
    """
    # 中文等价词映射
    _cn_titles = {
        'Data Availability': '数据可用性',
        'Code Availability': '代码可用性',
        'Code and Data Availability': '代码和数据可用性',
        'Author Contributions': '作者贡献',
        'Competing Interests': '利益冲突',
        'Sample Availability': '样品可用性',
        'Disclaimer': '免责声明',
        'Copyright Statement': '版权声明',
    }
    cn_title = _cn_titles.get(title, '')
    pattern = '\\' + cmd_name + '{'
    result = content
    title_patterns = [re.compile(r'^\s*' + re.escape(title) + r'\s*[.:;。：]\s*', re.IGNORECASE)]
    if cn_title:
        title_patterns.append(re.compile(r'^\s*' + re.escape(cn_title) + r'\s*[.:;。：]\s*'))
    while True:
        idx = result.find(pattern)
        if idx == -1:
            break
        brace_start = idx + len(pattern) - 1  # {的位置
        inner = _match_balanced_braces(result, brace_start)
        end_pos = brace_start + len(inner) + 2  # 包含{}两个字符
        # 去掉inner中开头的重复标题文本（英文或中文）
        cleaned_inner = inner
        for tp in title_patterns:
            cleaned_inner = tp.sub('', cleaned_inner)
        replacement = f'\\section{{{title}}}\n{cleaned_inner}\n'
        result = result[:idx] + replacement + result[end_pos:]
    return result


def _label_to_eq_num(label_key):
    r"""从 \label{eq2-1} 提取公式编号, 返回 '(2-1)' 等格式

    支持的 label 格式: eq1, eq2-1, eq3.1 等
    """
    num_m = re.search(r'eq(\d+)[-\.]?(\d+)?', label_key)
    if num_m:
        if num_m.group(2):
            return f'({num_m.group(1)}-{num_m.group(2)})'
        else:
            return f'({num_m.group(1)})'
    return f'({label_key})'


def _strip_explicit_tag(formula):
    """Return formula without ``\tag`` and its visible Word equation number."""
    tag_match = re.search(r'\\tag\{([^{}]+)\}', formula)
    if not tag_match:
        return formula, ''
    number = f'({tag_match.group(1).strip()})'
    return re.sub(r'\\tag\{[^{}]+\}', '', formula).strip(), number


def _wrap_display_math(result):
    """将display公式环境替换为占位符，收集公式数据

    按文件中实际出现顺序处理, 保证占位符编号 = display_formula_data 索引

    Returns:
        (result, display_formula_data): 替换后的文本和公式数据列表
    """
    display_formula_data = []  # [{latex, eq_num, env}, ...]
    _replacements = []  # [(start, end, replacement_text), ...] 按文件顺序

    # 先收集所有 display 公式环境及其位置, 按文件顺序排列
    _env_pattern = re.compile(
        r'\\begin\{(equation|gather)\}(.*?)\\end\{\1\}',
        re.DOTALL
    )
    _env_matches = list(_env_pattern.finditer(result))

    # 也收集 $$...$$ 公式
    _dd_pattern = re.compile(r'\$\$(.*?)\$\$', re.DOTALL)
    for m in _dd_pattern.finditer(result):
        if not any(m.start() >= em.start() and m.end() <= em.end() for em in _env_matches):
            _env_matches.append(m)

    # 按文件中出现顺序排序
    _env_matches.sort(key=lambda m: m.start())

    # 按文件顺序收集公式数据和替换文本
    for em in _env_matches:
        env_name = em.group(1) if len(em.groups()) >= 1 and em.group(1) else 'bracket'
        body = em.group(2) if len(em.groups()) >= 2 else em.group(1)

        if env_name == 'gather':
            # gather: 拆分为多行独立公式, label归属上一行
            raw_lines = body.split('\n')
            formula_lines = []  # [(公式文本, 编号), ...]
            i = 0
            while i < len(raw_lines):
                stripped = raw_lines[i].strip()
                if not stripped:
                    i += 1; continue
                # 独占一行的label
                if stripped.startswith('\\label'):
                    label_m = re.search(r'\\label\{([^}]+)\}', stripped)
                    if label_m and formula_lines:
                        eq_num = _label_to_eq_num(label_m.group(1))
                        formula_lines[-1] = (formula_lines[-1][0], eq_num)
                    i += 1; continue
                # 行内label: 公式文本末尾的 \label{xxx}
                label_m = re.search(r'\\label\{([^}]+)\}', stripped)
                eq_num = ''
                if label_m:
                    eq_num = _label_to_eq_num(label_m.group(1))
                    stripped = re.sub(r'\\label\{[^}]+\}', '', stripped).strip()
                stripped, tag_num = _strip_explicit_tag(stripped)
                eq_num = tag_num or eq_num
                dslash = '\\' + '\\'
                if stripped.endswith(dslash):
                    stripped = stripped[:-2].rstrip()
                if stripped:
                    formula_lines.append((stripped, eq_num))
                i += 1
            parts = []
            for eq_text, eq_num in formula_lines:
                idx = len(display_formula_data)
                # 每个占位符前后各两个空行，确保Pandoc生成独立段落
                placeholder = f'\n\n\n[DISPLAY_FORMULA_{idx}]\n\n\n'
                parts.append(placeholder)
                display_formula_data.append({'latex': eq_text, 'eq_num': eq_num, 'env': 'gather_line'})
            replacement = ''.join(parts)
            _replacements.append((em.start(), em.end(), replacement))

        elif env_name == 'equation':
            body_stripped = body.strip()
            eq_num = ''
            label_m = re.search(r'\\label\{([^}]+)\}', body_stripped)
            if label_m:
                eq_num = _label_to_eq_num(label_m.group(1))
                body_stripped = re.sub(r'\\label\{[^}]+\}', '', body_stripped).strip()
            body_stripped, tag_num = _strip_explicit_tag(body_stripped)
            eq_num = tag_num or eq_num
            idx = len(display_formula_data)
            # 前后各两个空行，确保Pandoc生成独立段落
            placeholder = f'\n\n\n[DISPLAY_FORMULA_{idx}]\n\n\n'
            display_formula_data.append({'latex': body_stripped, 'eq_num': eq_num, 'env': 'equation'})
            _replacements.append((em.start(), em.end(), placeholder))

        else:  # $$...$$
            body_stripped = body.strip()
            idx = len(display_formula_data)
            placeholder = f'\n\n\n[DISPLAY_FORMULA_{idx}]\n\n\n'
            display_formula_data.append({'latex': body_stripped, 'eq_num': '', 'env': 'bracket'})
            _replacements.append((em.start(), em.end(), placeholder))

    # 从后往前替换, 避免位置偏移
    for start, end, replacement in sorted(_replacements, key=lambda x: x[0], reverse=True):
        result = result[:start] + replacement + result[end:]

    return result, display_formula_data


def _protect_supertables(result):
    """Replace generated supertabular blocks with Word table placeholders."""
    import base64

    tables = []
    pattern = re.compile(
        r'% WORD_SUPERTABLE_BEGIN number=([^\n]*)\n'
        r'% WORD_SUPERTABLE_TIKZ=([A-Za-z0-9+/=]+)\n'
        r'(.*?)% WORD_SUPERTABLE_END',
        re.DOTALL,
    )

    def replace(match):
        block = match.group(3)
        caption = ''
        cap_match = re.search(r'\\tablecaption\{', block)
        if cap_match:
            caption = _match_balanced_braces(block, cap_match.end() - 1)
        tables.append({
            'caption': caption,
            'number': match.group(1).strip(),
            'tikz_body': base64.b64decode(match.group(2)).decode('utf-8'),
            'index': len(tables),
        })
        return f'\\textbf{{[TIKZ_TABLE_{len(tables) - 1}]}}'

    return pattern.sub(replace, result), tables


def _protect_tikz(result, start_index=0):
    """将TikZ表格替换为占位符，收集TikZ表格信息

    Returns:
        (result, tikz_tables): 替换后的文本和TikZ表格列表
    """
    tikz_table_pattern = r'\\begin\{table\*?\}(.*?)\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}(.*?)\\end\{table\*?\}'
    strip_body = r'(?:(?!\\end\{strip\}).)*?'
    strip_table_pattern = (
        r'\\begin\{strip\}(' + strip_body + r')'
        r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
        r'(' + strip_body + r')\\end\{strip\}'
    )
    tikz_tables = []
    table_idx = start_index

    def _make_placeholder(table_env, tikz_body, after_tikz, is_full_width=False):
        nonlocal table_idx

        # 提取caption
        table_full = table_env + after_tikz
        number = ''
        num_m = re.search(r'\\renewcommand\{\\thetable\}\{([^}]+)\}', table_full)
        if num_m:
            number = num_m.group(1).strip()

        caption = ''
        cap_m = re.search(r'\\caption\{', table_full)
        if cap_m:
            caption = _match_balanced_braces(table_full, cap_m.end() - 1)

        # 保存TikZ表格信息
        tikz_tables.append({
            'caption': caption,
            'number': number,
            'tikz_body': tikz_body,
            'is_full_width': is_full_width,
            'index': table_idx,
        })

        # 替换为占位符
        placeholder = f'\\textbf{{[TIKZ_TABLE_{table_idx}]}}'
        table_idx += 1
        return placeholder

    def _replace_strip_tikz_table(match):
        table_env = match.group(1)
        if not re.search(r'\\@captype\s*\{table\}|\\def\s*\\@captype\s*\{table\}', table_env):
            return match.group(0)
        return _make_placeholder(match.group(1), match.group(2), match.group(3), True)

    def _replace_tikz_table(match):
        return _make_placeholder(match.group(1), match.group(2), match.group(3), False)

    result = re.sub(strip_table_pattern, _replace_strip_tikz_table, result, flags=re.DOTALL)
    result = re.sub(tikz_table_pattern, _replace_tikz_table, result, flags=re.DOTALL)
    return result, tikz_tables


def _strip_figure_captions(result):
    """移除figure环境中的\\caption和\\small，避免与后处理重复

    使用平衡大括号匹配，支持caption中嵌套大括号的长文本。
    """
    def _strip_figure_caption(match):
        fig_body = match.group(1)
        # 用循环逐个移除 \caption{...}（含嵌套大括号）
        while True:
            cap_pos = fig_body.find('\\caption{')
            if cap_pos == -1:
                break
            brace_start = cap_pos + len('\\caption')
            inner = _match_balanced_braces(fig_body, brace_start)
            end_pos = brace_start + len(inner) + 2
            fig_body = fig_body[:cap_pos] + fig_body[end_pos:]
        fig_body = re.sub(r'\\small\s*', '', fig_body)
        return f'\\begin{{figure}}{fig_body}\\end{{figure}}'

    result = re.sub(r'\\begin\{figure\*?\}(.*?)\\end\{figure\*?\}', _strip_figure_caption, result, flags=re.DOTALL)
    return result


def _protect_figures(result):
    """Replace figure environments with stable placeholders for python-docx insertion."""
    figure_idx = 0

    def _replace_strip_figure(match):
        nonlocal figure_idx
        body = match.group(1)
        has_captype = re.search(
            r'\\@captype\s*\{figure\}|\\def\s*\\@captype\s*\{figure\}',
            body,
        )
        if not has_captype or '\\includegraphics' not in body:
            return match.group(0)
        placeholder = f'\n\n[FIGURE_{figure_idx}]\n\n'
        figure_idx += 1
        return placeholder

    def _replace_figure(match):
        nonlocal figure_idx
        placeholder = f'\n\n[FIGURE_{figure_idx}]\n\n'
        figure_idx += 1
        return placeholder

    result = re.sub(
        r'\\begin\{strip\}((?:(?!\\end\{strip\}).)*?)\\end\{strip\}',
        _replace_strip_figure,
        result,
        flags=re.DOTALL,
    )
    return re.sub(r'\\begin\{figure\*?\}.*?\\end\{figure\*?\}', _replace_figure, result, flags=re.DOTALL)


def _isolate_float_placeholders(result):
    """Keep figure/table/formula anchors in standalone paragraphs for Word reinsertion."""
    for pattern in (r'(\[FIGURE_\d+\])', r'(\\textbf\{\[TIKZ_TABLE_\d+\]\})', r'(\[DISPLAY_FORMULA_\d+\])'):
        result = re.sub(r'[ \t]*' + pattern + r'[ \t]*', r'\n\n\1\n\n', result)
    return re.sub(r'\n{4,}', '\n\n\n', result)


def _merge_float_adjacent_text(result):
    """合并被表格/图片浮动体截断的连续段落

    LaTeX中表格和图片是浮动体(floating body)，LaTeX编译器会将前后文字
    连续排版在PDF中。但预处理时替换为占位符后，前后的空行会导致Pandoc
    生成独立的Word段落，破坏文本的连续性。

    检测条件:
    - 浮动体前的文本不以句末标点(.!?。！？)结尾 → 未结束
    - 浮动体后的文本不以大写字母/数字/LaTeX命令开头 → 非新段
    - 前后内容均非占位符行(避免不同占位符互相吞噬)
    - 同时满足时，移除占位符前后的空行，合并为一个段落
    """
    sentence_end_chars = set('.!?。！？')

    # 匹配所有浮动体占位符行
    placeholder_re = re.compile(r'^\\textbf\{\[TIKZ_TABLE_\d+\]\}$')

    def _is_placeholder(line_text):
        return bool(placeholder_re.match(line_text.strip()))

    lines = result.split('\n')
    merged = 0
    i = 0
    while i < len(lines):
        stripped = lines[i].strip()
        if not placeholder_re.match(stripped):
            i += 1
            continue

        # 找到占位符行 → 向前找最近的有内容行
        prev_end = i - 1
        blanks_before = 0
        while prev_end >= 0 and not lines[prev_end].strip():
            blanks_before += 1
            prev_end -= 1

        # 向后找最近的有内容行
        next_start = i + 1
        blanks_after = 0
        while next_start < len(lines) and not lines[next_start].strip():
            blanks_after += 1
            next_start += 1

        # 需要前后都有内容行，且占位符前后有空行
        if (prev_end < 0 or next_start >= len(lines)
                or blanks_before < 1 or blanks_after < 1):
            i += 1
            continue

        prev_content = lines[prev_end].strip()
        next_content = lines[next_start].strip()

        # 前后内容都不能是占位符行(避免不同占位符合并后被一起删除)
        if _is_placeholder(prev_content) or _is_placeholder(next_content):
            i += 1
            continue

        prev_last = lines[prev_end].rstrip()[-1] if lines[prev_end].rstrip() else ''
        next_first = lines[next_start].lstrip()[0] if lines[next_start].lstrip() else ''

        should_merge = (
            prev_last not in sentence_end_chars and      # 前段未结束
            not next_first.isupper() and                  # 非大写开头
            not next_first.isdigit() and                  # 非数字开头
            next_first != '\\' and                        # 非LaTeX命令
            next_first != '#'                             # 非Markdown标题
        )

        if should_merge:
            # 合并: 前内容 + 占位符 + 后内容 在一行
            merged_line = lines[prev_end].rstrip() + ' ' + stripped + ' ' + lines[next_start].lstrip()
            lines[prev_end] = merged_line
            # 删除 prev_end+1 到 next_start 的所有行
            del lines[prev_end + 1: next_start + 1]
            merged += 1
            # 不递增i，重新检查当前位置
            continue

        i += 1

    if merged:
        print(f'  [float-merge] 合并了 {merged} 处浮动体截断段落')
    return '\n'.join(lines)


def _extract_abstract(result):
    """提取abstract环境内容（预留扩展）"""
    return result


def prepare_tex_for_pandoc(tex_path, work_dir, bbl_path=None):
    """预处理tex文件, 处理期刊自定义命令使其Pandoc可识别

    Args:
        tex_path: LaTeX源文件路径
        work_dir: 工作目录
        bbl_path: .bbl文件路径(可选), 注入参考文献并替换citep
    """
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    result = content

    # === 第零步: bbl注入 + citep替换 ===
    cite_map = {}
    if bbl_path and os.path.exists(bbl_path):
        bbl_info = _parse_bbl(bbl_path)
        cite_map = bbl_info['cite_map']
        numbered_cites = bool(cite_map) and all(
            re.fullmatch(r'\d+', str(value)) for value in cite_map.values())
        print(f'  [bbl] 解析到 {len(cite_map)} 条参考文献 (author-year格式)')

        # 替换引用为author-year格式
        # 重要策略: 多key的\citep{key1,key2}拆分为独立引用
        # 这样 cross_ref_builder 可以精确匹配每个引用并插入HYPERLINK跳转
        def _citation_token(key, mode):
            if key not in cite_map:
                return key
            return citation_marker(key, mode)

        def _replace_citep(match):
            keys_str = match.group(1)
            keys = [k.strip() for k in keys_str.split(',')]
            if numbered_cites:
                markers = [_citation_token(key, 'M') for key in keys]
                return '[' + ','.join(markers) + ']'
            if len(keys) == 1:
                key = keys[0]
                if key not in cite_map:
                    return f'({key})'
                return _citation_token(key, 'P')
            # 多引用保留外层括号，每个key独立生成跳转标记。
            return '(' + '; '.join(_citation_token(k, 'M') for k in keys) + ')'

        def _replace_citet(match):
            keys_str = match.group(1)
            keys = [k.strip() for k in keys_str.split(',')]
            if numbered_cites:
                return '[' + ','.join(_citation_token(key, 'M') for key in keys) + ']'
            return '; '.join(_citation_token(k, 'T') for k in keys)

        result = re.sub(r'\\citep\{([^}]+)\}', _replace_citep, result)
        result = re.sub(r'\\citet\{([^}]+)\}', _replace_citet, result)
        result = re.sub(r'\\cite\{([^}]+)\}', _replace_citep, result)

        # Keep a stable marker; python-docx builds the bibliography with bookmarks.
        result = re.sub(r'\\bibliographystyle\{[^}]*\}\s*', '', result)
        result = re.sub(r'\\bibliography\{[^}]*\}', '[REFERENCES_PLACEHOLDER]', result)

        print('  [bbl] citep/citet replaced; References placeholder inserted')
    else:
        result = re.sub(r'\\citep\{([^}]+)\}', r'\\cite{\1}', result)
        result = re.sub(r'\\citet\{([^}]+)\}', r'\\cite{\1}', result)

    # === 第一步: 处理statement命令(带平衡大括号的嵌套内容) ===
    statement_cmds = [
        ('authorcontribution', 'Author Contributions'),
        ('competinginterests', 'Competing Interests'),
        ('dataavailability', 'Data Availability'),
        ('codeavailability', 'Code Availability'),
        ('codedataavailability', 'Code and Data Availability'),
        ('sampleavailability', 'Sample Availability'),
        ('disclaimer', 'Disclaimer'),
        ('copyrightstatement', 'Copyright Statement'),
    ]
    for cmd, title in statement_cmds:
        result = _replace_statement_cmd(result, cmd, title)

    # === 第二步: 将TikZ表格替换为占位符 ===
    result, super_tables = _protect_supertables(result)
    result, tikz_tables = _protect_tikz(result, start_index=len(super_tables))
    tikz_tables = super_tables + tikz_tables

    # === 第二步b: 将figure环境中的\caption和\small移除 ===
    # 必须在 _protect_figures 之前执行，否则figure环境已被替换为占位符，
    # caption文本会残留在Pandoc输出中，导致Word中出现重复图例
    result = _strip_figure_captions(result)

    # === 第二步c: 将figure环境替换为稳定占位符 ===
    result = _protect_figures(result)

    # === 第二步d: 合并被浮动体截断的连续段落 ===
    result = _merge_float_adjacent_text(result)
    result = _isolate_float_placeholders(result)

    # === 第三步: 正则替换(简单模式) ===
    simple_replacements = [
        (r'^\\introduction\b', r'\\section{Introduction}'),
        (r'^\\conclusions\b', r'\\section{Conclusions}'),
        (r'\\begin\{acknowledgements\}', r'\\begin{acknowledgment}'),
        (r'\\end\{acknowledgements\}', r'\\end{acknowledgment}'),
        (r'\\Author(?:\[[^\]]*\]){0,2}\{[^}]*\}\{[^}]*\}', _author_to_latex),
        (r'\\affil\[(.*?)\]\{(.*?)\}', r'% \\affil[\1]{\2}'),
        (r'\\affil\{(.*?)\}', r'% \\affil{\1}'),
        (r'^\\runningtitle\{', r'% \\runningtitle{'),
        (r'^\\runningauthor\{', r'% \\runningauthor{'),
        (r'^\\received\{', r'% \\received{'),
        (r'^\\pubdiscuss\{', r'% \\pubdiscuss{'),
        (r'^\\revised\{', r'% \\revised{'),
        (r'^\\accepted\{', r'% \\accepted{'),
        (r'^\\published\{', r'% \\published{'),
        (r'^\\firstpage\{', r'% \\firstpage{'),
        (r'^\\noappendix\b', r'% \\noappendix'),
        (r'^\\appendixfigures\b', r'% \\appendixfigures'),
        (r'^\\appendixtables\b', r'% \\appendixtables'),
        (r'^\\correspondence\{', r'% \\correspondence{'),
    ]

    for pattern, replacement in simple_replacements:
        if callable(replacement):
            result = re.sub(pattern, replacement, result, flags=re.MULTILINE)
        else:
            result = re.sub(pattern, replacement, result, flags=re.MULTILINE)

    # === 第3.5步: 为section/subsection注入编号前缀 ===
    # Pandoc转Word时不带自动编号，需要手动注入以匹配PDF编译结果
    sec_counter = 0
    subsec_counter = 0
    lines = result.split('\n')
    for i, line in enumerate(lines):
        sec_m = re.match(r'^\\section\{(.+)\}$', line)
        subsec_m = re.match(r'^\\subsection\{(.+)\}$', line)
        if sec_m:
            sec_counter += 1
            subsec_counter = 0
            title = sec_m.group(1)
            lines[i] = f'\\section{{{sec_counter} {title}}}'
        elif subsec_m:
            subsec_counter += 1
            title = subsec_m.group(1)
            lines[i] = f'\\subsection{{{sec_counter}.{subsec_counter} {title}}}'
    result = '\n'.join(lines)

    # 修正bibliography路径（仅当bbl未注入时生效）
    if not (bbl_path and os.path.exists(bbl_path)):
        result = re.sub(r'\\bibliography\{[^}]*\}', r'\\bibliography{references}', result)
        result = re.sub(r'\\bibliographystyle\{[^}]*\}', r'\\bibliographystyle{copernicus}', result)

    # === 第三步b: 将 \textsubscript{N} 转为 Unicode 下标 ===
    _sub_map = str.maketrans('0123456789+-=()aeghijklmnoprstuvx',
                              '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₑₕᵢⱼₖₗₘₙₒₚᵣₛₜᵤᵥₓ')
    def _textsub_to_unicode(m):
        return m.group(1).translate(_sub_map)
    # 仅替换文本模式(非$...$内)的 \textsubscript
    result = re.sub(r'\\textsubscript\{([^}]+)\}', _textsub_to_unicode, result)

    # === 第三步c: 将 文本$_{N}$ 模式的inline下标转为Unicode ===
    # Pandoc把 "CO$_{2}$" 拆成 "CO"文本 + OMML下标，OMML含零宽空格导致显示异常
    # 解决方案：直接转为 Unicode 下标 (如 CO₂)，避免 Pandoc 产生问题 OMML
    def _inline_sub_to_unicode(m):
        prefix = m.group(1)       # CO, XCO 等前缀
        sub_content = m.group(2)  # 下标内容 (2, 3 等)
        return prefix + sub_content.translate(_sub_map)
    def _inline_sub_to_unicode_nobrace(m):
        prefix = m.group(1)
        sub_content = m.group(2)
        return prefix + sub_content.translate(_sub_map)
    # 匹配: 字母/数字后紧跟 $_{内容}$ 或 $_N$（仅限简单下标，不含复杂公式）
    result = re.sub(r'([A-Za-z]{1,5})\$_{([^}]+)}\$', _inline_sub_to_unicode, result)
    result = re.sub(r'([A-Za-z]{1,5})\$_([0-9])\$', _inline_sub_to_unicode_nobrace, result)
    # 匹配: 任意位置独立的 $_{N}$ 或 $_N$ 也转为Unicode
    result = re.sub(r'\$_{([0-9]+)}\$', lambda m: m.group(1).translate(_sub_map), result)
    result = re.sub(r'\$_([0-9])\$', lambda m: m.group(1).translate(_sub_map), result)

    # === 第四步: display公式替换为占位符 ===
    result, display_formula_data = _wrap_display_math(result)

    # 写入工作目录
    prepared_path = os.path.join(work_dir, 'prepared.tex')
    with open(prepared_path, 'w', encoding='utf-8') as f:
        f.write(result)

    return prepared_path, result, tikz_tables, display_formula_data, cite_map
