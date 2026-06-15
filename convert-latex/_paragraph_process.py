#!/usr/bin/env python3
"""段落遍历处理逻辑

从 assemble_tex.py 提取的段落遍历核心逻辑，包含：
- _process_paragraph_loop: 段落遍历主循环
- _process_heading_paragraph: 标题段落处理
- _process_formula_paragraph: 公式段落处理
- _process_empty_paragraph: 空段落处理
- _process_unknown_paragraph: 未知类型段落处理
"""

import re
import sys
import unicodedata
from pathlib import Path

# 包内绝对导入
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from template_spec_extract import APPENDIX_KEYWORDS
from _text_helpers import _strip_heading_number
from _image_table_insert import _insert_images_and_tables


def _should_insert_float_barrier(layout_spec):
    doc_spec = (layout_spec or {}).get('document', {})
    page_spec = (layout_spec or {}).get('page_geometry', {})
    try:
        column_count = int(page_spec.get('column_count', 1) or 1)
    except (TypeError, ValueError):
        column_count = 1
    if doc_spec.get('is_twocolumn') and column_count < 2:
        column_count = 2
    return not (
        column_count > 1 and doc_spec.get('supports_double_column_floats')
    )


def _clean_formula_latex(eq_latex):
    """Remove converter artifacts without changing the source equation intent."""
    if not eq_latex:
        return ''
    eq_latex = unicodedata.normalize('NFKC', eq_latex)
    eq_latex = eq_latex.replace('−', '-').replace('×', r'\times ')
    eq_latex = re.sub(r'\\times(?=[A-Za-z\\])', r'\\times ', eq_latex)
    eq_latex = re.sub(
        r'\\(?:mathbf|boldsymbol|mathrm)\{\\times\s*([A-Za-z])\}',
        r'\\times \1',
        eq_latex,
    )
    eq_latex = re.sub(
        r'\\mathrm\{\s*([+\-])\s*(\\[A-Za-z]+)\\times\s*\}',
        r'\1\2\\times ',
        eq_latex,
    )
    eq_latex = re.sub(r'\\mathrm\{(\\[A-Za-z]+)\}', r'\1', eq_latex)
    eq_latex = eq_latex.replace('\ufffc', '').replace('\ufffd', '')
    eq_latex = eq_latex.replace('$', '')
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf|textbf)\{\s*\}', '', eq_latex)
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf)\{\\times\s*([+\-])\}', r'\\times \1', eq_latex)
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf)\{\\times\s*(\\[A-Za-z]+)\}', r'\\times \1', eq_latex)
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf)\{\\times\}', r'\\times', eq_latex)
    eq_latex = re.sub(r'\\times\s*([+\-])', r'\1', eq_latex)
    greek_cmds = 'alpha|beta|gamma|delta|epsilon|varepsilon|zeta|eta|theta|vartheta|iota|kappa|lambda|mu|nu|xi|pi|rho|sigma|tau|upsilon|phi|varphi|chi|psi|omega'
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf)\{\\(' + greek_cmds + r')\}', r'\\\1', eq_latex)
    eq_latex = re.sub(r'\\(?:boldsymbol|mathbf)\{([=+\-])\}', r'\1', eq_latex)
    eq_latex = re.sub(r'\\textbf\{\s*\}', '', eq_latex)
    eq_latex = eq_latex.replace('\\mathrm{CO2}', '\\mathrm{CO}_{2}')
    eq_latex = eq_latex.replace('\\mathrm{X}\\mathrm{CO2}', '\\mathrm{XCO}_{2}')
    eq_latex = eq_latex.replace('\\mathrm{CO₂}', '\\mathrm{CO}_{2}')
    eq_latex = eq_latex.replace('\\mathrm{X}\\mathrm{CO₂}', '\\mathrm{XCO}_{2}')
    eq_latex = eq_latex.replace('\\mathrm{CO\\textsubscript{2}}', '\\mathrm{CO}_{2}')
    eq_latex = eq_latex.replace('\\mathrm{X}\\mathrm{CO\\textsubscript{2}}', '\\mathrm{XCO}_{2}')
    eq_latex = re.sub(r'\s+', ' ', eq_latex)
    return eq_latex.strip()


def _looks_like_float_caption_text(text, runs=None, doc_meta=None):
    """Detect explicit figure/table captions missed by style-based extraction.

    Uses font-size as the primary discriminator: captions are typically in a
    smaller font than body text (e.g. 9pt vs 12pt). Length alone is unreliable
    because real captions (e.g. Figure 4.1 with sub-figure descriptions) can
    be very long, while body text referencing "Table N." is short or long.

    Returns True only when the text matches caption pattern AND font size
    confirms it's smaller than body text. Without font-size data, only short
    text (<200 chars) is accepted as a heuristic fallback.
    """
    text = (text or '').strip()
    if not text:
        return False
    if not (re.match(r'^(?:Figure|Fig\.?|Table|Tab\.?)\s*\d+(?:\.\d+)*\s*[.:：．]\s+', text, re.I)
            or re.match(r'^(?:图|表)\s*\d+(?:\.\d+)*\s*[.:：．、]\s*\S+', text)):
        return False

    # 字号判断：caption字号 < 正文字号 → 确认为caption
    # 阈值用 strict <（不带-1），因为10pt vs 11pt也只差1pt
    if runs and doc_meta:
        _sizes = [r['size_pt'] for r in runs
                  if isinstance(r, dict) and r.get('size_pt') is not None]
        if _sizes:
            from collections import Counter
            dom_size = Counter(_sizes).most_common(1)[0][0]
            body_size = doc_meta.get('_body_font_size', 12.0)
            if dom_size < body_size:
                return True   # 字号小于正文 → caption
            else:
                return False  # 字号=正文 → body text引用表格/图片

    # 无字号信息时的兜底：短文本(<200)认为是caption
    return len(text) <= 200


def _process_heading_paragraph(p, body_lines, skeleton_info, spec, _concl_kw,
                               current_section, section_eq_counter, _has_placeins,
                               img_insert_map, tbl_insert_map, inserted_img_files,
                               inserted_tbl_ids, layout_spec, ref_map, fig_counter,
                               tab_counter, template_numbering):
    """处理标题段落

    Args:
        p: 段落信息字典
        body_lines: 输出行列表
        skeleton_info: 模板骨架信息
        spec: 模板规格
        _concl_kw: 结论关键词集合
        current_section: 当前section编号
        section_eq_counter: section内公式计数器
        _has_placeins: 是否加载了placeins包
        其他参数: 传递给 _insert_images_and_tables

    Returns:
        tuple: (current_section, section_eq_counter) 更新后的值
    """
    level = p['heading_level']
    latex = p['latex']
    text = p['text']
    pi = p['para_index']

    clean_latex = _strip_heading_number(latex, text)

    if level == 1:
        # 追踪section编号（用于公式label）
        current_section += 1
        section_eq_counter = 0  # 重置section内公式计数器
        heading_text_lower = clean_latex.lower().strip()
        # 声明标题检测：Data Availability / Competing Interests 等
        # 这些标题应该使用模板的专用命令而非 \section{}
        _heading_decl_map = {
            'data availability': 'dataavailability',
            'code availability': 'codeavailability',
            'code and data availability': 'codedataavailability',
            'author contributions': 'authorcontribution',
            'author contribution': 'authorcontribution',
            'competing interests': 'competinginterests',
            'competing interest': 'competinginterests',
            'sample availability': 'sampleavailability',
            'acknowledgements': 'acknowledgements',
            'acknowledgment': 'acknowledgements',
        }
        _decl_cmd_name = _heading_decl_map.get(heading_text_lower)
        if _decl_cmd_name and _decl_cmd_name in skeleton_info.get('statement_cmds', {}):
            cmd = f'\\{_decl_cmd_name}'
            body_lines.append(f'{cmd}{{{skeleton_info["statement_cmds"][_decl_cmd_name]}}}')
            body_lines.append('')
            # 标题后也检查是否有图片/表格需插入
            _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                      body_lines, inserted_img_files, inserted_tbl_ids,
                                      layout_spec=layout_spec, ref_map=ref_map,
                                      fig_counter=fig_counter, tab_counter=tab_counter,
                                      current_section=current_section,
                                      template_numbering=template_numbering)
            return current_section, section_eq_counter
        # 使用模板中提取的特殊命令（动态）
        if skeleton_info['conclusions_cmd'] and heading_text_lower in _concl_kw:
            body_lines.append(skeleton_info['conclusions_cmd'])
        else:
            intro_cmd = skeleton_info['introduction_cmd'] or '\\section'
            first_section_found = any(
                bl.startswith('\\section{') or bl.startswith('\\introduction')
                for bl in body_lines
            )
            if not first_section_found and skeleton_info['introduction_cmd']:
                body_lines.append(skeleton_info['introduction_cmd'])
            else:
                # 附录检测: 遇到"附录"/"Appendix"关键词时，先插入\appendix声明
                if heading_text_lower in APPENDIX_KEYWORDS and not any('\\appendix' in bl for bl in body_lines):
                    app_fmt = spec.get('appendix_format', {}) if spec else {}
                    if app_fmt.get('type') == 'environment':
                        body_lines.append('\\begin{appendices}')
                    else:
                        body_lines.append('\\appendix')
                body_lines.append(f'\\section{{{clean_latex}}}')
    elif level == 2:
        body_lines.append(f'\\subsection{{{clean_latex}}}')
    elif level == 3:
        body_lines.append(f'\\subsubsection{{{clean_latex}}}')
    elif level == 4:
        body_lines.append(f'\\paragraph{{{clean_latex}}}')

    # Keep barriers out of two-column starred-float templates; otherwise
    # LaTeX has too few legal pages for table*/figure* placement.
    if level <= 2 and _has_placeins and _should_insert_float_barrier(layout_spec):
        body_lines.append('\\FloatBarrier')
    body_lines.append('')

    # 标题后也检查是否有图片/表格需插入
    _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                              body_lines, inserted_img_files, inserted_tbl_ids,
                              layout_spec=layout_spec, ref_map=ref_map,
                              fig_counter=fig_counter, tab_counter=tab_counter,
                              current_section=current_section,
                              template_numbering=template_numbering)

    return current_section, section_eq_counter


def _insert_label_before_outer_math_end(latex, label):
    """Insert a label before the outermost trailing math environment end."""
    if not label:
        return latex
    matches = list(re.finditer(r'\\end\{', latex))
    if not matches:
        return latex + '\n' + label
    pos = matches[-1].start()
    return latex[:pos] + f'  {label}\n' + latex[pos:]


def _process_formula_paragraph(p, body_lines, ref_map, eq_counter, used_eq_labels,
                               current_section, section_eq_counter, template_numbering,
                               img_insert_map, tbl_insert_map, inserted_img_files,
                               inserted_tbl_ids, layout_spec, fig_counter, tab_counter):
    """处理独立公式段落

    Args:
        p: 段落信息字典
        body_lines: 输出行列表
        ref_map: 编号映射字典
        eq_counter: 公式计数器
        used_eq_labels: 已使用的公式label集合
        current_section: 当前section编号
        section_eq_counter: section内公式计数器
        template_numbering: 编号模式
        其他参数: 传递给 _insert_images_and_tables

    Returns:
        tuple: (eq_counter, section_eq_counter) 更新后的值
    """
    pi = p['para_index']
    text = p['text']
    latex = p['latex']

    eq_counter += 1
    # 根据编号模式生成label
    if template_numbering == 'sectioned' and current_section:
        section_eq_counter += 1
        label = f'\\label{{eq{current_section}-{section_eq_counter}}}'
    else:
        label = f'\\label{{eq{eq_counter}}}'

    eq_num = re.search(r'\(([\d\-\.]+)\)', text)
    eq_latex = latex
    src_eq_num = ''
    if eq_num:
        # 构建ref_map: 所有模式都替换公式编号为\eqref
        src_eq_num = eq_num.group(1)  # 如 "3" 或 "2-1"
        label = f'\\label{{eq{src_eq_num}}}'
        eq_label_name = f'eq{src_eq_num}'
        ref_map[f'({src_eq_num})'] = eq_label_name
        eq_latex = re.sub(r'\s*\\textbf\{\s*\(' + re.escape(eq_num.group(1)) + r'\)\s*\}', '', eq_latex)
        eq_latex = re.sub(r'\s*\(' + re.escape(eq_num.group(1)) + r'\)\s*$', '', eq_latex)

    # 保留公式后的文本（如 ", i = 1, 2, ... , N"），只删除编号
    eq_latex = eq_latex.replace('$', '')
    # Keep trailing equation text such as ", i = 1, 2, ... , N" while removing
    # converter artifacts that otherwise become invisible glyphs or bad operators.
    eq_latex = _clean_formula_latex(eq_latex)
    # 剥离文本提取器自带的 \begin{equation}...\end{equation} 包裹
    _eq_wrap = re.match(r'^\\begin\{equation\}\s*(.*?)\s*\\end\{equation\}(?:\s*\(.*?\))?\s*$', eq_latex, re.DOTALL)
    if _eq_wrap:
        eq_latex = _eq_wrap.group(1).strip()
    if src_eq_num:
        eq_latex = f'{eq_latex} \\tag{{{src_eq_num}}}'

    # 检测公式是否已包含数学环境（gather/align/multline等，不含equation）
    _math_envs = ('gather', 'align', 'aligned', 'multline', 'split', 'cases', 'array', 'eqnarray')
    _has_math_env = any(f'\\begin{{{e}}}' in eq_latex for e in _math_envs)
    if _has_math_env:
        # 公式已自带环境，不包裹equation，直接输出
        if label:
            eq_latex = _insert_label_before_outer_math_end(eq_latex, label)
        body_lines.append(eq_latex)
    else:
        body_lines.append(f'\\begin{{equation}}')
        body_lines.append(f'  {eq_latex}')
        if label:
            body_lines.append(f'  {label}')
        body_lines.append(f'\\end{{equation}}')

    _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                              body_lines, inserted_img_files, inserted_tbl_ids,
                              layout_spec=layout_spec, ref_map=ref_map,
                              fig_counter=fig_counter, tab_counter=tab_counter,
                              current_section=current_section,
                              template_numbering=template_numbering)

    return eq_counter, section_eq_counter

def _prepare_formula_record(p, ref_map, eq_counter, current_section,
                            section_eq_counter, template_numbering):
    """Prepare one display formula for equation/gather output."""
    text = p['text']
    latex = p['latex']
    eq_counter += 1
    if template_numbering == 'sectioned' and current_section:
        section_eq_counter += 1
        label_name = f'eq{current_section}-{section_eq_counter}'
    else:
        label_name = f'eq{eq_counter}'
    label = f'\\label{{{label_name}}}'

    eq_num = re.search(r'\(([\d\-\.]+)\)', text)
    eq_latex = latex
    src_eq_num = ''
    if eq_num:
        src_eq_num = eq_num.group(1)
        label_name = f'eq{src_eq_num}'
        label = f'\\label{{{label_name}}}'
        ref_map[f'({src_eq_num})'] = label_name
        eq_latex = re.sub(r'\s*\\textbf\{\s*\(' + re.escape(src_eq_num) + r'\)\s*\}', '', eq_latex)
        eq_latex = re.sub(r'\s*\(' + re.escape(src_eq_num) + r'\)\s*$', '', eq_latex)

    eq_latex = eq_latex.replace('$', '')
    eq_latex = _clean_formula_latex(eq_latex)
    eq_wrap = re.match(r'^\\begin\{equation\}\s*(.*?)\s*\\end\{equation\}(?:\s*\(.*?\))?\s*$', eq_latex, re.DOTALL)
    if eq_wrap:
        eq_latex = eq_wrap.group(1).strip()
    if src_eq_num:
        eq_latex = f'{eq_latex} \\tag{{{src_eq_num}}}'
    return {'latex': eq_latex, 'label': label, 'para_index': p['para_index']}, eq_counter, section_eq_counter


def _process_formula_group(formula_paras, body_lines, ref_map, eq_counter,
                           section_eq_counter, current_section,
                           template_numbering, img_insert_map, tbl_insert_map,
                           inserted_img_files, inserted_tbl_ids, layout_spec,
                           fig_counter, tab_counter):
    """Emit consecutive display formulas as one gather environment.

    This avoids stacking separate display skips while leaving the actual spacing
    values to the target LaTeX template/class.
    """
    records = []
    math_envs = ('gather', 'align', 'aligned', 'multline', 'split', 'cases', 'array', 'eqnarray')
    for fp in formula_paras:
        rec, eq_counter, section_eq_counter = _prepare_formula_record(
            fp, ref_map, eq_counter, current_section, section_eq_counter,
            template_numbering)
        records.append(rec)

    if any(any(f'\\begin{{{env}}}' in rec['latex'] for env in math_envs) for rec in records):
        for rec in records:
            body_lines.append('\\begin{equation}')
            body_lines.append(f"  {rec['latex']}")
            body_lines.append(f"  {rec['label']}")
            body_lines.append('\\end{equation}')
    elif len(records) == 1:
        rec = records[0]
        body_lines.append('\\begin{equation}')
        body_lines.append(f"  {rec['latex']}")
        body_lines.append(f"  {rec['label']}")
        body_lines.append('\\end{equation}')
    else:
        body_lines.append('\\begin{gather}')
        for idx, rec in enumerate(records):
            line = f"  {rec['latex']} {rec['label']}"
            if idx < len(records) - 1:
                line += r' \\'
            body_lines.append(line)
        body_lines.append('\\end{gather}')

    for fp in formula_paras:
        _insert_images_and_tables(fp['para_index'], img_insert_map, tbl_insert_map,
                                  body_lines, inserted_img_files, inserted_tbl_ids,
                                  layout_spec=layout_spec, ref_map=ref_map,
                                  fig_counter=fig_counter, tab_counter=tab_counter,
                                  current_section=current_section,
                                  template_numbering=template_numbering)
    return eq_counter, section_eq_counter

def _process_empty_paragraph(p, body_lines, img_by_para, img_insert_map, tbl_insert_map,
                             inserted_img_files, inserted_tbl_ids, layout_spec, ref_map,
                             fig_counter, tab_counter, current_section,
                             template_numbering='simple'):
    """处理空段落

    即使是empty段落，也检查是否有图片/表格需插入
    """
    pi = p['para_index']

    if pi in img_by_para:
        # 这是图片段落，插入图片
        _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                  body_lines, inserted_img_files, inserted_tbl_ids,
                                  layout_spec=layout_spec, ref_map=ref_map,
                                  fig_counter=fig_counter, tab_counter=tab_counter,
                                  current_section=current_section,
                                  template_numbering=template_numbering)
    else:
        # 普通empty段落，检查是否有图片/表格需插入
        _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                  body_lines, inserted_img_files, inserted_tbl_ids,
                                  layout_spec=layout_spec, ref_map=ref_map,
                                  fig_counter=fig_counter, tab_counter=tab_counter,
                                  current_section=current_section,
                                  template_numbering=template_numbering)
    body_lines.append('')


def _process_unknown_paragraph(p, body_lines):
    """处理未知类型段落

    短小的 unknown 段落（<30字）合并到上一段末尾
    """
    text = p['text']
    latex = p['latex']

    if len(text) < 30 and latex and latex.strip():
        for bi in range(len(body_lines) - 1, -1, -1):
            if body_lines[bi].strip() and not body_lines[bi].startswith('\\'):
                body_lines[bi] = body_lines[bi] + latex.strip()
                break
        else:
            body_lines.append(latex)
            body_lines.append('')
    else:
        # 普通段落
        if latex and latex.strip():
            body_lines.append(latex)
            body_lines.append('')
        else:
            body_lines.append('')


def _process_declaration_paragraph(p, body_lines, skeleton_info, _decl_kw):
    """处理声明段落

    Args:
        p: 段落信息字典
        body_lines: 输出行列表
        skeleton_info: 模板骨架信息
        _decl_kw: 声明关键词映射
    """
    text = p['text']
    latex = p['latex']

    decl_match = None
    decl_type = None
    for cmd_name, keywords in _decl_kw.items():
        for kw in keywords:
            if re.match(kw, text, re.IGNORECASE):
                decl_type = cmd_name
                decl_match = True
                break
        if decl_match:
            break

    if decl_type and decl_type in skeleton_info['statement_cmds']:
        cmd = f'\\{decl_type}'
        body_lines.append(f'{cmd}{{{latex}}}')
        body_lines.append('')
        return

    # 致谢走单独逻辑
    if decl_type == 'acknowledgements':
        ack_env = skeleton_info['ack_env'] or 'acknowledgements'
        body_lines.append(f'\\begin{{{ack_env}}}')
        body_lines.append(latex)
        body_lines.append(f'\\end{{{ack_env}}}')
        body_lines.append('')
        return

    # 未匹配的声明，原样输出
    body_lines.append(latex)
    body_lines.append('')


def _process_paragraph_loop(paragraphs, ref_start_para, img_by_para, img_insert_map,
                            tbl_insert_map, layout_spec, ref_map, skeleton_info, spec,
                            _concl_kw, _decl_kw, template_numbering, _has_placeins):
    """段落遍历主循环

    Args:
        paragraphs: 段落列表
        ref_start_para: 参考文献起始位置
        img_by_para: 图片索引映射
        img_insert_map: 图片插入映射
        tbl_insert_map: 表格插入映射
        layout_spec: 排版规格
        ref_map: 编号映射字典
        skeleton_info: 模板骨架信息
        spec: 模板规格
        _concl_kw: 结论关键词集合
        _decl_kw: 声明关键词映射
        template_numbering: 编号模式
        _has_placeins: 是否加载了placeins包

    Returns:
        tuple: (body_lines, inserted_img_files, inserted_tbl_ids, current_section,
                eq_counter, section_eq_counter, used_eq_labels)
    """
    body_lines = []
    inserted_img_files = set()
    inserted_tbl_ids = set()
    used_eq_labels = set()
    eq_counter = 0
    fig_counter = [0]
    tab_counter = [0]
    current_section = 0
    section_eq_counter = 0

    skip_para_indices = set()
    _doc_meta = {}  # 累积正文字号，供caption判断使用
    for idx, p in enumerate(paragraphs):
        pi = p['para_index']
        if pi in skip_para_indices:
            continue
        semantic = p.get('semantic_type', 'unknown')

        # 参考文献之后全部跳过
        if ref_start_para and pi >= ref_start_para:
            continue

        text = p['text']
        latex = p['latex']
        is_heading = p['heading_level'] is not None

        # -- 标题（已单独提取，此处跳过） --
        if semantic == 'title':
            continue

        # -- 作者/机构（已单独提取，此处跳过） --
        if semantic in ('author', 'affiliation', 'front_matter'):
            continue

        # -- 摘要标签（由 \\begin{abstract} 自带，跳过） --
        if semantic == 'abstract_label':
            continue

        # -- 摘要（已单独提取，此处跳过） --
        if semantic == 'abstract':
            continue

        # -- 关键词（已单独提取，此处跳过） --
        if semantic == 'keywords':
            continue

        # -- 图片段落（已通过 _insert_images_and_tables 插入，此处跳过） --
        if pi in img_by_para:
            _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                      body_lines, inserted_img_files, inserted_tbl_ids,
                                      layout_spec=layout_spec, ref_map=ref_map,
                                      fig_counter=fig_counter, tab_counter=tab_counter,
                                      current_section=current_section,
                                      template_numbering=template_numbering)
            continue

        # -- 标题段落 --
        if is_heading:
            current_section, section_eq_counter = _process_heading_paragraph(
                p, body_lines, skeleton_info, spec, _concl_kw,
                current_section, section_eq_counter, _has_placeins,
                img_insert_map, tbl_insert_map, inserted_img_files,
                inserted_tbl_ids, layout_spec, ref_map, fig_counter,
                tab_counter, template_numbering)
            continue

        # -- 独立公式段落 --
        if semantic == 'display_formula':
            formula_group = [p]
            j = idx + 1
            while j < len(paragraphs):
                np = paragraphs[j]
                if ref_start_para and np.get('para_index') >= ref_start_para:
                    break
                if np.get('semantic_type') != 'display_formula':
                    break
                formula_group.append(np)
                skip_para_indices.add(np['para_index'])
                j += 1
            if len(formula_group) > 1:
                eq_counter, section_eq_counter = _process_formula_group(
                    formula_group, body_lines, ref_map, eq_counter,
                    section_eq_counter, current_section, template_numbering,
                    img_insert_map, tbl_insert_map, inserted_img_files,
                    inserted_tbl_ids, layout_spec, fig_counter, tab_counter)
            else:
                eq_counter, section_eq_counter = _process_formula_paragraph(
                    p, body_lines, ref_map, eq_counter, used_eq_labels,
                    current_section, section_eq_counter, template_numbering,
                    img_insert_map, tbl_insert_map, inserted_img_files,
                    inserted_tbl_ids, layout_spec, fig_counter, tab_counter)
            continue

        # -- 图说明/表说明 → 跳过，已在图片/表格插入时通过 caption_full 处理 --
        if semantic in ('figure_caption', 'table_caption'):
            _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                      body_lines, inserted_img_files, inserted_tbl_ids,
                                      layout_spec=layout_spec, ref_map=ref_map,
                                      fig_counter=fig_counter, tab_counter=tab_counter,
                                      current_section=current_section,
                                      template_numbering=template_numbering)
            continue

        if _looks_like_float_caption_text(text, p.get('runs'), _doc_meta):
            _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                      body_lines, inserted_img_files, inserted_tbl_ids,
                                      layout_spec=layout_spec, ref_map=ref_map,
                                      fig_counter=fig_counter, tab_counter=tab_counter,
                                      current_section=current_section,
                                      template_numbering=template_numbering)
            continue

        # -- 致谢 --
        if semantic == 'acknowledgement':
            ack_env = skeleton_info.get('ack_env') or 'acknowledgements'
            body_lines.append(f'\\begin{{{ack_env}}}')
            body_lines.append(latex)
            body_lines.append(f'\\end{{{ack_env}}}')
            body_lines.append('')
            continue

        # -- 声明 --
        if semantic == 'declaration':
            _process_declaration_paragraph(p, body_lines, skeleton_info, _decl_kw)
            continue

        # -- 正文/unknown --
        if semantic == 'empty' or (not text and not p['has_formula']):
            _process_empty_paragraph(p, body_lines, img_by_para, img_insert_map,
                                     tbl_insert_map, inserted_img_files, inserted_tbl_ids,
                                     layout_spec, ref_map, fig_counter, tab_counter,
                                     current_section, template_numbering)
            continue

        # 短小的 unknown 段落（<30字）合并到上一段末尾
        if semantic == 'unknown' and len(text) < 30 and latex and latex.strip():
            for bi in range(len(body_lines) - 1, -1, -1):
                candidate = body_lines[bi].strip()
                if not candidate:
                    continue
                is_structural = (
                    candidate.startswith('\\') or
                    candidate.startswith('{\\') or
                    candidate.startswith('%') or
                    '\\caption{' in candidate or
                    '\\begin{' in candidate or
                    '\\end{' in candidate or
                    '\\draw' in candidate or
                    '\\node' in candidate
                )
                if not is_structural:
                    body_lines[bi] = body_lines[bi] + latex.strip()
                    break
            else:
                body_lines.append(latex)
                body_lines.append('')
            continue

        # 普通段落（非caption、非heading、非formula的正文）
        if latex and latex.strip():
            body_lines.append(latex)
            body_lines.append('')
        else:
            body_lines.append('')
        # 累积正文字号到 _doc_meta，供后续 caption 判断使用
        _runs = p.get('runs', [])
        if _runs and '_body_font_size' not in _doc_meta:
            _sizes = [r['size_pt'] for r in _runs
                      if isinstance(r, dict) and r.get('size_pt') is not None]
            if _sizes:
                from collections import Counter
                _doc_meta['_body_font_size'] = Counter(_sizes).most_common(1)[0][0]

        # 段落输出后，检查是否有图片/表格需在该位置插入
        _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                                  body_lines, inserted_img_files, inserted_tbl_ids,
                                  layout_spec=layout_spec, ref_map=ref_map,
                                  fig_counter=fig_counter, tab_counter=tab_counter,
                                  current_section=current_section,
                                  template_numbering=template_numbering)

    return (body_lines, inserted_img_files, inserted_tbl_ids, current_section,
            eq_counter, section_eq_counter, used_eq_labels, fig_counter, tab_counter)
