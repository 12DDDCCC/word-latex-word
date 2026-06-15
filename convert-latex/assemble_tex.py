#!/usr/bin/env python3
"""核心整合：将所有提取结果合并为完整 .tex 文件

包含：
- assemble_tex: 核心整合函数
- build_image_map: 图片索引构建（re-export from _image_table_insert）
- _strip_heading_number: 标题编号剥离（re-export from _text_helpers）
- _find_reference_start: 参考文献起始检测（re-export from _text_helpers）
- _extract_abstract_keywords: 摘要关键词提取（re-export from _text_helpers）
"""

import re
import sys
from pathlib import Path

# 包内绝对导入（非Python包，不能使用相对导入）
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

# 包内模块导入
from numbering_system import (
    detect_template_numbering_mode, detect_source_numbering_mode,
    REF_KEYWORDS,
)
from skeleton_builder import (
    parse_template_skeleton, clean_preamble, extract_skeleton_commands,
    _build_preamble_from_spec, class_options_from_spec,
)
from template_spec_extract import _extend_keywords_from_spec, APPENDIX_KEYWORDS

# 从子模块导入辅助函数并 re-export
from _text_helpers import (
    _strip_heading_number,
    _find_reference_start,
    _extract_abstract_keywords,
)
from _image_table_insert import (
    _insert_images_and_tables,
    _append_full_width_block,
    _image_required_space_mm,
    _includegraphics_options,
    _table_required_space_mm,
    build_image_map,
    image_requires_full_width,
    normal_figure_width,
)
from _paragraph_process import _process_paragraph_loop

# shared模块导入
from shared.caption_utils import strip_caption_prefix as _strip_caption_prefix


def _derive_table_rule_layout(table_layout, spec, template_content=''):
    """Derive table rule behavior from extracted template signals."""
    table_layout = dict(table_layout or {})
    table_format = spec.get('table_format', {}) if spec else {}
    table_packages = set(table_format.get('table_packages', []) or [])
    hline_commands = set(table_format.get('hline_commands', []) or [])

    if template_content:
        for cmd in (
            'tophline', 'middlehline', 'bottomhline',
            'toprule', 'midrule', 'bottomrule',
            'hline', 'cline', 'cmidrule'
        ):
            if re.search(r'\\' + cmd + r'\b', template_content):
                hline_commands.add(cmd)

    if not table_layout.get('rule_style'):
        if table_format.get('rule_style'):
            table_layout['rule_style'] = table_format['rule_style']
        elif {'tophline', 'middlehline', 'bottomhline'}.issubset(hline_commands):
            table_layout['rule_style'] = 'template_hlines'
        elif 'booktabs' in table_packages or {'toprule', 'midrule', 'bottomrule'}.issubset(hline_commands):
            table_layout['rule_style'] = 'booktabs'

    if hline_commands and not table_layout.get('hline_commands'):
        table_layout['hline_commands'] = sorted(hline_commands)

    if not table_layout.get('vertical_rules'):
        if table_format.get('vertical_rules'):
            table_layout['vertical_rules'] = table_format['vertical_rules']
        elif template_content:
            tabular_specs = re.findall(
                r'\\begin\{(?:tabular|tabularx|longtable)\}\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}',
                template_content
            )
            if tabular_specs:
                table_layout['vertical_rules'] = 'source' if any('|' in colspec for colspec in tabular_specs) else 'none'

    if 'no_vertical_rules' not in table_layout:
        if table_layout.get('vertical_rules') == 'none':
            table_layout['no_vertical_rules'] = True
        elif table_layout.get('vertical_rules') == 'source':
            table_layout['no_vertical_rules'] = False
        elif table_layout.get('rule_style') in ('booktabs', 'template_hlines'):
            table_layout['no_vertical_rules'] = True

    return table_layout


def _template_keeps_inline_identifier_math(spec, layout_spec):
    """Return True only when the template explicitly asks for italic/math identifiers."""
    eq_fmt = spec.get('equation_format', {}) if spec else {}
    eq_layout = layout_spec.get('equation', {}) if layout_spec else {}
    style = (
        eq_fmt.get('inline_identifier_style')
        or eq_layout.get('inline_identifier_style')
        or ''
    ).lower()
    if style in ('math_italic', 'italic', 'math', 'keep_math'):
        return True
    return False


def _simple_identifier_math_to_text(math_content):
    """Convert simple inline identifier math to upright body text, if safe."""
    candidate = re.sub(r'\s+', '', math_content or '')
    if not candidate:
        return None
    if re.search(r'\\(?!_)', candidate):
        return None
    if re.search(r'[\^=+\-*/(),;:\[\]&|<>]', candidate):
        return None

    has_name_marker = bool(re.search(r'(?:_\{\d+\}|_\d|\\_|[A-Z]{2,}\d)', candidate))
    if not has_name_marker:
        return None

    identifier_re = re.compile(
        r'[A-Za-z][A-Za-z0-9]*(?:_\{\d+\}|_\d)?(?:\\_[A-Za-z][A-Za-z0-9]*)*'
    )
    if not identifier_re.fullmatch(candidate):
        return None

    text = re.sub(
        r'([A-Za-z][A-Za-z0-9]*)_\{(\d+)\}',
        r'\1\\textsubscript{\2}',
        candidate
    )
    text = re.sub(
        r'([A-Za-z][A-Za-z0-9]*)_(\d+)',
        r'\1\\textsubscript{\2}',
        text
    )
    text = re.sub(
        r'\b(X?CO|CO)(2)\b',
        r'\1\\textsubscript{\2}',
        text
    )
    return text


def _apply_inline_identifier_style(full_tex, spec, layout_spec):
    """Apply template-derived inline identifier style without touching real formulas."""
    if _template_keeps_inline_identifier_math(spec, layout_spec):
        return full_tex

    begin = full_tex.find('\\begin{document}')
    if begin < 0:
        prefix, body = '', full_tex
    else:
        prefix, body = full_tex[:begin], full_tex[begin:]

    dollar_re = re.compile(r'(?<!\\)\$(?!\$)([^$\n]+?)(?<!\\)\$')

    def repl(match):
        replacement = _simple_identifier_math_to_text(match.group(1))
        return replacement if replacement is not None else match.group(0)

    return prefix + dollar_re.sub(repl, body)


def _first_content_boundary(paragraphs):
    for p in paragraphs:
        text = (p.get('text') or '').strip()
        semantic = p.get('semantic_type')
        if semantic in ('heading', 'abstract', 'abstract_label', 'keywords'):
            return p.get('para_index', 10**9)
        if p.get('heading_level') is not None:
            return p.get('para_index', 10**9)
        if re.match(r'^\s*Abstract\s*[:：.]', text, re.IGNORECASE):
            return p.get('para_index', 10**9)
    return 10**9


def _prepare_metadata_paragraphs(paragraphs, title_para, author_paras):
    """Mark generated DOCX front matter so it is not emitted as body text."""
    boundary = _first_content_boundary(paragraphs)
    candidates = [
        p for p in paragraphs
        if p.get('para_index', 10**9) < boundary and (p.get('text') or '').strip()
    ]
    if not title_para and candidates:
        title_para = candidates[0]
        title_para['semantic_type'] = 'title'

    if not author_paras and title_para:
        for p in candidates:
            if p is title_para:
                continue
            text = (p.get('text') or '').strip()
            if re.match(r'^Correspondence\s*:', text, re.IGNORECASE):
                p['semantic_type'] = 'front_matter'
                continue
            if not re.match(r'^\s*Abstract\s*[:：.]', text, re.IGNORECASE):
                p['semantic_type'] = 'author'
                author_paras = [p]
                break

    for p in candidates:
        text = (p.get('text') or '').strip()
        if re.match(r'^Correspondence\s*:', text, re.IGNORECASE):
            p['semantic_type'] = 'front_matter'
    return title_para, author_paras


def _append_uninserted_images(body_lines, image_result, inserted_img_files, layout_spec):
    fig_spec_d = (layout_spec or {}).get('figure', {})
    fig_float = fig_spec_d.get('float_position', 'htbp')
    fig_width = normal_figure_width(layout_spec)
    # 从模板动态检测双栏模式
    # CLS模板规定: 普通figure使用\columnwidth(半栏), 不用figure*
    fig_env = 'figure'
    for img_info in image_result:
        img_file = img_info.get('image_file', '')
        if not img_file or img_file in inserted_img_files:
            continue
        cap_source = (
            img_info.get('caption_full') or img_info.get('caption') or
            img_info.get('context_below_text') or img_info.get('context_below') or ''
        )
        cap = _strip_caption_prefix(cap_source)
        src_num = ''
        num_m = re.match(r'\s*(?:图|Figure|Fig\.?)\s*(\d+(?:\.\d+)*)', cap_source, re.IGNORECASE)
        if num_m:
            src_num = num_m.group(1)
        use_full_fig = image_requires_full_width(img_info, layout_spec)
        fig_env_actual = f'{fig_env}*' if use_full_fig else fig_env
        fig_width_actual = '\\textwidth' if use_full_fig else fig_width
        fig_lines = [
            '\\centering',
            f'\\includegraphics[{_includegraphics_options(fig_width_actual, layout_spec, img_info)}]{{fig/{img_file}}}',
        ]
        if src_num:
            fig_lines.append(
                '{\\renewcommand{\\thefigure}{' + src_num + '}\\caption{' + cap + '}}')
        else:
            fig_lines.append(f'\\caption{{{cap}}}')
        if use_full_fig:
            _append_full_width_block(
                body_lines, 'figure', fig_lines, layout_spec,
                required_space_mm=_image_required_space_mm(
                    img_info, layout_spec, fig_width_actual))
        else:
            body_lines.append(f'\\begin{{{fig_env_actual}}}[{fig_float}]')
            body_lines.extend(fig_lines)
            body_lines.append(f'\\end{{{fig_env_actual}}}')
        body_lines.append('')
        inserted_img_files.add(img_file)


def _image_caption_source(img_info):
    return (
        img_info.get('caption_full') or img_info.get('caption') or
        img_info.get('context_below_text') or img_info.get('context_below') or ''
    )


def _image_source_number(img_info):
    cap_source = _image_caption_source(img_info)
    match = re.match(r'\s*(?:图|Figure|Fig\.?)\s*(\d+(?:\.\d+)*)', cap_source, re.I)
    return match.group(1) if match else ''


def _fill_missing_image_caption(img_info):
    if img_info.get('caption_full') or img_info.get('caption'):
        return
    cap_source = _image_caption_source(img_info)
    if cap_source:
        img_info['caption'] = cap_source
        img_info['caption_full'] = cap_source


def _is_float_or_empty_semantic(semantic):
    return semantic in ('figure_caption', 'table_caption', 'empty')


def _previous_valid_paragraph(paragraphs, target_pi, para_semantic):
    prev = None
    for p in paragraphs:
        pi = p.get('para_index')
        if pi is None or pi >= target_pi:
            break
        if _is_float_or_empty_semantic(para_semantic.get(pi, 'unknown')):
            continue
        prev = p
    return prev


def _recover_late_image_anchor(img_info, paragraphs, ref_start_para, para_semantic):
    """Recover images that came from a previously bad DOCX after References."""
    num = _image_source_number(img_info)
    if not num:
        return None
    figure_re = re.compile(
        r'(?:图|Figure|Fig\.?)\s*' + re.escape(num) + r'(?:\b|[a-zA-Z])',
        re.I,
    )
    fallback = None
    for p in paragraphs:
        pi = p.get('para_index')
        if pi is None or (ref_start_para and pi >= ref_start_para):
            break
        semantic = para_semantic.get(pi, 'unknown')
        text = p.get('text') or ''
        if _is_float_or_empty_semantic(semantic) or not figure_re.search(text):
            continue
        prev = _previous_valid_paragraph(paragraphs, pi, para_semantic)
        if prev is None:
            fallback = pi
            continue
        if prev.get('heading_level') is not None or para_semantic.get(prev['para_index']) == 'heading':
            return prev['para_index']
        fallback = prev['para_index']
    return fallback


def assemble_tex(text_result, image_result, table_result,
                 template_result, bib_path, output_dir, docx_path, doc_options=None):
    """Merge extracted Word content into a complete template-driven .tex file."""
    from tikz_table_gen import process_table, table_requires_full_width

    skeleton_info = template_result.get('skeleton_info') or {}
    template_content = ''
    if not skeleton_info:
        template_tex_path = template_result['tex_path']
        template_content = Path(template_tex_path).read_text(encoding='utf-8')
        skeleton_info = parse_template_skeleton(template_content)
    if not template_content and template_result.get('tex_path'):
        try:
            template_content = Path(template_result['tex_path']).read_text(encoding='utf-8')
        except Exception:
            template_content = ''

    layout_spec = template_result.get('layout_spec', {})
    spec = template_result.get('spec', {})
    layout_spec = dict(layout_spec or {})
    table_layout = _derive_table_rule_layout(layout_spec.get('table', {}), spec, template_content)
    if table_layout:
        layout_spec['table'] = table_layout
    _intro_kw, _concl_kw, _app_kw, _decl_kw = _extend_keywords_from_spec(spec)

    paragraphs = text_result['paragraphs']
    img_by_para = build_image_map(image_result)

    template_dir_path = Path(template_result.get('output_dir', ''))
    template_numbering = detect_template_numbering_mode(
        template_dir_path, layout_spec, doc_options=doc_options)
    source_numbering = detect_source_numbering_mode(paragraphs)
    numbering_status = 'different' if template_numbering != source_numbering else 'same'
    print(f'[numbering] template={template_numbering}, source={source_numbering}, visible rewrite disabled ({numbering_status})')

    text_para_indices = sorted(p['para_index'] for p in paragraphs)
    para_semantic = {p['para_index']: p.get('semantic_type', 'unknown') for p in paragraphs}
    ref_start_para = _find_reference_start(paragraphs)

    tbl_tables = table_result.get('tables', [])
    tbl_insert_map = {}
    for tbl_data in tbl_tables:
        tbl_pos = tbl_data.get('position', {})
        tbl_pi = tbl_pos.get('paragraph_index')
        if tbl_pi is None:
            tbl_insert_map.setdefault(-1, []).append(tbl_data)
            continue
        prev_pi = None
        for tpi in text_para_indices:
            if tpi >= tbl_pi:
                break
            semantic = para_semantic.get(tpi, 'unknown')
            if semantic in ('figure_caption', 'table_caption', 'empty'):
                continue
            prev_pi = tpi
        if prev_pi is not None:
            tbl_insert_map.setdefault(prev_pi, []).append(tbl_data)
        else:
            tbl_insert_map.setdefault(-1, []).append(tbl_data)

    img_insert_map = {}
    for img_pi, img_items in img_by_para.items():
        for img_info in img_items:
            if ref_start_para and img_pi >= ref_start_para:
                recovered_pi = _recover_late_image_anchor(
                    img_info, paragraphs, ref_start_para, para_semantic)
                if recovered_pi is not None:
                    _fill_missing_image_caption(img_info)
                    img_insert_map.setdefault(recovered_pi, []).append(img_info)
                    continue
            prev_pi = None
            for tpi in text_para_indices:
                if tpi >= img_pi:
                    break
                semantic = para_semantic.get(tpi, 'unknown')
                if semantic in ('figure_caption', 'table_caption', 'empty'):
                    continue
                prev_pi = tpi
            if prev_pi is not None:
                img_insert_map.setdefault(prev_pi, []).append(img_info)
            else:
                img_insert_map.setdefault(-1, []).append(img_info)

    if -1 in tbl_insert_map:
        print(f"WARNING: {len(tbl_insert_map[-1])} tables were not mapped to a body paragraph; appending at the end")

    if spec:
        preamble = _build_preamble_from_spec(spec, template_result['journal'])
        config_mode = next(
            (opt for opt in (doc_options or [])
             if opt in ('final', 'manuscript', 'discussions')),
            None,
        )
        cls_options = class_options_from_spec(
            spec, template_result['journal'], config_mode)
        option_text = ','.join(cls_options)
        replacement = (
            rf'\\documentclass[{option_text}]'
            if option_text else r'\\documentclass'
        )
        preamble = re.sub(
            r'\\documentclass(?:\[[^\]]*\])?', replacement, preamble, count=1)
    else:
        doc_begin_pos = template_content.find('\\begin{document}')
        preamble = template_content[:doc_begin_pos] if doc_begin_pos >= 0 else ''
    preamble = clean_preamble(preamble)

    # 如果 doc_options 指定了非 manuscript 模式，替换 documentclass 选项
    if doc_options:
        _mode = None
        for _opt in doc_options:
            if _opt in ('final', 'discussions'):
                _mode = _opt
                break
        if _mode:
            preamble = re.sub(
                r'(\\documentclass\[)([^\]]*?)\bmanuscript\b',
                rf'\1\2{_mode}',
                preamble
            )

    import platform as _platform
    if 'ctex' not in preamble:
        _fontset = 'windows' if _platform.system() == 'Windows' else 'auto'
        ctex_line = f'\\usepackage[UTF8,fontset={_fontset}]{{ctex}}\n'
        docclass_pos = preamble.find('\\documentclass')
        docclass_line_end = preamble.find('\n', docclass_pos)
        if docclass_pos >= 0 and docclass_line_end >= 0:
            preamble = preamble[:docclass_line_end+1] + ctex_line + preamble[docclass_line_end+1:]
        else:
            preamble = ctex_line + preamble

    doc_spec = layout_spec.get('document', {}) if layout_spec else {}
    if doc_spec.get('needs_fontspec') and 'fontspec' not in preamble:
        fontspec_lines = ['\\usepackage{fontspec}']
        if doc_spec.get('main_font'):
            fontspec_lines.append(f'\\setmainfont{{{doc_spec["main_font"]}}}')
        if doc_spec.get('sans_font'):
            fontspec_lines.append(f'\\setsansfont{{{doc_spec["sans_font"]}}}')
        if doc_spec.get('mono_font'):
            fontspec_lines.append(f'\\setmonofont{{{doc_spec["mono_font"]}}}')
        preamble = preamble.rstrip() + '\n' + '\n'.join(fontspec_lines) + '\n'
    # Math font is controlled by the target template/class.

    spec_pkgs = set(spec.get('required_packages', {}).keys()) if spec else set()
    cls_builtin_pkgs = list(spec_pkgs)
    preamble_lines = preamble.split('\n')
    _pkg_re = re.compile(r'\\usepackage\s*(?:\[.*?\])?\{([^}]+)\}')
    filtered = []
    for line in preamble_lines:
        m = _pkg_re.search(line)
        if m:
            pkgs_in_line = [p.strip() for p in m.group(1).split(',')]
            remaining = [p for p in pkgs_in_line if p not in cls_builtin_pkgs]
            if not remaining:
                continue
            if remaining != pkgs_in_line:
                line = line.replace(m.group(1), ', '.join(remaining))
        filtered.append(line)
    preamble = '\n'.join(filtered)

    preamble = _apply_spec_to_preamble(preamble, spec, layout_spec, spec_pkgs)
    preamble = _ensure_url_breaking(preamble)
    preamble = _ensure_cjk_min_parindent(preamble, paragraphs)
    _has_placeins = 'placeins' in preamble

    title_para = next((p for p in paragraphs if p.get('semantic_type') == 'title'), None)
    author_paras = [p for p in paragraphs if p.get('semantic_type') == 'author']
    title_para, author_paras = _prepare_metadata_paragraphs(
        paragraphs, title_para, author_paras)
    affil_paras = [p for p in paragraphs if p.get('semantic_type') == 'affiliation']
    abstract_lines, keywords_lines = _extract_abstract_keywords(
        paragraphs, ref_start_para, skeleton_info, layout_spec)

    ref_map = {}
    (body_lines, inserted_img_files, inserted_tbl_ids, current_section,
     eq_counter, section_eq_counter, used_eq_labels,
     fig_counter, tab_counter) = _process_paragraph_loop(
        paragraphs, ref_start_para, img_by_para, img_insert_map,
        tbl_insert_map, layout_spec, ref_map, skeleton_info, spec,
        _concl_kw, _decl_kw, template_numbering, _has_placeins)

    # 追加模板要求的 statement 命令占位符（仅当 body_lines 中完全没有时）
    # 去重：同时检查命令标签和关键词文本（避免声明段落已被作为正文输出时再追加）
    _decl_title_map = {
        'codeavailability': ('code availability', '代码可用性'),
        'dataavailability': ('data availability', '数据可用性'),
        'sampleavailability': ('sample availability', '样品可用性'),
        'competinginterests': ('competing interest', '利益冲突'),
        'authorcontribution': ('author contribution', '作者贡献'),
    }
    for cmd_name, placeholder in skeleton_info.get('statement_cmds', {}).items():
        cmd_tag = f'\\{cmd_name}'
        # 检查命令标签是否已存在
        if any(cmd_tag in bl for bl in body_lines):
            continue
        # 检查对应关键词文本是否已存在于某行中（声明段落可能已被作为正文输出）
        keywords = _decl_title_map.get(cmd_name, ())
        if keywords:
            title_found = any(
                kw in bl.lower()
                for kw in keywords
                for bl in body_lines
                if bl and not bl.strip().startswith('%')
            )
            if title_found:
                continue
        statement_content = placeholder
        if cmd_name == 'acknowledgements' and not str(placeholder).strip():
            statement_content = '\\mbox{}'
        body_lines.append(f'{cmd_tag}{{{statement_content}}}')
        body_lines.append('')

    if -1 in tbl_insert_map:
        for tbl_data in tbl_insert_map[-1]:
            tbl_id = tbl_data.get('table_index', -1)
            if tbl_id in inserted_tbl_ids:
                continue
            tikz_code = process_table(tbl_data, tbl_id, layout_spec=layout_spec)
            if tikz_code:
                tbl_spec = (layout_spec or {}).get('table', {})
                tbl_float = tbl_spec.get('float_position', 'htbp')
                use_full_table = table_requires_full_width(tbl_data, layout_spec)
                tbl_env = 'table*' if use_full_table else 'table'
                tbl_lines = []
                if tbl_spec.get('alignment', 'center') == 'center':
                    tbl_lines.append('\\centering')
                tbl_lines.append(tikz_code)
                if use_full_table:
                    _append_full_width_block(
                        body_lines, 'table', tbl_lines, layout_spec,
                        required_space_mm=_table_required_space_mm(
                            tbl_data, layout_spec))
                else:
                    body_lines.append(f'\\begin{{{tbl_env}}}[{tbl_float}]')
                    body_lines.extend(tbl_lines)
                    body_lines.append(f'\\end{{{tbl_env}}}')
                body_lines.append('')
            inserted_tbl_ids.add(tbl_id)
    if -1 in img_insert_map:
        # CLS模板规定: 普通figure使用\columnwidth(半栏), 不用figure*
        fig_env2 = 'figure'
        for img_info in img_insert_map[-1]:
            img_file = img_info.get('image_file', '')
            if img_file in inserted_img_files:
                continue
            cap = img_info.get('caption_full', '') or img_info.get('caption', '')
            cap = _strip_caption_prefix(cap)
            fig_spec_d = (layout_spec or {}).get('figure', {})
            fig_float = fig_spec_d.get('float_position', 'htbp')
            fig_width = normal_figure_width(layout_spec)
            use_full_fig = image_requires_full_width(img_info, layout_spec)
            fig_env_actual = f'{fig_env2}*' if use_full_fig else fig_env2
            fig_width_actual = '\\textwidth' if use_full_fig else fig_width
            fig_lines = [
                '\\centering',
                f'\\includegraphics[{_includegraphics_options(fig_width_actual, layout_spec, img_info)}]{{fig/{img_file}}}',
                f'\\caption{{{cap}}}',
            ]
            if use_full_fig:
                _append_full_width_block(
                    body_lines, 'figure', fig_lines, layout_spec,
                    required_space_mm=_image_required_space_mm(
                        img_info, layout_spec, fig_width_actual))
            else:
                body_lines.append(f'\\begin{{{fig_env_actual}}}[{fig_float}]')
                body_lines.extend(fig_lines)
                body_lines.append(f'\\end{{{fig_env_actual}}}')
            body_lines.append('')
            inserted_img_files.add(img_file)

    all_tbl_ids = set(tbl_data.get('table_index', -1) for tbl_data in tbl_tables)
    missing_tbls = all_tbl_ids - inserted_tbl_ids
    if missing_tbls:
        print(f'WARNING: Tables not inserted: {missing_tbls}')
    _append_uninserted_images(body_lines, image_result, inserted_img_files, layout_spec)
    all_img_files = set(img_info.get('image_file', '') for img_info in image_result)
    missing_imgs = all_img_files - inserted_img_files
    if missing_imgs:
        print(f'WARNING: Images not inserted: {missing_imgs}')

    bib_spec_d = (layout_spec or {}).get('bibliography', {}) if layout_spec else {}
    bib_style = bib_spec_d.get('bst_file', '') or skeleton_info.get('bib_style', '') or template_result.get('journal', '')
    bib_filename = bib_spec_d.get('bib_filename', '') or skeleton_info.get('bib_filename', 'references')
    if body_lines and body_lines[-1]:
        body_lines.append('')
    body_lines.append('\\par')
    body_lines.append(f'\\bibliographystyle{{{bib_style}}}')
    body_lines.append(f'\\bibliography{{{bib_filename}}}')
    body_lines.append('')

    full_tex = preamble + '\\begin{document}\n\n'
    full_tex = _assemble_metadata(full_tex, template_result, template_content,
                                  paragraphs, title_para, author_paras, affil_paras,
                                  skeleton_info, spec, layout_spec)
    full_tex = _insert_abstract_keywords(full_tex, abstract_lines, keywords_lines,
                                         skeleton_info, layout_spec, paragraphs)
    full_tex += '\n'.join(body_lines) + '\n\n\\end{document}\n'
    if spec.get('bibliography_format', {}).get('citation_command') == 'cite':
        full_tex = re.sub(r'\\cite[pt]\{', r'\\cite{', full_tex)

    from tex_postprocess import postprocess_tex
    full_tex = _apply_inline_identifier_style(full_tex, spec, layout_spec)
    full_tex = postprocess_tex(full_tex, layout_spec=layout_spec)

    tex_path = Path(output_dir) / f'{template_result["journal"]}_full.tex'
    with open(tex_path, 'w', encoding='utf-8') as f:
        f.write(full_tex)

    return tex_path, skeleton_info


def _ensure_url_breaking(preamble):
    """Enable URL line breaks without changing the bibliography style."""
    if 'SKILL-URL-BREAKING' in preamble:
        return preamble
    block = '\n'.join([
        '% SKILL-URL-BREAKING',
        '\\PassOptionsToPackage{hyphens}{url}',
        '\\IfFileExists{xurl.sty}{\\usepackage{xurl}}{}',
        '\\makeatletter',
        '\\AtBeginDocument{%',
        '  \\Urlmuskip=0mu plus 1mu%',
        '  \\g@addto@macro\\UrlBreaks{\\do\\/\\do-\\do\\.\\do\\_\\do\\?\\do\\&\\do\\=\\do\\#}%',
        '}',
        '\\makeatother',
    ])
    docclass = re.search(r'\\documentclass(?:\[[^\]]*\])?\{[^}]+\}\s*', preamble)
    if docclass:
        return preamble[:docclass.end()] + block + '\n' + preamble[docclass.end():]
    return block + '\n' + preamble


_CJK_RE = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff]')
_NON_BODY_SEMANTICS = {
    'title', 'author', 'affiliation', 'caption', 'figure_caption',
    'table_caption', 'reference', 'references', 'bibliography',
}


def _has_cjk_body_text(paragraphs):
    for para in paragraphs or []:
        semantic = (para.get('semantic_type') or '').lower()
        if semantic in _NON_BODY_SEMANTICS:
            continue
        if _CJK_RE.search(para.get('text') or ''):
            return True
    return False


def _ensure_cjk_min_parindent(preamble, paragraphs):
    """Use at least two CJK characters for Chinese body text."""
    if 'SKILL-CJK-PARINDENT' in preamble or not _has_cjk_body_text(paragraphs):
        return preamble
    block = '\n'.join([
        '% SKILL-CJK-PARINDENT',
        r'\IfFileExists{indentfirst.sty}{\usepackage{indentfirst}}{}',
        r'\makeatletter',
        r'\AtBeginDocument{\ifdim\parindent<2em\setlength{\parindent}{2em}\fi\let\@afterindentfalse\@afterindenttrue}',
        r'\makeatother',
    ])
    docclass = re.search(r'\\documentclass(?:\[[^\]]*\])?\{[^}]+\}\s*', preamble)
    if docclass:
        return preamble[:docclass.end()] + block + '\n' + preamble[docclass.end():]
    return block + '\n' + preamble


def _uses_strip_full_width(layout_spec):
    if not (layout_spec or {}).get('document', {}).get('is_twocolumn'):
        return False
    for kind in ('figure', 'table'):
        if (layout_spec or {}).get(kind, {}).get('full_width_container') == 'strip':
            return True
    return False


def _uses_native_star_full_width(layout_spec):
    if not (layout_spec or {}).get('document', {}).get('is_twocolumn'):
        return False
    for kind in ('figure', 'table'):
        if (layout_spec or {}).get(kind, {}).get('full_width_container') == 'native-star':
            return True
    return False


def _apply_spec_to_preamble(preamble, spec, layout_spec, spec_pkgs):
    """Add only compile-enabling commands; formatting stays in the target template/class."""
    eq_fmt = spec.get('equation_format', {}) if spec else {}
    amsmath_opts = eq_fmt.get('amsmath_options', '')
    if 'amsmath' not in preamble and 'amsmath' not in spec_pkgs:
        option_block = (
            f'[{amsmath_opts}]'
            if amsmath_opts and '\\' not in amsmath_opts and '@' not in amsmath_opts
            else ''
        )
        preamble += f'\\usepackage{option_block}{{amsmath}}\n'

    fig_fmt = spec.get('figure_format', {}) if spec else {}
    if fig_fmt.get('graphics_path') and '\\graphicspath' not in preamble:
        preamble += f'\\graphicspath{{{fig_fmt["graphics_path"]}}}\n'
    if fig_fmt.get('graphics_extensions') and '\\DeclareGraphicsExtensions' not in preamble:
        preamble += f'\\DeclareGraphicsExtensions{{{fig_fmt["graphics_extensions"]}}}\n'

    bib_fmt = spec.get('bibliography_format', {}) if spec else {}
    natbib_opts = bib_fmt.get('natbib_options', '')
    if (natbib_opts and '\\' not in natbib_opts and '@' not in natbib_opts
            and 'natbib' not in preamble and 'natbib' not in spec_pkgs):
        preamble += f'\\usepackage[{natbib_opts}]{{natbib}}\n'
    bibpunct = bib_fmt.get('bibpunct', {})
    if bibpunct and '\\bibpunct' not in preamble:
        bp_args = [bibpunct.get(key, '') for key in ('open', 'close', 'style', 'sep', 'between', 'after')]
        preamble += '\\bibpunct{' + '}{'.join(bp_args) + '}\n'

    tbl_fmt = spec.get('table_format', {}) if spec else {}
    extra_pkgs = []
    if 'graphicx' not in preamble and 'graphicx' not in spec_pkgs:
        extra_pkgs.append('\\usepackage{graphicx}')
    if 'url' not in preamble and 'url' not in spec_pkgs:
        extra_pkgs.append('\\usepackage{url}')
    if 'tikz' not in preamble and 'tikz' not in spec_pkgs:
        extra_pkgs.append('\\usepackage{tikz}')
        extra_pkgs.append('\\usetikzlibrary{calc}')
    if 'soul' not in preamble and 'soul' not in spec_pkgs:
        extra_pkgs.append('\\usepackage{soul}')
    if 'placeins' not in preamble and 'placeins' not in spec_pkgs:
        extra_pkgs.append('\\usepackage{placeins}')
    if _uses_strip_full_width(layout_spec):
        if 'cuted' not in preamble and 'cuted' not in spec_pkgs:
            extra_pkgs.append(
                '\\IfFileExists{cuted.sty}{\\usepackage{cuted}}{}')
        extra_pkgs.append(
            r'\IfFileExists{needspace.sty}{\usepackage{needspace}'
            r'\newcommand{\skillneedspace}[1]{\Needspace{#1}}}'
            r'{\newcommand{\skillneedspace}[1]{\par}}')
    if (_uses_native_star_full_width(layout_spec)
            and 'stfloats' not in preamble and 'stfloats' not in spec_pkgs
            and 'dblfloatfix' not in preamble and 'dblfloatfix' not in spec_pkgs):
        extra_pkgs.append(
            '\\IfFileExists{stfloats.sty}{\\usepackage{stfloats}}'
            '{\\IfFileExists{dblfloatfix.sty}{\\usepackage{dblfloatfix}}{}}')
    extra_pkgs.append(
        '\\providecommand{\\chem}[1]{\\ensuremath{\\mathrm{#1}}}')
    if extra_pkgs:
        preamble = preamble.rstrip() + '\n' + '\n'.join(extra_pkgs) + '\n'

    return preamble

def _assemble_metadata(full_tex, template_result, template_content,
                       paragraphs, title_para, author_paras, affil_paras,
                       skeleton_info, spec, layout_spec):
    """组装元数据命令区

    处理 metadata_block、title/author/affil 占位符替换
    """
    # 提取元数据命令区
    if 'metadata_block' in template_result:
        skeleton_cmds = template_result['metadata_block']
        if skeleton_cmds.strip():
            full_tex += skeleton_cmds + '\n'
        _title_para = title_para
        _author_paras = author_paras
        if not _title_para:
            for p in paragraphs:
                if p['text'].strip() and p.get('semantic_type') not in ('abstract', 'abstract_label', 'keywords', 'empty'):
                    _title_para = p
                    break
        if not _author_paras and _title_para:
            for p in paragraphs:
                if p['para_index'] > _title_para['para_index'] and p['text'].strip():
                    if p.get('semantic_type') not in ('abstract', 'abstract_label', 'keywords', 'heading'):
                        _author_paras = [p]
                        break
        if _title_para or _author_paras:
            meta_lines = []
            if _title_para:
                title_text = _title_para.get('latex', _title_para['text'])
                title_fmt = spec.get('title_format', {}) if spec else {}
                if title_fmt.get('title_args') == 'short_full':
                    short_title = re.sub(r'\$[^$]+\$', '', _title_para['text'])[:50].rstrip()
                    meta_lines.append(f'\\title[{short_title}]{{{title_text}}}')
                else:
                    meta_lines.append(f'\\title{{{title_text}}}')
                meta_lines.append('')
            if _author_paras:
                author_text = ' '.join(p.get('latex', p['text']) for p in _author_paras)
                parts = author_text.split()
                given, surname = '', ''
                if len(parts) == 1:
                    given, surname = '', parts[0]
                elif len(parts) >= 2:
                    given, surname = ' '.join(parts[:-1]), parts[-1]
                author_fmt = (template_result.get('spec', {}) or {}).get('author_format', {})
                if author_fmt.get('author_args') == 2:
                    meta_lines.append(f'\\Author[][EMAIL]{{{given}}}{{{surname}}}')
                else:
                    meta_lines.append(f'\\author{{{author_text}}}')
                meta_lines.append('')
            if affil_paras:
                affil_text = ' '.join(p['latex'] for p in affil_paras)
                meta_lines.append(f'\\affil[]{{{affil_text}}}')
                meta_lines.append('')
            tmpl = (template_result.get('spec', {}) or {}).get('template_specific', {})
            if _title_para and tmpl.get('runningtitle', {}).get('exists'):
                running_title = _title_para['text'].strip()
                if len(running_title) > 60:
                    running_title = running_title[:57] + '...'
                meta_lines.append(f'\\runningtitle{{{running_title}}}')
            if _author_paras and tmpl.get('runningauthor', {}).get('exists'):
                running_author = _author_paras[0]['text'].strip()
                meta_lines.append(f'\\runningauthor{{{running_author}}}')
            if tmpl.get('correspondence', {}).get('exists'):
                meta_lines.append('\\correspondence{[EMAIL PLACEHOLDER]}')
            meta_lines.append('')
            full_tex = full_tex.replace('\\maketitle', '\n'.join(meta_lines) + '\\maketitle')
    elif template_content:
        doc_begin_pos = template_content.find('\\begin{document}')
        doc_end_pos = template_content.find('\\end{document}')
        if doc_begin_pos >= 0 and doc_end_pos >= 0:
            skeleton_body = template_content[doc_begin_pos + len('\\begin{document}'):doc_end_pos]
            skeleton_cmds = extract_skeleton_commands(skeleton_body)
            if skeleton_cmds.strip():
                full_tex += skeleton_cmds + '\n'

    # 替换模板占位符为实际内容
    if title_para:
        full_tex = full_tex.replace('\\title{TEXT}', f"\\title{{{title_para['latex']}}}")
        full_tex = full_tex.replace('\\title{TITLE}', f"\\title{{{title_para['latex']}}}")

    if author_paras:
        author_names = ' '.join(p['latex'] for p in author_paras)
        parts = author_names.split()
        if len(parts) == 1:
            given, surname = '', parts[0]
        else:
            given, surname = ' '.join(parts[:-1]), parts[-1]
        full_tex = full_tex.replace('\\Author[][EMAIL]{}{}', f"\\Author[][EMAIL]{{{given}}}{{{surname}}}")
        full_tex = full_tex.replace('\\Author[]{given_name}{surname}', f"\\Author[]{{{given}}}{{{surname}}}")
        full_tex = full_tex.replace('\\Author[][EMAIL]{given_name}{surname}', f"\\Author[][EMAIL]{{{given}}}{{{surname}}}")
        full_tex = re.sub(
            r'\\author\{[^}]*\}',
            lambda _match: f'\\author{{{author_names}}}',
            full_tex,
        )
        full_tex = re.sub(r'\\Author\[\]\{\}\{\}\s*$', '', full_tex, flags=re.MULTILINE)

    full_tex = re.sub(r'\\Author\[\]\[EMAIL\]\{given_name\}\{surname\}\s*%%.*$', '', full_tex, flags=re.MULTILINE)

    if affil_paras:
        affil_text = ' '.join(p['latex'] for p in affil_paras)
        full_tex = full_tex.replace('\\affil[]{ADDRESS}', f"\\affil[]{{{affil_text}}}")

    full_tex = full_tex.replace('\\affil[]{ADDRESS}', "\\affil[]{[AFFILIATION PLACEHOLDER]}")
    full_tex = full_tex.replace('\\affil[]{}', "\\affil[]{[AFFILIATION PLACEHOLDER]}")

    if title_para:
        running_title = title_para['text'].strip()
        if len(running_title) > 60:
            running_title = running_title[:57] + '...'
        full_tex = full_tex.replace('\\runningtitle{SHORT TITLE}', f'\\runningtitle{{{running_title}}}')
        full_tex = full_tex.replace('\\runningtitle{}', f'\\runningtitle{{{running_title}}}')

    if author_paras:
        running_author = author_paras[0]['text'].strip()
        full_tex = full_tex.replace('\\runningauthor{SHORT AUTHOR}', f'\\runningauthor{{{running_author}}}')
        full_tex = full_tex.replace('\\runningauthor{}', f'\\runningauthor{{{running_author}}}')
    else:
        full_tex = full_tex.replace('\\runningauthor{SHORT AUTHOR}', '\\runningauthor{}')

    full_tex = full_tex.replace('\\correspondence{EMAIL}', "\\correspondence{[EMAIL PLACEHOLDER]}")
    full_tex = full_tex.replace('\\correspondence{}', "\\correspondence{[EMAIL PLACEHOLDER]}")

    return full_tex


def _insert_abstract_keywords(full_tex, abstract_lines, keywords_lines,
                              skeleton_info, layout_spec, paragraphs):
    """插入摘要和关键词块

    根据 skeleton_info 中的 abstract_after_maketitle 决定插入位置，
    根据 layout_spec 中的 keywords_inside_abstract 决定关键词是否在摘要内
    """
    # 如果存在实际摘要，移除模板骨架中的空 abstract 环境
    has_abstract = any(p.get('semantic_type') == 'abstract' for p in paragraphs)
    if has_abstract:
        abs_env_re = skeleton_info.get('abstract_env', 'abstract')
        empty_abs_pattern = r'\\begin\{' + abs_env_re + r'\}\s*\\end\{' + abs_env_re + r'\}'
        full_tex = re.sub(empty_abs_pattern, '', full_tex)

    abstract_block = '\n'.join(abstract_lines) if abstract_lines else ''
    keywords_block = '\n'.join(keywords_lines) if keywords_lines else ''

    abs_spec = layout_spec.get('abstract', {}) if layout_spec else {}
    keywords_inside = abs_spec.get('keywords_inside', False)

    if abstract_block or keywords_block:
        if keywords_inside and abstract_block and keywords_block:
            abs_kw_block = '\n' + abstract_block + '\n' + keywords_block + '\n'
        else:
            abs_kw_block = ''
            if abstract_block:
                abs_kw_block += '\n' + abstract_block + '\n\n'
            if keywords_block:
                abs_kw_block += keywords_block + '\n\n'

        if skeleton_info['abstract_after_maketitle']:
            full_tex = full_tex.replace('\\maketitle', '\\maketitle' + abs_kw_block)
        else:
            full_tex = full_tex.replace('\\maketitle', abs_kw_block + '\\maketitle')

    return full_tex
