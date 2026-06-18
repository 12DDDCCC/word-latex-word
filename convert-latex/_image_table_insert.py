#!/usr/bin/env python3
"""图片和表格插入辅助函数"""

import re
import sys
from pathlib import Path

# shared模块导入
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.caption_utils import (
    strip_caption_prefix as _strip_caption_prefix,
    clean_caption as _clean_caption,
)


def _source_number_value(src_num):
    """Return the visible source number part from a Figure/Table caption prefix."""
    if not src_num:
        return ''
    m = re.search(r'(\d+(?:\.\d+)*)', str(src_num))
    return m.group(1) if m else ''


def _caption_with_source_number(kind, caption, src_num):
    """Let the class format captions while preserving the source visible number."""
    num = _source_number_value(src_num)
    if not num:
        return f'\\caption{{{caption}}}'
    counter = 'figure' if kind == 'figure' else 'table'
    return '{\\renewcommand{\\the' + counter + '}{' + num + '}\\caption{' + caption + '}}'


def _supports_double_column_kind(layout_spec, kind):
    doc_spec = (layout_spec or {}).get('document', {})
    page_spec = (layout_spec or {}).get('page_geometry', {})
    if not doc_spec.get('is_twocolumn'):
        return False
    try:
        column_count = int(page_spec.get('column_count', 2) or 2)
    except (TypeError, ValueError):
        column_count = 2
    if doc_spec.get('is_twocolumn') and column_count < 2:
        column_count = 2
    specific = doc_spec.get(f'supports_double_column_{kind}')
    if specific is None:
        specific = doc_spec.get('supports_double_column_floats')
    return bool(specific) and column_count > 1


def _supports_double_column_floats(layout_spec):
    return _supports_double_column_kind(layout_spec, 'tables')


def _supports_double_column_figures(layout_spec):
    return _supports_double_column_kind(layout_spec, 'figures')


def _column_width_pt(layout_spec):
    doc_spec = (layout_spec or {}).get('document', {})
    page_spec = (layout_spec or {}).get('page_geometry', {})
    try:
        textwidth_mm = float(page_spec.get('textwidth_mm') or 0)
        column_count = int(page_spec.get('column_count', 1) or 1)
        column_sep_mm = float(page_spec.get('column_sep_mm', 0) or 0)
    except (TypeError, ValueError):
        return None
    if doc_spec.get('is_twocolumn') and column_count < 2:
        column_count = 2
    if textwidth_mm <= 0 or column_count <= 1:
        return None
    column_mm = (textwidth_mm - (column_count - 1) * column_sep_mm) / column_count
    return column_mm * 72.27 / 25.4


def image_requires_full_width(img_info, layout_spec=None):
    """Use a starred float only when extracted template columns cannot fit the source image."""
    figure_spec = (layout_spec or {}).get('figure', {})
    if not _supports_double_column_figures(layout_spec):
        return False
    allow_full = figure_spec.get('allow_full_width')
    if allow_full is None:
        allow_full = True
    if not allow_full:
        return False
    explicit_full = (
        img_info.get('is_full_width') or
        img_info.get('full_width') or
        img_info.get('layout_width') == 'textwidth'
    )
    column_pt = _column_width_pt(layout_spec)
    width_pt = img_info.get('width_pt') or img_info.get('display_width_pt')
    if not width_pt and img_info.get('width_twips'):
        width_pt = float(img_info.get('width_twips')) / 20.0
    if not column_pt:
        return bool(explicit_full and not width_pt)
    if not width_pt:
        return True
    try:
        return float(width_pt) > column_pt * 1.02
    except (TypeError, ValueError):
        return bool(explicit_full)


def normal_figure_width(layout_spec=None):
    """Return the template width command for non-starred figures."""
    figure_spec = (layout_spec or {}).get('figure', {})
    doc_spec = (layout_spec or {}).get('document', {})
    page_spec = (layout_spec or {}).get('page_geometry', {})
    try:
        column_count = int(page_spec.get('column_count', 1) or 1)
    except (TypeError, ValueError):
        column_count = 1
    if doc_spec.get('is_twocolumn') or column_count > 1:
        return '\\columnwidth'
    return figure_spec.get('width', '\\textwidth')


def _max_float_body_height_mm(layout_spec, width_command=None, aspect_ratio=None):
    """浮动体高度上限(mm)。

    aspect_ratio 为图片宽/高比。横向图(>1.3)宽度先满,不受 height 约束,
    保持原逻辑;正方形/纵向图(≤1.3)必受 height 约束,当宽度铺满会被
    现有上限截断时(textwidth > height上限),把上限提升到 textwidth 让宽度
    能铺满(接受多 1 页代价,图更完整)。传 None(无图片信息)时走原逻辑,
    保持向后兼容。
    """
    page_spec = (layout_spec or {}).get('page_geometry', {}) if layout_spec else {}
    doc_spec = (layout_spec or {}).get('document', {}) if layout_spec else {}
    policy = (layout_spec or {}).get('float_policy', {}) if layout_spec else {}
    try:
        textheight_mm = float(page_spec.get('textheight_mm') or 0)
    except (TypeError, ValueError):
        textheight_mm = 0
    if not textheight_mm:
        return None

    is_double = width_command == '\\textwidth' and bool(doc_spec.get('is_twocolumn'))
    key = 'dbltopfraction' if is_double else 'topfraction'
    try:
        fraction = float(policy.get(key) or 0)
    except (TypeError, ValueError):
        fraction = 0
    if not 0 < fraction <= 1:
        fraction = 0.82
    height_limit = textheight_mm * fraction - 18.0

    # 正方形/纵向图:宽度铺满会被高度截断时,提升上限让宽度能铺满
    if aspect_ratio is not None and aspect_ratio <= 1.3:
        target_w = _target_graphic_width_mm(width_command, layout_spec)
        if target_w and target_w > height_limit:
            height_limit = target_w
    return max(20.0, height_limit)


def _image_aspect_ratio(img_info):
    """从 img_info 提取宽/高比,无法确定时返回 None。"""
    if not img_info:
        return None
    w = img_info.get('width_pt') or img_info.get('display_width_pt')
    h = img_info.get('height_pt') or img_info.get('display_height_pt')
    if not w and img_info.get('width_twips'):
        w = float(img_info.get('width_twips')) / 20.0
    if not h and img_info.get('height_twips'):
        h = float(img_info.get('height_twips')) / 20.0
    try:
        ratio = float(w) / float(h)
    except (TypeError, ValueError, ZeroDivisionError):
        return None
    return ratio if ratio > 0 else None


def _includegraphics_options(width_command, layout_spec=None, img_info=None):
    opts = [f'width={width_command}']
    aspect_ratio = _image_aspect_ratio(img_info)
    max_height = _max_float_body_height_mm(layout_spec, width_command, aspect_ratio)
    if max_height:
        opts.extend([f'height={max_height:.1f}mm', 'keepaspectratio'])
    return ','.join(opts)


def _target_graphic_width_mm(width_command, layout_spec=None):
    page_spec = (layout_spec or {}).get('page_geometry', {}) if layout_spec else {}
    try:
        textwidth = float(page_spec.get('textwidth_mm') or 0)
        column_count = int(page_spec.get('column_count', 1) or 1)
        column_sep = float(page_spec.get('column_sep_mm', 0) or 0)
    except (TypeError, ValueError):
        return None
    if width_command == '\\textwidth':
        return textwidth or None
    if width_command == '\\columnwidth' and column_count > 1:
        return (textwidth - (column_count - 1) * column_sep) / column_count
    return textwidth or None


def _image_required_space_mm(img_info, layout_spec, width_command):
    target_width = _target_graphic_width_mm(width_command, layout_spec)
    width_pt = img_info.get('width_pt') or img_info.get('display_width_pt')
    height_pt = img_info.get('height_pt') or img_info.get('display_height_pt')
    if not target_width or not width_pt or not height_pt:
        return None
    try:
        image_height = target_width * float(height_pt) / float(width_pt)
    except (TypeError, ValueError, ZeroDivisionError):
        return None
    max_height = _max_float_body_height_mm(
        layout_spec, width_command, _image_aspect_ratio(img_info))
    if max_height:
        image_height = min(image_height, max_height)
    return image_height + 18.0


def _table_required_space_mm(tbl_data, layout_spec):
    """Reserve the caption plus the table's final, possibly scaled height."""
    from supertabular_gen import rendered_table_output_height_mm

    return rendered_table_output_height_mm(tbl_data, layout_spec) + 18.0


def _full_width_container(layout_spec, kind):
    spec = (layout_spec or {}).get(kind, {})
    return spec.get('full_width_container', '')


def _text_bearing_float_position(pos):
    """Keep two-column floats on text pages instead of float-only pages."""
    raw = (pos or '').strip()
    if not raw:
        return 't'
    prefix = '!' if '!' in raw else ''
    allowed = ''.join(ch for ch in raw if ch in 'htb')
    return prefix + (allowed or 't')


def _full_width_position(layout_spec, kind):
    """Place double-column floats at the next page top, keeping text flowing."""
    doc_spec = (layout_spec or {}).get('document', {})
    if doc_spec.get('is_twocolumn') and kind in ('figure', 'table'):
        return '!t'
    spec = (layout_spec or {}).get(kind, {})
    return _text_bearing_float_position(
        spec.get('full_width_float_position') or spec.get('float_position'))


def _append_full_width_block(body_lines, kind, content_lines, layout_spec,
                             required_space_mm=None):
    """Append an indivisible full-text-width float block."""
    if _full_width_container(layout_spec, kind) == 'strip':
        pos = _full_width_position(layout_spec, kind)
        body_lines.append(f'% SKILL-FULLWIDTH-FLOAT {kind} pos={pos}')
        body_lines.append('\\begin{strip}')
        body_lines.append('\\noindent\\begin{minipage}{\\textwidth}')
        body_lines.append('\\begingroup')
        body_lines.append(
            f'\\makeatletter\\def\\@captype{{{kind}}}\\makeatother')
        body_lines.extend(content_lines)
        body_lines.append('\\endgroup')
        body_lines.append('\\end{minipage}')
        body_lines.append('\\end{strip}')
        return

    env = f'{kind}*'
    pos = _full_width_position(layout_spec, kind)
    body_lines.append(f'\\begin{{{env}}}[{pos}]')
    body_lines.append('\\noindent\\begin{minipage}{\\textwidth}')
    body_lines.extend(content_lines)
    body_lines.append('\\end{minipage}')
    body_lines.append(f'\\end{{{env}}}')


def _insert_images_and_tables(pi, img_insert_map, tbl_insert_map,
                              body_lines, inserted_img_files, inserted_tbl_ids,
                              layout_spec=None, ref_map=None, fig_counter=None,
                              tab_counter=None, current_section=0,
                              template_numbering='simple'):
    """在段落 pi 之后插入关联的图片和表格

    使用 para_index 直接映射，不再依赖 context 文本匹配。
    图片/表格的 caption 使用 caption_full（含图例/表例）。
    layout_spec 控制图片/表格的 caption 位置和字体大小。
    双栏模式下使用 figure*/table*（跨栏浮动体）。
    ref_map: 源文档编号→label映射，用于替换正文中的硬编码编号
    fig_counter/tab_counter: list[int,1]，可变计数器
    """
    fig_env = 'figure'
    tbl_env = 'table'
    # 插入图片
    if pi in img_insert_map:
        for img_info in img_insert_map[pi]:
            img_file = img_info.get('image_file', '')
            if img_file in inserted_img_files:
                continue
            inserted_img_files.add(img_file)

            cap = img_info.get('caption_full', '') or img_info.get('caption', '')
            cap_text = img_info.get('caption', '') or img_info.get('caption_full', '')
            src_num = ''
            num_m = re.match(r'(图\s*\d+(?:\.\d+)?|Figure\s*\d+(?:\.\d+)?|Fig\.?\s*\d+(?:\.\d+)?)', cap_text, re.IGNORECASE)
            if num_m:
                src_num = num_m.group(1).strip()
            cap = _clean_caption(_strip_caption_prefix(cap))

            fig_spec = (layout_spec or {}).get('figure', {})
            cap_spec = (layout_spec or {}).get('caption', {})
            fig_cap_pos = fig_spec.get('caption_position', cap_spec.get('figure_position', 'below'))
            fig_float = fig_spec.get('float_position', 'htbp')
            fig_width = normal_figure_width(layout_spec)

            if fig_counter is not None:
                fig_counter[0] += 1
                if template_numbering == 'sectioned' and current_section:
                    fig_label = f'\\label{{fig{current_section}-{fig_counter[0]}}}'
                else:
                    fig_label = f'\\label{{fig{fig_counter[0]}}}'
            else:
                fig_label = ''

            # Let the active LaTeX template/class format the caption.
            cap_cmd = _caption_with_source_number('figure', cap, src_num)

            use_full_fig = image_requires_full_width(img_info, layout_spec)
            fig_env_actual = f'{fig_env}*' if use_full_fig else fig_env
            fig_width_actual = '\\textwidth' if use_full_fig else fig_width

            fig_lines = ['\\centering']
            if fig_cap_pos == 'above':
                fig_lines.append(cap_cmd)
            fig_lines.append(
                f'\\includegraphics[{_includegraphics_options(fig_width_actual, layout_spec, img_info)}]{{fig/{img_file}}}')
            if fig_cap_pos != 'above':
                fig_lines.append(cap_cmd)
            if fig_label:
                fig_lines.append(fig_label)

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

            if ref_map is not None and src_num and fig_label:
                if template_numbering == 'sectioned' and current_section:
                    label_name = f'fig{current_section}-{fig_counter[0]}'
                else:
                    label_name = f'fig{fig_counter[0]}' if fig_counter else 'fig'
                ref_map[src_num] = label_name
                if src_num.startswith('Figure') or src_num.startswith('Fig'):
                    num_part = re.sub(r'^(?:Figure|Fig\.?)\s*', '', src_num)
                    ref_map[f'图{num_part}'] = label_name
                    ref_map[f'图 {num_part}'] = label_name
    # 插入表格
    if pi in tbl_insert_map:
        for tbl_data in tbl_insert_map[pi]:
            tbl_idx = tbl_data.get('table_index', 0)
            if tbl_idx in inserted_tbl_ids:
                continue
            inserted_tbl_ids.add(tbl_idx)
            # 提取源文档中的表编号
            tbl_pos = tbl_data.get('position', {})
            tbl_cap_text = (
                tbl_pos.get('caption', '') or
                tbl_pos.get('table_caption', '') or
                tbl_pos.get('caption_full', '')
            )
            src_tbl_num = ''
            # 同时匹配中文和英文表编号
            tbl_num_m = re.match(r'(表\s*\d+(?:\.\d+)?|Table\s*\d+(?:\.\d+)?)', tbl_cap_text, re.IGNORECASE)
            if tbl_num_m:
                src_tbl_num = tbl_num_m.group(1).strip()

            # 生成表格label（根据编号模式）
            if tab_counter is not None:
                tab_counter[0] += 1
                if template_numbering == 'sectioned' and current_section:
                    tbl_label_name = f'tab{current_section}-{tab_counter[0]}'
                else:
                    tbl_label_name = f'tab{tab_counter[0]}'
                tbl_label = f'\\label{{{tbl_label_name}}}'
            else:
                tbl_label = ''

            # Pass layout_spec to the lossless table renderer, then add the
            # LaTeX table/caption shell here so captions follow the template.
            from tikz_table_gen import process_table, table_requires_full_width
            use_full_table = table_requires_full_width(tbl_data, layout_spec)

            raw_caption = (
                tbl_pos.get('caption_full', '') or
                tbl_pos.get('table_caption', '') or
                tbl_pos.get('caption', '')
            )
            if not raw_caption:
                rows = tbl_data.get('rows', [])
                grid_cols = tbl_data.get('grid_cols', [])
                first_cells = rows[0].get('cells', []) if rows else []
                if first_cells and first_cells[0].get('gridSpan', 1) == len(grid_cols):
                    raw_caption = first_cells[0].get('text', '')
            tbl_caption = _clean_caption(_strip_caption_prefix(raw_caption))

            tbl_spec = (layout_spec or {}).get('table', {})
            cap_spec = (layout_spec or {}).get('caption', {})
            tbl_cap_pos = tbl_spec.get('caption_position', cap_spec.get('table_position', 'above'))
            tbl_float = tbl_spec.get('float_position', 'htbp')
            tbl_align = tbl_spec.get('alignment', 'center')
            if tbl_caption:
                # Let the active LaTeX template/class format the caption.
                cap_cmd = _caption_with_source_number('table', tbl_caption, src_tbl_num)
            else:
                cap_cmd = ''

            table_chunks = [tbl_data]
            for chunk_idx, chunk_data in enumerate(table_chunks):
                tikz_code = process_table(
                    chunk_data, tbl_idx, layout_spec=layout_spec)
                if not tikz_code:
                    continue
                from supertabular_gen import build_supertabular, requires_multipage_table
                if requires_multipage_table(chunk_data, layout_spec) and not use_full_table:
                    body_lines.append(build_supertabular(
                        chunk_data, tikz_code, tbl_caption, tbl_label_name,
                        _source_number_value(src_tbl_num), layout_spec))
                    body_lines.append('')
                    continue
                tbl_env_actual = f'{tbl_env}*' if use_full_table else tbl_env
                tbl_lines = []
                if tbl_align == 'center':
                    tbl_lines.append('\\centering')
                part_cap_cmd = cap_cmd if chunk_idx == 0 else ''
                part_label = tbl_label if chunk_idx == 0 else ''
                if tbl_cap_pos == 'above' and part_cap_cmd:
                    tbl_lines.append(part_cap_cmd)
                    if part_label:
                        tbl_lines.append(part_label)
                tbl_lines.append(tikz_code)
                if tbl_cap_pos != 'above' and part_cap_cmd:
                    tbl_lines.append(part_cap_cmd)
                    if part_label:
                        tbl_lines.append(part_label)
                elif tbl_cap_pos == 'above' and not part_cap_cmd and part_label:
                    tbl_lines.append(part_label)

                if use_full_table:
                    _append_full_width_block(
                        body_lines, 'table', tbl_lines, layout_spec,
                        required_space_mm=_table_required_space_mm(
                            chunk_data, layout_spec))
                else:
                    body_lines.append(f'\\begin{{{tbl_env_actual}}}[{tbl_float}]')
                    body_lines.extend(tbl_lines)
                    body_lines.append(f'\\end{{{tbl_env_actual}}}')
                body_lines.append('')

            # 构建ref_map: 源文档编号→label
            if ref_map is not None and src_tbl_num and tbl_label:
                ref_map[src_tbl_num] = tbl_label_name
                # 同时添加中文key（正文引用可能是中文"表N"而非"Table N"）
                if src_tbl_num.startswith('Table') or src_tbl_num.startswith('Talbe'):
                    num_part = re.sub(r'^(?:Table|Talbe)\s*', '', src_tbl_num, flags=re.IGNORECASE)
                    ref_map[f'表{num_part}'] = tbl_label_name
                    ref_map[f'表 {num_part}'] = tbl_label_name


def build_image_map(image_result):
    """将图片结果按 para_index 建索引

    image_result 是 extract_all_images_with_position 返回的列表，
    每项含 para_index, image_file, caption, context_above, context_below 等。
    返回 dict: para_index → [{image_file, caption, context_above, context_below, ...}, ...]
    """
    img_by_para = {}
    for img in image_result:
        pi = img.get('para_index')
        if pi is not None:
            img_by_para.setdefault(pi, []).append(img)
    return img_by_para
