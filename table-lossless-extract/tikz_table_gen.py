#!/usr/bin/env python3
r"""
JSON -> TikZ LaTeX 精确边框还原生成器
每条边框线独立\draw，实现与Word 100%边框一致
支持合并单元格(gridSpan/vMerge)、粗/细边框线段合并、行高列宽精确还原

输入: all_tables_complete.json 格式
输出: TikZ LaTeX .tex 文件，可xelatex编译为PDF
"""
import json, os, sys, re
from pathlib import Path
from collections import defaultdict

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.caption_utils import clean_caption
from shared.latex_text_utils import escape_latex as esc
from shared.unit_convert import width_to_pt


def bw(borders, direction):
    """获取边框宽度(sz值, 0 if val!=single)"""
    b = borders.get(direction, {})
    if b.get('val') != 'single': return 0
    return int(b.get('sz', '0') or 0)


def calc_actual_rs(rows, ri, col_pos, gs, num_rows):
    """计算vMerge restart cell的实际rowspan"""
    actual_rs = 1
    for nr in range(ri + 1, num_rows):
        ncp = 0
        found = False
        for nc in rows[nr]['cells']:
            ngs = nc.get('gridSpan', 1)
            nvm = nc.get('vMerge', None)
            if ncp == col_pos and nvm == 'continue' and ngs == gs:
                found = True
                break
            ncp += ngs
        if found: actual_rs += 1
        else: break
    return actual_rs


def process_table(tbl_data, table_idx, layout_spec=None):
    """处理单个表格 -> TikZ代码

    layout_spec 控制表格 caption 位置、字体大小等。
    """
    import re  # 确保在函数内可用
    rows = tbl_data['rows']
    num_rows = len(rows)
    num_cols = len(tbl_data['grid_cols'])

    # Keep table text sizing under the active LaTeX template; only structural line style may come from the template spec.
    tbl_spec = (layout_spec or {}).get('table', {})
    rule_style = tbl_spec.get('rule_style', 'default')
    no_vertical_rules = bool(tbl_spec.get('no_vertical_rules')) or rule_style in ('booktabs', 'template_hlines')

    # 列宽 (twips -> cm)
    full_width = table_requires_full_width(tbl_data, layout_spec)
    col_widths_cm = [gc['width_twips'] / 567.0 for gc in tbl_data['grid_cols']]
    col_widths_cm = _normalize_col_widths_cm(col_widths_cm, layout_spec, full_width)
    if full_width:
        col_widths_cm = _rebalance_text_columns(rows, col_widths_cm)

    # 标题行检测：首行gridSpan==num_cols则为内部caption行
    fc = rows[0]['cells'][0]
    gs0 = fc.get('gridSpan', 1)
    data_start = 1 if gs0 == num_cols else 0
    internal_caption = fc.get('text', '') if data_start == 1 else ''
    internal_caption = internal_caption.strip()

    # Caption优先级：内部caption行 > position.table_caption（来自Word字体样式检测）
    # 确保无内部caption行的表格也能正确显示caption
    position_caption = tbl_data.get('position', {}).get('table_caption', '').strip()
    caption = internal_caption or position_caption

    num_data_rows = num_rows - data_start

    # 行高 (twips -> cm)
    row_heights_cm = []
    for ri in range(data_start, num_rows):
        h = rows[ri].get('row_height', None)
        if h is None or (isinstance(h, str) and not h.isdigit()):
            h = 400
        h = int(h)
        if h <= 0: h = 400
        row_heights_cm.append(h / 567.0)

    # 自动检测长文本行并增加行高
    for ri in range(num_rows):
        col_pos = 0
        for ci, cell in enumerate(rows[ri]['cells']):
            gs = cell.get('gridSpan', 1)
            text = cell.get('text', '').strip()
            if text and gs > 0:
                col_w = sum(col_widths_cm[col_pos:col_pos + gs])
                char_units = sum(1.0 if ord(ch) > 127 else 0.55 for ch in text)
                est_lines = max(1, int(char_units * 0.42 / col_w) + 1) if col_w > 0 else 1
                dr = ri - data_start
                if est_lines > 1 and 0 <= dr < len(row_heights_cm):
                    needed_h = est_lines * 0.55
                    if row_heights_cm[dr] < needed_h:
                        row_heights_cm[dr] = needed_h
            col_pos += gs

    # 提取边框线段 + 文字节点
    h_segments = []
    v_segments = []
    cell_nodes = []

    for ri in range(data_start, num_rows):
        col_pos = 0
        for ci, cell in enumerate(rows[ri]['cells']):
            gs = cell.get('gridSpan', 1)
            vm = cell.get('vMerge', None)
            borders = cell.get('borders', {})

            if vm == 'continue':
                col_pos += gs
                continue

            actual_rs = 1
            if vm == 'restart':
                actual_rs = calc_actual_rs(rows, ri, col_pos, gs, num_rows)

            y_top = ri - data_start
            y_bottom = ri - data_start + actual_rs
            x_start = col_pos
            x_end = col_pos + gs

            w = bw(borders, 'top')
            if w > 0: h_segments.append((y_top, x_start, x_end, w))
            w = bw(borders, 'bottom')
            if w > 0: h_segments.append((y_bottom, x_start, x_end, w))
            w = bw(borders, 'right')
            if w > 0: v_segments.append((x_end, y_top, y_bottom, w))
            w = bw(borders, 'left')
            if w > 0: v_segments.append((x_start, y_top, y_bottom, w))

            text = cell.get('text', '').strip()
            is_bold = cell.get('bold', False)
            if not is_bold:
                for p in cell.get('paragraphs', []):
                    for r in p.get('runs', []):
                        fmt = r.get('format', {})
                        if fmt.get('bold'):
                            is_bold = True
                            break

            if text:
                cell_nodes.append({
                    'col_start': col_pos, 'col_end': x_end,
                    'row_start': y_top, 'row_end': y_bottom,
                    'text': text, 'is_bold': is_bold,
                })

            col_pos += gs

    # 合并线段
    h_merged = _merge_h(h_segments)
    v_merged = _merge_v(v_segments)

    # 累积坐标
    x_pos = [0.0]
    for w in col_widths_cm:
        x_pos.append(x_pos[-1] + w)
    y_pos = [0.0]
    for h in row_heights_cm[:num_data_rows]:
        y_pos.append(y_pos[-1] + h)

    # 生成TikZ代码
    lines = []
    lines.append(r'\begin{tikzpicture}')
    if data_start == 1:
        lines.append('  % meta:has_caption_row=1')
        lines.append(f'  % meta:caption_text={internal_caption}')

    # 注入完整元数据注释（逆向解析时优先读取，实现100%无损）
    # 格式: % meta:json=<base64编码的JSON>，包含 x_pos/y_pos/列宽/行高/cell信息
    _meta = _build_meta(tbl_data, data_start, num_rows, x_pos, y_pos,
                        col_widths_cm, row_heights_cm, num_data_rows)
    lines.append(f'  % meta:json={_meta}')
    if full_width:
        lines.append('  % meta:full_width=1')

    if rule_style == 'template_hlines':
        # Keep the target template's no-vertical-rule table style, but do not
        # discard horizontal separators that existed in the Word source table.
        num_h = len(h_merged)
        for idx, (seg_type, y, xs, xe, w) in enumerate(h_merged):
            if y >= len(y_pos) or xs >= len(x_pos) or xe >= len(x_pos):
                continue
            x1, x2, yy = x_pos[xs], x_pos[xe], y_pos[y]
            lw = 1.2 if idx == 0 or idx == num_h - 1 else 0.4
            lines.append(f'  \\draw[line width={lw:.1f}pt] ({x1:.3f},{-yy:.3f}) -- ({x2:.3f},{-yy:.3f});')
    else:
        num_h = len(h_merged)
        for idx, (seg_type, y, xs, xe, w) in enumerate(h_merged):
            lw = width_to_pt(w)
            if lw == 0 or y >= len(y_pos) or xs >= len(x_pos) or xe >= len(x_pos): continue
            x1, x2, yy = x_pos[xs], x_pos[xe], y_pos[y]
            # booktabs ?????/??????????
            if rule_style == 'booktabs' and lw > 0:
                if idx == 0:
                    lw = max(lw, 0.8)  # toprule ?
                elif idx == num_h - 1:
                    lw = max(lw, 0.8)  # bottomrule ?
                else:
                    lw = min(lw, 0.4)  # midrule ?
            lines.append(f'  \\draw[line width={lw:.1f}pt] ({x1:.3f},{-yy:.3f}) -- ({x2:.3f},{-yy:.3f});')

    # Template-driven no-vertical-rule styles do not draw source vertical borders.
    if not no_vertical_rules:
        for seg_type, x, ys, ye, w in v_merged:
            lw = width_to_pt(w)
            if lw == 0 or x >= len(x_pos) or ys >= len(y_pos) or ye >= len(y_pos): continue
            xx, y1, y2 = x_pos[x], y_pos[ys], y_pos[ye]
            lines.append(f'  \\draw[line width={lw:.1f}pt] ({xx:.3f},{-y1:.3f}) -- ({xx:.3f},{-y2:.3f});')

    for node in cell_nodes:
        cs, ce = node['col_start'], node['col_end']
        rs, re = node['row_start'], node['row_end']
        text, is_bold = node['text'], node['is_bold']
        if cs >= len(x_pos) or ce >= len(x_pos) or rs >= len(y_pos) or re >= len(y_pos): continue
        xc = (x_pos[cs] + x_pos[ce]) / 2.0
        yc = -(y_pos[rs] + y_pos[re]) / 2.0
        et = esc(text)
        if is_bold and et: et = r'\textbf{' + et + '}'
        col_w = x_pos[ce] - x_pos[cs]
        row_h = y_pos[re] - y_pos[rs]
        est_lines = max(1, int(len(text) * 0.22 / col_w) + 1) if col_w > 0 else 1
        font_attr = ''
        if est_lines > 1:
            lines.append(f'  \\node[anchor=center, text width={col_w - 0.2:.2f}cm, align=center{font_attr}] at ({xc:.3f},{yc:.3f}) {{{et}}};')
        else:
            lines.append(f'  \\node[anchor=center{font_attr}] at ({xc:.3f},{yc:.3f}) {{{et}}};')

    lines.append(r'\end{tikzpicture}')
    tikz = '\n'.join(lines)
    width_command = _table_width_command(layout_spec, full_width)
    resize_command = _table_resize_command(
        layout_spec, width_command, y_pos[-1] * 10.0 if y_pos else 0.0)
    return f'{resize_command}{{%\n{tikz}\n}}'


def table_requires_full_width(tbl_data, layout_spec=None):
    """Return True only when the template's extracted column width cannot fit the source table."""
    if not _supports_double_column_floats(layout_spec):
        return False
    column_width_cm = _column_width_cm(layout_spec)
    if not column_width_cm:
        return False
    source_width_cm = sum(
        (gc.get('width_twips', 0) or 0) / 567.0
        for gc in tbl_data.get('grid_cols', [])
    )
    return source_width_cm > column_width_cm * 1.02


def _supports_double_column_floats(layout_spec):
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
    support = doc_spec.get('supports_double_column_tables')
    if support is None:
        support = doc_spec.get('supports_double_column_floats')
    return bool(support) and column_count > 1


def _column_width_cm(layout_spec):
    doc_spec = (layout_spec or {}).get('document', {}) if layout_spec else {}
    page_spec = (layout_spec or {}).get('page_geometry', {}) if layout_spec else {}
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
    return (textwidth_mm - (column_count - 1) * column_sep_mm) / column_count / 10.0


def _table_width_command(layout_spec, full_width=False):
    if full_width or not (layout_spec or {}).get('document', {}).get('is_twocolumn'):
        return r'\textwidth'
    return r'\columnwidth'


def _table_resize_command(layout_spec, width_command, rendered_height_mm):
    """Fit table width without shrinking template-derived body text by height."""
    return f'\\resizebox{{{width_command}}}{{!}}'


def _max_table_width_cm(layout_spec, full_width=False):
    table_spec = (layout_spec or {}).get('table', {}) if layout_spec else {}
    page_spec = (layout_spec or {}).get('page_geometry', {}) if layout_spec else {}
    doc_spec = (layout_spec or {}).get('document', {}) if layout_spec else {}
    width_mm = (
        table_spec.get('max_width_mm') or
        table_spec.get('textwidth_mm') or
        page_spec.get('textwidth_mm')
    )
    try:
        width_mm = float(width_mm) if width_mm else None
        column_count = int(page_spec.get('column_count', 1) or 1)
        if doc_spec.get('is_twocolumn') and column_count < 2:
            column_count = 2
        if doc_spec.get('is_twocolumn') and column_count > 1 and width_mm and not full_width:
            column_sep = float(page_spec.get('column_sep_mm', 0) or 0)
            width_mm = (width_mm - (column_count - 1) * column_sep) / column_count
        return width_mm / 10.0 if width_mm else None
    except (TypeError, ValueError):
        return None


def _normalize_col_widths_cm(col_widths_cm, layout_spec, full_width=False):
    max_width_cm = _max_table_width_cm(layout_spec, full_width)
    total_width_cm = sum(col_widths_cm)
    if not max_width_cm or total_width_cm <= 0:
        return col_widths_cm
    scale = max_width_cm / total_width_cm
    return [max(w * scale, 0.2) for w in col_widths_cm]


def _rebalance_text_columns(rows, col_widths_cm):
    """Give long text columns enough width while preserving the extracted total table width."""
    max_text = [0] * len(col_widths_cm)
    for row in rows:
        col_pos = 0
        for cell in row.get('cells', []):
            span = cell.get('gridSpan', 1)
            text = (cell.get('text') or '').strip()
            if span == 1 and col_pos < len(max_text) and not _looks_numeric(text):
                max_text[col_pos] = max(max_text[col_pos], len(text))
            col_pos += span

    targets = []
    for idx, length in enumerate(max_text):
        if length >= 16:
            desired = min(max(col_widths_cm[idx], length * 0.16), 4.5)
            if desired > col_widths_cm[idx]:
                targets.append((idx, desired - col_widths_cm[idx]))
    if not targets:
        return col_widths_cm

    adjusted = list(col_widths_cm)
    total_growth = sum(growth for _, growth in targets)
    donors = [
        idx for idx in range(len(adjusted))
        if idx not in {target for target, _ in targets}
    ]
    available = sum(max(adjusted[idx] - 1.15, 0) for idx in donors)
    if available <= 0:
        return col_widths_cm

    growth_scale = min(1.0, available / total_growth)
    actual_growth = 0.0
    for idx, growth in targets:
        delta = growth * growth_scale
        adjusted[idx] += delta
        actual_growth += delta
    for idx in donors:
        spare = max(adjusted[idx] - 1.15, 0)
        adjusted[idx] -= actual_growth * spare / available
    return adjusted


def _looks_numeric(text):
    if not text:
        return True
    return bool(re.fullmatch(r'[-+±\d.,\s%/()]+', text))


def _build_meta(tbl_data, data_start, num_rows, x_pos, y_pos,
                 col_widths_cm, row_heights_cm, num_data_rows):
    """构建逆向解析所需的元数据，base64编码后嵌入TikZ注释

    包含：x_pos/y_pos精确坐标、列宽/行高、每个cell的完整信息
    逆向解析器优先读取此元数据，跳过所有推断算法，实现100%无损
    """
    import base64

    num_cols = len(tbl_data['grid_cols'])

    # 列宽(twips)
    col_widths = [max(1, int(round(w_cm * 567.0))) for w_cm in col_widths_cm]

    # 行高(twips)
    row_heights = []
    for ri in range(data_start, num_rows):
        h = tbl_data['rows'][ri].get('row_height', None)
        if h is None or (isinstance(h, str) and not h.isdigit()):
            h = 400
        h = int(h)
        if h <= 0:
            h = 400
        row_heights.append(h)

    # cell信息：包含所有cell（含vMerge=continue），保留完整边框信息
    cells = []
    for ri in range(data_start, num_rows):
        col_pos = 0
        for ci, cell in enumerate(tbl_data['rows'][ri]['cells']):
            gs = cell.get('gridSpan', 1)
            vm = cell.get('vMerge', None)
            text = cell.get('text', '').strip()

            # 检测bold
            is_bold = cell.get('bold', False)
            if not is_bold:
                for p in cell.get('paragraphs', []):
                    for r in p.get('runs', []):
                        if r.get('format', {}).get('bold'):
                            is_bold = True
                            break

            # 边框信息（包含所有方向，含nil）
            borders = cell.get('borders', {})
            border_info = {}
            for direction in ('top', 'bottom', 'left', 'right'):
                b = borders.get(direction, {})
                val = b.get('val', '')
                if val == 'single':
                    border_info[direction] = {
                        'val': 'single',
                        'sz': b.get('sz', '4'),
                        'color': b.get('color', '000000'),
                    }
                elif val == 'nil':
                    border_info[direction] = {'val': 'nil', 'sz': '0'}

            # shading / vAlign
            shading = cell.get('shading', {})
            v_align = cell.get('vAlign', '')

            # 段落对齐
            align = ''
            for p in cell.get('paragraphs', []):
                align = p.get('alignment', '')
                if align:
                    break

            cell_info = {
                'r': ri - data_start,
                'c': col_pos,
                'gs': gs,
            }
            if vm is not None:
                cell_info['vm'] = vm
            if text:
                cell_info['t'] = text
            if is_bold:
                cell_info['b'] = True
            if border_info:
                cell_info['bd'] = border_info
            if shading and shading.get('fill'):
                cell_info['sh'] = shading['fill']
            if v_align:
                cell_info['va'] = v_align
            if align:
                cell_info['al'] = align

            cells.append(cell_info)
            col_pos += gs

    meta = {
        'xp': [round(p, 4) for p in x_pos],
        'yp': [round(p, 4) for p in y_pos[:num_data_rows + 1]],
        'cw': col_widths,
        'rh': row_heights,
        'nc': num_cols,
        'ds': data_start,
        'cells': cells,
    }

    json_str = json.dumps(meta, separators=(',', ':'), ensure_ascii=False)
    return base64.b64encode(json_str.encode('utf-8')).decode('ascii')


def _merge_h(segments):
    """合并相邻同宽度的水平线段"""
    if not segments: return []
    grouped = defaultdict(list)
    for y, xs, xe, w in segments:
        grouped[(y, w)].append((xs, xe))
    result = []
    for (y, w), segs in grouped.items():
        s = sorted(segs)
        start, end = s[0]
        for a, b in s[1:]:
            if a <= end: end = max(end, b)
            else: result.append(('h', y, start, end, w)); start, end = a, b
        result.append(('h', y, start, end, w))
    return sorted(result)


def _merge_v(segments):
    """合并相邻同宽度的垂直线段"""
    if not segments: return []
    grouped = defaultdict(list)
    for x, ys, ye, w in segments:
        grouped[(x, w)].append((ys, ye))
    result = []
    for (x, w), segs in grouped.items():
        s = sorted(segs)
        start, end = s[0]
        for a, b in s[1:]:
            if a <= end: end = max(end, b)
            else: result.append(('v', x, start, end, w)); start, end = a, b
        result.append(('v', x, start, end, w))
    return sorted(result)


def generate_tikz_document(json_path, output_dir=None, compile_pdf=True):
    """从JSON生成完整的TikZ LaTeX文档，可选编译PDF"""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    tables = data.get('tables', [])

    if output_dir is None:
        output_dir = os.path.dirname(json_path)

    os.makedirs(output_dir, exist_ok=True)

    import platform as _platform
    _fontset = 'windows' if _platform.system() == 'Windows' else 'auto'

    tikz_parts = []
    for t in tables:
        idx = t.get('table_index', 0)
        tikz_parts.append(process_table(t, idx))

    tex = (r"""\documentclass[a4paper,10pt]{article}
\usepackage[UTF8,""" + f"fontset={_fontset}" + r"""]{ctex}
\usepackage{tikz}
\usetikzlibrary{calc}

\begin{document}
""" + '\n\n'.join(tikz_parts) + r"""

\end{document}
""")

    tex_path = os.path.join(output_dir, 'tikz_tables.tex')
    with open(tex_path, 'w', encoding='utf-8') as f:
        f.write(tex)

    result = {'tex_path': tex_path, 'tables_count': len(tables)}

    if compile_pdf:
        import subprocess
        try:
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'
            proc = subprocess.run(
                ['xelatex', '-interaction=nonstopmode', tex_path],
                capture_output=True, cwd=output_dir, timeout=120, env=env
            )
            pdf_path = os.path.join(output_dir, 'tikz_tables.pdf')
            if os.path.exists(pdf_path):
                result['pdf_path'] = pdf_path
                print(f'PDF compiled: {pdf_path}')
            else:
                result['compile_error'] = 'xelatex compilation error'
                print(f'Compilation failed')
        except FileNotFoundError:
            print('xelatex not found, skipping PDF compilation')
        except subprocess.TimeoutExpired:
            print('xelatex compilation timed out')

    return result


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python tikz_table_gen.py <input.json> [output_dir] [--no-pdf]")
        print("  input.json: all_tables_complete.json 格式的表格数据")
        print("  output_dir: LaTeX/PDF输出目录(默认: 输入文件所在目录)")
        print("  --no-pdf: 只生成.tex不编译PDF")
        sys.exit(1)

    json_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') else None
    compile_pdf = '--no-pdf' not in sys.argv

    result = generate_tikz_document(json_path, output_dir, compile_pdf)
    print(f'Generated {result["tables_count"]} tables -> {result["tex_path"]}')
