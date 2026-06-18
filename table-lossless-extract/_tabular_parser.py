#!/usr/bin/env python3
"""tabular 表格解析模块

将 LaTeX tabular 表格解析为 JSON 格式，实现部分无损转换。
只有线型规则，无逐线宽度值。
"""
import re
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import match_balanced_braces as _match_balanced_braces, to_subscript as _to_subscript
from shared.unit_convert import cm_to_twips as _cm_to_twips


# LaTeX font 命令 → Word size_pt
_FONT_MAP = {
    'tiny': 6, 'scriptsize': 8, 'footnotesize': 9,
    'small': 10, 'normalsize': 12, 'large': 14,
    'Large': 16, 'LARGE': 18, 'huge': 20, 'Huge': 24,
}

# 列对齐映射
_ALIGN_MAP = {'l': 'left', 'c': 'center', 'r': 'right'}


def tabular_to_json(tabular_body, col_format, table_env='', layout_spec=None):
    """tabular 表格 → JSON (部分无损)

    Args:
        tabular_body: \\begin{tabular}{fmt}...\\end{tabular} 内的内容
        col_format: 列格式字符串, 如 |l|c|r|
        table_env: table 环境（含 caption/label）
        layout_spec: 排版规格

    Returns:
        dict: 符合 gen_table_from_json.py 输入格式的 JSON
    """
    layout_spec = layout_spec or {}

    # 1. 解析列格式
    cols_info = _parse_col_format(col_format)
    num_cols = len(cols_info)
    if num_cols == 0:
        return _empty_json()

    # 2. 拆分行 + 记录行前规则
    raw_lines = re.split(r'\\\\', tabular_body)
    parsed_rows = []  # [(rule_type, [cell_texts], clines)]
    pending_rule = ''
    for line in raw_lines:
        line = line.strip()
        if not line:
            continue

        rules_found = []
        for m in re.finditer(r'\\(toprule|midrule|bottomrule|hline)\b', line):
            rules_found.append(m.group(1))

        clines = [(int(a), int(b)) for a, b in re.findall(r'\\cline\{(\d+)-(\d+)\}', line)]

        cleaned = re.sub(r'\\(toprule|midrule|bottomrule|hline)\b', '', line).strip()
        cleaned = re.sub(r'\\cline\{\d+-\d+\}', '', cleaned).strip()

        main_rule = rules_found[-1] if rules_found else ''

        if not cleaned:
            pending_rule = _merge_rules(pending_rule, main_rule)
            continue

        cells = _split_cells_balanced(cleaned, num_cols)
        actual_rule = _merge_rules(pending_rule, main_rule)
        pending_rule = main_rule
        parsed_rows.append((actual_rule, cells, clines))

    if not parsed_rows:
        return _empty_json()

    # 3. 生成 grid_cols
    grid_cols = []
    total_width = _cm_to_twips(layout_spec.get('table_width_cm', 15))
    fixed_total = sum(c['width_twips'] for c in cols_info if c['width_twips'] > 0)
    remaining = total_width - fixed_total
    flex_count = sum(1 for c in cols_info if c['width_twips'] == 0)
    flex_width = remaining // flex_count if flex_count > 0 else 0

    for c in cols_info:
        w = c['width_twips'] if c['width_twips'] > 0 else flex_width
        grid_cols.append({'width_twips': w})

    # 4. 提取 caption
    caption = ''
    cap_m = re.search(r'\\caption\{', table_env)
    if cap_m:
        caption = _match_balanced_braces(table_env, cap_m.end() - 1)

    # 5. 组装 rows → JSON
    rows_data = []
    multirow_tracker = {}

    for ri, (rule_type, cells, clines) in enumerate(parsed_rows):
        is_first = (ri == 0)
        is_last = (ri == len(parsed_rows) - 1)

        cells_data = []
        ci = 0
        while ci < num_cols:
            if ci in multirow_tracker and multirow_tracker[ci] > 0:
                multirow_tracker[ci] -= 1
                cells_data.append({
                    'col_start': ci,
                    'gridSpan': 1,
                    'vMerge': 'continue',
                    'borders': _build_cell_borders_tabular(
                        ri, ci, num_cols, rule_type, clines,
                        is_first, is_last, cols_info),
                })
                ci += 1
                continue

            cell_text = cells[ci] if ci < len(cells) else ''
            gs = 1
            vm = None

            mc_m = re.match(r'\\multicolumn\{(\d+)\}\{([^}]*)\}\{', cell_text)
            if mc_m:
                gs = int(mc_m.group(1))
                inner = _match_balanced_braces(cell_text, cell_text.index('{', cell_text.index('}') + 1))
                cell_text = inner

            mr_m = re.match(r'\\multirow\{(\d+)\}\{[^}]*\}\{', cell_text)
            if mr_m:
                n_rows = int(mr_m.group(1))
                inner = _match_balanced_braces(cell_text, cell_text.index('{', 2 + cell_text.index('{', 1)) )
                cell_text = inner
                vm = 'restart'
                multirow_tracker[ci] = n_rows - 1

            text, is_bold, is_italic, size_pt = _parse_cell_text(cell_text)

            paragraphs = []
            if text:
                run = {'text': text}
                fmt = {'font_ascii': 'Times New Roman'}
                if is_bold:
                    fmt['bold'] = True
                if is_italic:
                    fmt['italic'] = True
                if size_pt and size_pt != 12:
                    fmt['size_pt'] = size_pt
                run['format'] = fmt
                paragraphs.append({
                    'align': cols_info[min(ci, len(cols_info)-1)]['align'],
                    'runs': [run],
                })

            borders = _build_cell_borders_tabular(
                ri, ci, num_cols, rule_type, clines,
                is_first, is_last, cols_info, gs)

            cell_data = {
                'col_start': ci,
                'gridSpan': gs,
                'vMerge': vm or '',
                'borders': borders,
            }
            if paragraphs:
                cell_data['paragraphs'] = paragraphs
            cells_data.append(cell_data)

            for skip in range(1, gs):
                if ci + skip < num_cols:
                    skip_cell = {
                        'col_start': ci + skip,
                        'gridSpan': 1,
                        'vMerge': '',
                        'borders': {},
                    }
                    cells_data.append(skip_cell)

            ci += gs

        row_data = {
            'row_height': '400',
            'row_height_rule': 'atLeast',
            'cells': cells_data,
        }
        if ri == 0 and any(c.get('paragraphs', []) and
                           any(r.get('format', {}).get('bold') for r in c.get('paragraphs', [{}]))
                           for c in cells_data if 'vMerge' not in c):
            row_data['is_header'] = True

        rows_data.append(row_data)

    result = {
        'grid_cols': grid_cols,
        'table_properties': {
            'borders': {},
            'width': str(sum(gc['width_twips'] for gc in grid_cols)),
        },
        'rows': rows_data,
    }

    if caption:
        result['position'] = {'table_caption': caption}

    return result


# ── tabular 专用辅助函数 ─────────────────────────────────────

def _empty_json():
    """返回空表格 JSON"""
    return {
        'grid_cols': [],
        'table_properties': {'borders': {}, 'width': '0'},
        'rows': [],
    }


def _parse_col_format(fmt):
    """解析列格式字符串 → [{align, width_twips, vline_left}]"""
    cols = []
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch in 'lcr':
            cols.append({'align': _ALIGN_MAP[ch], 'width_twips': 0, 'vline_left': False})
            i += 1
        elif ch == '|':
            if cols:
                cols[-1]['vline_left'] = True
            i += 1
        elif ch == 'p':
            m = re.match(r'p\{([0-9.]+)(cm|in|mm)\}', fmt[i:])
            if m:
                w = float(m.group(1))
                unit = m.group(2)
                if unit == 'cm':
                    tw = _cm_to_twips(w)
                elif unit == 'in':
                    tw = round(w * 1440)
                elif unit == 'mm':
                    tw = _cm_to_twips(w / 10)
                else:
                    tw = 0
                cols.append({'align': 'left', 'width_twips': tw, 'vline_left': False})
                i += m.end()
            else:
                i += 1
        elif ch == '@':
            brace_start = fmt.index('{', i)
            end = _find_matching_brace(fmt, brace_start)
            i = end + 1 if end > 0 else i + 1
        else:
            i += 1
    return cols


def _find_matching_brace(text, start):
    """找到与 start 位置 { 配对的 } 位置"""
    depth = 0
    for i in range(start, len(text)):
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                return i
    return -1


def _merge_rules(existing, new):
    """合并规则类型（后出现的优先）"""
    if new:
        return new
    return existing


def _split_cells_balanced(line, num_cols):
    """用平衡大括号匹配拆分单元格"""
    cells = []
    depth = 0
    current = []
    for ch in line:
        if ch == '&' and depth == 0:
            cells.append(''.join(current).strip())
            current = []
        else:
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
            current.append(ch)
    cells.append(''.join(current).strip())

    while len(cells) < num_cols:
        cells.append('')
    return cells[:num_cols]


def _parse_cell_text(text):
    """解析单元格文本 → (clean_text, is_bold, is_italic, size_pt)"""
    is_bold = False
    is_italic = False
    size_pt = None

    for cmd, pt in _FONT_MAP.items():
        if f'\\{cmd}' in text:
            size_pt = pt
            text = text.replace(f'\\{cmd}', '')
            break

    if '\\textbf{' in text:
        is_bold = True
        text = re.sub(r'\\textbf\{([^}]*)\}', r'\1', text)

    if '\\textit{' in text:
        is_italic = True
        text = re.sub(r'\\textit\{([^}]*)\}', r'\1', text)

    if '\\emph{' in text:
        is_italic = True
        text = re.sub(r'\\emph\{([^}]*)\}', r'\1', text)

    if '\\bfseries' in text:
        is_bold = True
        text = text.replace('\\bfseries', '')

    if '\\itshape' in text:
        is_italic = True
        text = text.replace('\\itshape', '')

    text = re.sub(r'\$_\{([^}]+)\}\$', lambda m: _to_subscript(m.group(1)), text)
    text = re.sub(r'\$_([^$]+)\$', lambda m: _to_subscript(m.group(1)), text)

    text = re.sub(r'CO\$_2\$', 'CO₂', text)
    text = re.sub(r'XCO\$_2\$', 'XCO₂', text)

    text = re.sub(r'\\underline\{([^}]*)\}', r'\1', text)
    text = re.sub(r'\\textsubscript\{([^}]*)\}', lambda m: _to_subscript(m.group(1)), text)
    text = re.sub(r'\\\w+\{([^}]*)\}', r'\1', text)
    text = text.replace('\\%', '%').replace('\\_', '_').replace('\\&', '&')
    text = re.sub(r'[${}\\]', '', text).strip()

    return text, is_bold, is_italic, size_pt


def _build_cell_borders_tabular(ri, ci, num_cols, rule_type, clines,
                                 is_first, is_last, cols_info, gridSpan=1):
    """构建 tabular 单元格边框"""
    borders = {}

    if is_first:
        if rule_type == 'toprule':
            borders['top'] = {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'}
        elif rule_type == 'hline':
            borders['top'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}
        else:
            borders['top'] = {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'}
    else:
        if rule_type == 'midrule':
            borders['top'] = {'val': 'single', 'sz': '6', 'color': '000000', 'space': '0'}
        elif rule_type == 'hline':
            borders['top'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}
        elif rule_type == 'toprule':
            borders['top'] = {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'}
        else:
            borders['top'] = {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'}

    if is_last:
        borders['bottom'] = {'val': 'single', 'sz': '12', 'color': '000000', 'space': '0'}
    else:
        borders['bottom'] = {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'}

    for cs, ce in clines:
        cs, ce = int(cs) - 1, int(ce) - 1
        if cs <= ci <= ce:
            borders['top'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}

    if ci == 0:
        borders['left'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}
    if ci == num_cols - 1 or ci + gridSpan >= num_cols:
        borders['right'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}
    if ci < len(cols_info) and cols_info[ci].get('vline_left') and ci > 0:
        borders['left'] = {'val': 'single', 'sz': '4', 'color': '000000', 'space': '0'}

    for edge in ('top', 'bottom', 'left', 'right'):
        if edge not in borders:
            borders[edge] = {'val': 'nil', 'sz': '0', 'color': 'auto', 'space': '0'}

    return borders
