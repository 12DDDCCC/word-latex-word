#!/usr/bin/env python3
r"""TeX提取模块 — 从LaTeX源文件提取公式、表格、图片、TikZ表格

所有从LaTeX源文件中提取结构化数据的函数集中在此模块。
"""
import re, os, sys
from pathlib import Path

# 导入共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import match_balanced_braces as _match_balanced_braces

# 导入 latex_to_omml skill (公式提取委托)
_SKILL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'omml-to-latex')
if os.path.isdir(_SKILL_DIR):
    sys.path.insert(0, _SKILL_DIR)
try:
    from latex_to_omml import extract_formulas_from_tex as _extract_formulas_skill
    _HAS_SKILL = True
except ImportError:
    _HAS_SKILL = False


def extract_images_from_tex(tex_path):
    """从LaTeX源文件提取图片路径和完整图片环境信息（包括caption和legend）

    Returns:
        list[dict]: 图片信息列表, 每项包含 {
            'path': str,      # 图片路径
            'width': str,     # 宽度参数
            'caption': str,   # 图片caption
            'legend': str,    # 图片legend
            'start': int,     # 在tex中的起始位置
            'end': int,       # 在tex中的结束位置
        }
    """
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    images = []

    def _append_image(figure_body, is_full_width, start, end):
        img_match = re.search(r'\\includegraphics(?:\[[^\]]*\])?\{([^}]+)\}', figure_body)
        img_path = img_match.group(1) if img_match else ''

        width = ''
        if img_match:
            width_match = re.search(r'\\includegraphics\[([^\]]*)\]', figure_body)
            if width_match:
                width = width_match.group(1)

        number = ''
        num_match = re.search(r'\\renewcommand\{\\thefigure\}\{([^}]+)\}', figure_body)
        if num_match:
            number = num_match.group(1).strip()

        caption = ''
        cap_match = re.search(r'\\caption\{', figure_body)
        if cap_match:
            cap_brace_start = cap_match.end() - 1
            caption = _match_balanced_braces(figure_body, cap_brace_start)
            cap_end = cap_brace_start + len(caption) + 2
        else:
            cap_end = 0

        legend = ''
        after_caption = figure_body[cap_end:] if cap_match else ''
        legend_lines = []
        for line in after_caption.split('\n'):
            line = line.strip()
            if line and not line.startswith('\\') and not line.startswith('%'):
                legend_lines.append(line)
        if legend_lines:
            legend = '\n'.join(legend_lines)

        if img_path:
            images.append({
                'path': img_path,
                'width': width,
                'number': number,
                'caption': caption,
                'legend': legend,
                'is_full_width': is_full_width,
                'start': start,
                'end': end,
            })

    figure_pattern = r'\\begin\{(figure\*?)\}(.*?)\\end\{\1\}'
    for m in re.finditer(figure_pattern, content, re.DOTALL):
        _append_image(m.group(2), m.group(1).endswith('*'), m.start(), m.end())

    strip_pattern = r'\\begin\{strip\}((?:(?!\\end\{strip\}).)*?)\\end\{strip\}'
    for m in re.finditer(strip_pattern, content, re.DOTALL):
        body = m.group(1)
        has_captype = re.search(
            r'\\@captype\s*\{figure\}|\\def\s*\\@captype\s*\{figure\}',
            body,
        )
        if has_captype and '\\includegraphics' in body:
            _append_image(body, True, m.start(), m.end())

    return images


def extract_tikz_tables_from_tex(tex_path):
    """从LaTeX源文件提取TikZ表格

    Returns:
        list[dict]: TikZ表格信息列表
    """
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    tikz_tables = []

    def _append_tikz_table(table_full, tikz_body, is_full_width, start, end):
        number = ''
        num_m = re.search(r'\\renewcommand\{\\thetable\}\{([^}]+)\}', table_full)
        if num_m:
            number = num_m.group(1).strip()

        caption = ''
        cap_m = re.search(r'\\caption\{', table_full)
        if cap_m:
            caption = _match_balanced_braces(table_full, cap_m.end() - 1)

        tikz_tables.append({
            'caption': caption,
            'number': number,
            'tikz_body': tikz_body,
            'is_full_width': is_full_width,
            'start': start,
            'end': end,
        })

    pattern = r'\\begin\{(table\*?)\}(.*?)\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}(.*?)\\end\{\1\}'
    for m in re.finditer(pattern, content, re.DOTALL):
        env_name = m.group(1)
        table_env = m.group(2)
        tikz_body = m.group(3)
        after_tikz = m.group(4)
        _append_tikz_table(
            table_env + after_tikz, tikz_body, env_name.endswith('*'),
            m.start(), m.end())

    strip_body = r'(?:(?!\\end\{strip\}).)*?'
    strip_pattern = (
        r'\\begin\{strip\}(' + strip_body + r')'
        r'\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}'
        r'(' + strip_body + r')\\end\{strip\}'
    )
    for m in re.finditer(strip_pattern, content, re.DOTALL):
        table_full = m.group(1) + m.group(3)
        has_captype = re.search(
            r'\\@captype\s*\{table\}|\\def\s*\\@captype\s*\{table\}',
            table_full,
        )
        if has_captype:
            _append_tikz_table(table_full, m.group(2), True, m.start(), m.end())

    super_pattern = re.compile(
        r'% WORD_SUPERTABLE_BEGIN number=([^\n]*)\n'
        r'% WORD_SUPERTABLE_TIKZ=([A-Za-z0-9+/=]+)\n'
        r'(.*?)% WORD_SUPERTABLE_END',
        re.DOTALL,
    )
    for m in super_pattern.finditer(content):
        import base64

        block = m.group(3)
        caption = ''
        cap_m = re.search(r'\\tablecaption\{', block)
        if cap_m:
            caption = _match_balanced_braces(block, cap_m.end() - 1)
        tikz_body = base64.b64decode(m.group(2)).decode('utf-8')
        tikz_tables.append({
            'caption': caption,
            'number': m.group(1).strip(),
            'tikz_body': tikz_body,
            'is_full_width': '% meta:full_width=1' in tikz_body,
            'start': m.start(),
            'end': m.end(),
        })

    return tikz_tables


def parse_tikz_table(tikz_body):
    """解析TikZ表格为结构化数据

    Returns:
        dict: {
            'headers': list[str],  # 表头
            'rows': list[list[str]],  # 数据行
        }
    """
    node_data = []
    node_pattern = r'\\node\[([^\]]*anchor=center[^\]]*)\] at \(([^)]+)\) \{'
    for m in re.finditer(node_pattern, tikz_body):
        pos = m.group(2)
        brace_start = m.end() - 1
        text = _match_balanced_braces(tikz_body, brace_start)
        text = re.sub(r'\\textbf\{([^}]+)\}', r'\1', text)
        text = re.sub(r'\\textit\{([^}]+)\}', r'\1', text)
        text = re.sub(r'\\emph\{([^}]+)\}', r'\1', text)
        text = re.sub(r'\\underline\{([^}]+)\}', r'\1', text)
        text = re.sub(r'\\\w+\{([^}]*)\}', r'\1', text)
        text = re.sub(r'[${}]', '', text).strip()
        is_bold = '\\textbf' in tikz_body[m.start():m.start() + len(m.group(0)) + len(text) + 10]
        node_data.append({'pos': pos, 'text': text, 'is_bold': is_bold})

    if not node_data:
        return {'headers': [], 'rows': []}

    rows = {}
    for node in node_data:
        pos = node['pos']
        text = node['text']
        is_bold = node.get('is_bold', False)
        try:
            x, y = pos.split(',')
            x = float(x.strip())
            y = float(y.strip())
            y_key = round(y, 2)
            if y_key not in rows:
                rows[y_key] = []
            rows[y_key].append({'x': x, 'text': text, 'is_bold': is_bold})
        except ValueError:
            continue

    sorted_rows = []
    for y_key in sorted(rows.keys(), reverse=True):
        row_cells = sorted(rows[y_key], key=lambda c: c['x'])
        sorted_rows.append(row_cells)

    if not sorted_rows:
        return {'headers': [], 'rows': []}

    header_cells = sorted_rows[0] if sorted_rows else []
    headers = [c['text'] for c in header_cells]
    data_rows = [[c['text'] for c in row] for row in sorted_rows[1:]] if len(sorted_rows) > 1 else []

    return {
        'headers': headers,
        'rows': data_rows,
    }


def extract_formulas_from_tex(tex_path):
    """从LaTeX源文件提取数学公式 (委托给 latex_to_omml skill)"""
    if _HAS_SKILL:
        return _extract_formulas_skill(tex_path)
    # fallback: 简单提取
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    formulas = []
    idx = 0
    for m in re.finditer(r'\\begin\{(equation|gather)\*?\}(.*?)\\end\{\1\*?\}', content, re.DOTALL):
        formulas.append({
            'id': idx, 'type': 'display', 'env': m.group(1),
            'latex': m.group(2).strip(), 'label': None, 'eq_num': f'({idx+1})',
            'start': m.start(), 'end': m.end(),
        })
        idx += 1
    return formulas


def extract_tables_from_tex(tex_path):
    """从LaTeX源文件提取表格结构"""
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    tables = []
    idx = 0

    for m in re.finditer(r'\\begin\{table\*?\}(.*?)\\end\{table\*?\}', content, re.DOTALL):
        table_body = m.group(1)
        table_data = _parse_tabular(table_body, idx)
        if table_data:
            tables.append(table_data)
            idx += 1

    for m in re.finditer(r'\\begin\{tabular\*?\}(.*?)\\end\{tabular\*?\}', content, re.DOTALL):
        if any(m.start() >= t['start'] and m.end() <= t['end'] for t in tables):
            continue
        table_data = _parse_tabular(m.group(0), idx)
        if table_data:
            table_data['standalone'] = True
            tables.append(table_data)
            idx += 1

    return tables


def _count_tabular_columns(col_format):
    """Count visible LaTeX tabular columns from a column format string."""
    fmt = re.sub(r'[@!><]\{[^{}]*\}', '', col_format or '')
    fmt = re.sub(r'[|*\s]', '', fmt)
    return len(re.findall(r'[lcrX]|[pmb]\{[^{}]*\}', fmt))


def _parse_tabular(table_body, table_id):
    """解析tabular内容为结构化数据"""
    col_fmt_m = re.search(r'\\begin\{tabular\*?\}\{([^}]+)\}', table_body)
    if not col_fmt_m:
        return None
    col_format = col_fmt_m.group(1)
    col_count = _count_tabular_columns(col_format)
    if col_count == 0:
        return None

    caption = ''
    cap_m = re.search(r'\\caption\{', table_body)
    if cap_m:
        caption = _match_balanced_braces(table_body, cap_m.end() - 1)

    label = ''
    lab_m = re.search(r'\\label\{([^}]+)\}', table_body)
    if lab_m:
        label = lab_m.group(1)

    tab_m = re.search(r'\\begin\{tabular\*?\}\{[^}]+\}(.*?)\\end\{tabular\*?\}', table_body, re.DOTALL)
    if not tab_m:
        return None
    tab_content = tab_m.group(1)

    rows = []
    for line in tab_content.split('\\\\'):
        line = line.strip()
        if not line or line in ('\\hline', '\\toprule', '\\midrule', '\\bottomrule'):
            continue
        line = re.sub(r'\\(toprule|midrule|bottomrule|hline)\b', '', line)
        cells = _split_cells(line, col_count)
        if cells:
            rows.append(cells)

    return {
        'id': table_id,
        'col_format': col_format,
        'col_count': col_count,
        'caption': caption,
        'label': label,
        'rows': rows,
        'start': 0,
        'end': len(table_body),
    }


def _split_cells(line, col_count):
    """分割一行中的单元格"""
    raw_cells = line.split('&')
    cells = []
    for c in raw_cells:
        c = c.strip()
        mc = re.match(r'\\multicolumn\{(\d+)\}\{[^}]*\}\{([^}]*)\}', c)
        if mc:
            n = int(mc.group(1))
            text = mc.group(2).strip()
            for _ in range(n):
                cells.append(text)
        else:
            clean = re.sub(r'\\(textbf|textit|emph|underline)\{([^}]*)\}', r'\2', c)
            clean = re.sub(r'\\\w+\{([^}]*)\}', r'\1', clean)
            clean = re.sub(r'[${}\\]', '', clean).strip()
            cells.append(clean)

    while len(cells) < col_count:
        cells.append('')

    return cells[:col_count]


def _collect_table_captions(tex_path):
    """收集所有表格的caption信息"""
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    captions = []
    for m in re.finditer(r'\\caption\{([^}]+)\}', content):
        captions.append(m.group(1))
    return captions
