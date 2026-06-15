"""
从JSON数据生成LaTeX表格
根据表格位置信息确定LaTeX中的放置位置
支持合并单元格(\multirow/\multicolumn)、边框(\hline/\cline)、对齐
"""
import json
import sys
import re


def sanitize_label(text):
    """生成LaTeX label（去掉特殊字符）"""
    t = re.sub(r'[^\w\s-]', '', text.lower())
    return re.sub(r'\s+', '-', t.strip())[:50]


def analyze_hlines(table_data):
    """分析哪些行需要hline，哪些列需要cline"""
    rows = table_data['rows']
    num_cols = len(table_data['grid_cols'])
    hline_rows = []  # 哪些行后有hline
    cline_info = {}  # row_idx -> [(col_start, col_end)]

    for ri, row in enumerate(rows):
        has_bottom = False
        bottom_cols = set()

        for cell in row['cells']:
            borders = cell.get('borders', {})
            bottom = borders.get('bottom', {})
            if bottom.get('val') == 'single':
                sz = bottom.get('sz', '4')
                if sz == '8':  # 粗线 = \hline
                    has_bottom = True
                elif sz == '4':  # 细线 = \cline
                    for c in range(cell['col_start'], cell['col_start'] + cell['gridSpan']):
                        bottom_cols.add(c)

        if has_bottom:
            hline_rows.append(ri)
        elif bottom_cols:
            cline_info[ri] = sorted(bottom_cols)

    return hline_rows, cline_info


def to_latex(table_data, position=None):
    """将表格数据转换为LaTeX格式"""
    rows = table_data['rows']
    num_cols = len(table_data['grid_cols'])
    col_spec = '|' + '|'.join(['c'] * num_cols) + '|'

    # 获取caption和label
    caption = ''
    if position:
        caption = position.get('table_caption', '')
        # 从caption提取标题文本
        if not caption:
            # 用第一个cell的text作为caption fallback
            if rows and rows[0]['cells']:
                caption = rows[0]['cells'][0]['text'][:80]

    label = sanitize_label(caption) if caption else f"tab{table_data['table_index']}"

    # 分析边框线
    hline_rows, cline_info = analyze_hlines(table_data)

    # 生成表格行
    latex_lines = []

    for ri, row in enumerate(rows):
        # 如果是标题行（Row 0 跨全列的caption），跳过
        if ri == 0 and len(row['cells']) == 1 and row['cells'][0]['gridSpan'] == num_cols:
            caption_text = row['cells'][0]['text'].strip()
            if caption_text and not caption:
                caption = caption_text
                label = sanitize_label(caption)
            continue

        # 跳过vMerge=continue的行（已在multirow中）
        if any(c.get('vMerge') == 'continue' for c in row['cells']):
            # 只输出非continue的cell内容（用&连接，留空multirow位置）
            cell_parts = []
            for c in row['cells']:
                if c.get('vMerge') == 'continue':
                    cell_parts.append('')  # multirow已经占了位置
                else:
                    text = c['text'].strip().replace('&', '\\&')
                    cell_parts.append(text)
            latex_lines.append(' & '.join(cell_parts) + ' \\\\')
            continue

        # 正常行或vMerge=restart行
        cell_parts = []
        for c in row['cells']:
            text = c['text'].strip().replace('&', '\\&')

            # 横向合并
            if c['gridSpan'] > 1:
                text = f'\\multicolumn{{{c["gridSpan"]}}}{{|c|}}{{{text}}}'

            # 纵向合并起点
            if c['vMerge'] == 'restart':
                # 计算合并行数
                merge_rows = 1
                for later_ri in range(ri + 1, len(rows)):
                    for later_c in rows[later_ri]['cells']:
                        if later_c['col_start'] == c['col_start'] and later_c.get('vMerge') == 'continue':
                            merge_rows += 1
                            break
                    else:
                        break
                text = f'\\multirow{{{merge_rows}}}{{*}}{{{text}}}'

            cell_parts.append(text)

        line = ' & '.join(cell_parts) + ' \\\\'

        # 添加hline或cline
        if ri in hline_rows:
            line += ' \\hline'
        elif ri in cline_info:
            cols = cline_info[ri]
            # 生成\cline{start-end}
            ranges = []
            start = cols[0]
            prev = start
            for c in cols[1:]:
                if c != prev + 1:
                    ranges.append((start + 1, prev + 1))  # LaTeX列号1-based
                    start = c
                prev = c
            ranges.append((start + 1, prev + 1))
            clines = ' '.join(f'\\cline{{{s}-{e}}}' for s, e in ranges)
            line += ' ' + clines

        latex_lines.append(line)

    # 第一行前面加hline
    top_lines = ['\\hline']
    if 0 in hline_rows:
        # Row 0有粗底边框，前面加\hline
        pass

    # 组装完整LaTeX
    latex = []
    latex.append('\\begin{table}[htbp]')
    latex.append(f'  \\centering')
    latex.append(f'  \\caption{{{caption}}}')
    latex.append(f'  \\label{{{label}}}')

    # 位置提示（用于LaTeX编译时确定表格位置）
    if position:
        heading = position.get('current_heading', '')
        if heading:
            latex.append(f'  % 位置: {heading}章节')

    latex.append(f'  \\begin{{tabular}}{{{col_spec}}}')
    latex.append('    \\hline')
    for line in latex_lines:
        latex.append(f'    {line}')
    latex.append('    \\hline')
    latex.append(f'  \\end{{tabular}}')
    latex.append('\\end{table}')

    return '\n'.join(latex)


def to_latex_all(data):
    """将所有表格转换为LaTeX"""
    results = []
    for t in data['tables']:
        position = t.get('position', {})
        latex = to_latex(t, position)
        results.append({
            'table_index': t['table_index'],
            'caption': position.get('table_caption', ''),
            'heading': position.get('current_heading', ''),
            'latex': latex
        })
    return results


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python to_latex.py <input.json> [output.tex]")
        sys.exit(1)

    json_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'tables_latex.tex'

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    results = to_latex_all(data)

    with open(output_path, 'w', encoding='utf-8') as f:
        for r in results:
            f.write(f"% Table {r['table_index']}: {r['caption']}\n")
            f.write(f"% 位置: {r['heading']}\n\n")
            f.write(r['latex'])
            f.write('\n\n')

    print(f"Generated {len(results)} LaTeX tables -> {output_path}")
    for r in results:
        print(f"  Table {r['table_index']}: heading='{r['heading']}' caption='{r['caption'][:30]}'")