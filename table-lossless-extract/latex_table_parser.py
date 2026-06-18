#!/usr/bin/env python3
"""LaTeX 表格 → JSON 解析器 (反向转换)

将 LaTeX 源码中的表格解析为 gen_table_from_json.py 可消费的 JSON 格式，
实现 LaTeX → Word 的零损失表格转换。

支持两种表格类型：
  1. TikZ 表格 → 100% 无损（所有边框/位置/字体信息都在源码中）
  2. tabular 表格 → 部分无损（只有线型规则，无逐线宽度值）

用法:
  python latex_table_parser.py <input.tex> [output.json]
  python latex_table_parser.py --tikz-body "<tikz_code>" [--table-env "<table_env>"]
"""
import re
import sys
import json
import os
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 从子模块导入核心解析函数
from _tikz_parser import tikz_to_json
from _tabular_parser import tabular_to_json


def parse_tex_tables(tex_path, layout_spec=None):
    """从 .tex 文件提取所有表格并转为 JSON

    Returns:
        list[dict]: 每个 dict 是一个表格的 JSON 数据
    """
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    tables = []
    idx = 0

    # TikZ 表格
    tikz_pattern = r'\\begin\{table\*?\}(.*?)\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}(.*?)\\end\{table\*?\}'
    for m in re.finditer(tikz_pattern, content, re.DOTALL):
        table_env = m.group(1) + m.group(3)
        tikz_body = m.group(2)
        json_data = tikz_to_json(tikz_body, table_env, layout_spec)
        json_data['table_index'] = idx + 1
        tables.append(json_data)
        idx += 1

    # 标准 tabular 表格
    tab_pattern = r'\\begin\{table\*?\}(.*?)\\end\{table\*?\}'
    for m in re.finditer(tab_pattern, content, re.DOTALL):
        if any(m.start() >= t.get('_src_start', -1) and m.end() <= t.get('_src_end', -1) for t in tables):
            continue

        table_body = m.group(1)

        if '\\begin{tikzpicture}' in table_body:
            continue

        tab_m = re.search(r'\\begin\{tabular\*?\}\{([^}]+)\}(.*?)\\end\{tabular\*?\}', table_body, re.DOTALL)
        if not tab_m:
            continue

        col_format = tab_m.group(1)
        tabular_body = tab_m.group(2)
        json_data = tabular_to_json(tabular_body, col_format, table_body, layout_spec)
        json_data['table_index'] = idx + 1
        tables.append(json_data)
        idx += 1

    # 独立 tabular（不在 table 环境内）
    for m in re.finditer(r'\\begin\{tabular\*?\}\{([^}]+)\}(.*?)\\end\{tabular\*?\}', content, re.DOTALL):
        if any(m.start() >= t.get('_src_start', -1) for t in tables):
            continue

        col_format = m.group(1)
        tabular_body = m.group(2)
        json_data = tabular_to_json(tabular_body, col_format, '', layout_spec)
        json_data['table_index'] = idx + 1
        tables.append(json_data)
        idx += 1

    return tables


# ── CLI ───────────────────────────────────────────────────

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='LaTeX 表格 → JSON 解析器')
    parser.add_argument('input', help='输入 .tex 文件路径')
    parser.add_argument('-o', '--output', help='输出 .json 文件路径')
    parser.add_argument('--table-width', type=float, default=15, help='默认表格宽度(cm)')
    args = parser.parse_args()

    layout_spec = {'table_width_cm': args.table_width}
    tables = parse_tex_tables(args.input, layout_spec)

    output = {
        'source_file': args.input,
        'total_tables': len(tables),
        'tables': tables,
    }

    if args.output:
        out_path = args.output
    else:
        base = os.path.splitext(args.input)[0]
        out_path = base + '_tables.json'

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f'Extracted {len(tables)} tables → {out_path}')
