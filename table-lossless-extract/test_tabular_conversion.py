#!/usr/bin/env python3
"""tabular 转换测试

验证: LaTeX tabular 源码 → JSON → Word 的转换质量。
"""
import json
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SKILL_DIR)

from latex_table_parser import tabular_to_json
from gen_table_from_json import generate_docx


# ── 测试用例 ──────────────────────────────────────────────

TEST_CASES = [
    {
        'name': '简单 booktabs 表格',
        'col_format': 'lcc',
        'tabular_body': r"""\toprule
Site & Value & Unit \\
\midrule
A & 1.0 & ppm \\
B & 2.0 & ppb \\
\bottomrule""",
        'table_env': r'\caption{Test Table}',
        'expected': {
            'num_cols': 3,
            'num_rows': 3,
            'caption': 'Test Table',
        },
    },
    {
        'name': '带竖线的 hline 表格',
        'col_format': '|l|c|r|',
        'tabular_body': r"""\hline
Name & Age & Score \\
\hline
Alice & 20 & 95 \\
Bob & 22 & 88 \\
\hline""",
        'table_env': r'\caption{With Vertical Lines}',
        'expected': {
            'num_cols': 3,
            'num_rows': 3,
        },
    },
    {
        'name': '带 p{} 列宽的表格',
        'col_format': 'p{3cm}cp{4cm}',
        'tabular_body': r"""\toprule
Description & ID & Details \\
\midrule
Long text here & 1 & More details \\
Short & 2 & OK \\
\bottomrule""",
        'table_env': '',
        'expected': {
            'num_cols': 3,
            'num_rows': 3,
        },
    },
    {
        'name': '\\multicolumn 表格',
        'col_format': 'lccc',
        'tabular_body': r"""\toprule
\multicolumn{4}{c}{Group Header} \\
\midrule
Item & A & B & C \\
X & 1 & 2 & 3 \\
\bottomrule""",
        'table_env': '',
        'expected': {
            'num_cols': 4,
            'num_rows': 3,
        },
    },
    {
        'name': '粗体+斜体文本',
        'col_format': 'lcc',
        'tabular_body': r"""\toprule
\textbf{Site} & \textbf{Value} & \textbf{Unit} \\
\midrule
\textit{Total} & 3.0 & ppm \\
\bottomrule""",
        'table_env': r'\caption{Formatted Text}',
        'expected': {
            'num_cols': 3,
            'num_rows': 2,
            'has_bold': True,
            'has_italic': True,
        },
    },
]


def run_test(case, output_dir=None):
    """运行单个测试"""
    name = case['name']
    print(f"  [{name}]", end=" ")

    json_data = tabular_to_json(
        case['tabular_body'],
        case['col_format'],
        case['table_env'],
    )

    expected = case['expected']
    checks = []

    # 列数
    num_cols = len(json_data.get('grid_cols', []))
    checks.append(('num_cols', num_cols == expected.get('num_cols', 0),
                    f"{num_cols} vs {expected.get('num_cols', '?')}"))

    # 行数
    num_rows = len(json_data.get('rows', []))
    checks.append(('num_rows', num_rows == expected.get('num_rows', 0),
                    f"{num_rows} vs {expected.get('num_rows', '?')}"))

    # caption
    if 'caption' in expected:
        cap = json_data.get('position', {}).get('table_caption', '')
        checks.append(('caption', cap == expected['caption'],
                        f"'{cap}' vs '{expected['caption']}'"))

    # bold 检测
    if expected.get('has_bold'):
        has_bold = False
        for row in json_data.get('rows', []):
            for c in row.get('cells', []):
                for p in c.get('paragraphs', []):
                    for r in p.get('runs', []):
                        if r.get('format', {}).get('bold'):
                            has_bold = True
        checks.append(('bold', has_bold, str(has_bold)))

    # italic 检测
    if expected.get('has_italic'):
        has_italic = False
        for row in json_data.get('rows', []):
            for c in row.get('cells', []):
                for p in c.get('paragraphs', []):
                    for r in p.get('runs', []):
                        if r.get('format', {}).get('italic'):
                            has_italic = True
        checks.append(('italic', has_italic, str(has_italic)))

    # 基本边框存在性
    has_borders = False
    for row in json_data.get('rows', []):
        for cell in row.get('cells', []):
            for edge in ('top', 'bottom', 'left', 'right'):
                b = cell.get('borders', {}).get(edge, {})
                if b.get('val') == 'single':
                    has_borders = True
                    break
    checks.append(('borders_exist', has_borders, str(has_borders)))

    # 结果
    all_pass = all(ok for _, ok, _ in checks)
    if all_pass:
        print("PASS", end="")
    else:
        print("FAIL", end="")

    for cname, ok, msg in checks:
        status = "✓" if ok else "✗"
        print(f"  {status}{cname}:{msg}", end="")
    print()

    # 生成 Word
    if output_dir and json_data.get('grid_cols'):
        try:
            idx = TEST_CASES.index(case) + 1
            safe_name = name.replace(' ', '_').replace('\\', '').replace('/', '').replace('{', '').replace('}', '')
            docx_path = os.path.join(output_dir, f'tabular_test_{idx}_{safe_name}.docx')
            generate_docx(json_data, docx_path)
            print(f"    → Word: {docx_path}")
        except Exception as e:
            print(f"    → Word 生成失败: {e}")

    # 输出 JSON 供调试
    if output_dir:
        idx = TEST_CASES.index(case) + 1
        json_path = os.path.join(output_dir, f'tabular_test_{idx}.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)

    return all_pass


def main():
    output_dir = sys.argv[1] if len(sys.argv) > 1 else None

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    print(f"{'='*60}")
    print(f"tabular 转换测试 ({len(TEST_CASES)} 用例)")
    print(f"{'='*60}\n")

    total_pass = 0
    total_fail = 0
    for case in TEST_CASES:
        if run_test(case, output_dir):
            total_pass += 1
        else:
            total_fail += 1

    print(f"\n{'='*60}")
    print(f"结果: {total_pass} PASS, {total_fail} FAIL")
    print(f"{'='*60}")

    return total_fail == 0


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
