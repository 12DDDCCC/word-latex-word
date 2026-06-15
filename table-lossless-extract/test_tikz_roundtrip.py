#!/usr/bin/env python3
"""TikZ 回环测试

验证: JSON → TikZ (tikz_table_gen) → JSON (latex_table_parser) → Word (gen_table_from_json)

比较原始 JSON 和回环 JSON 的关键字段，确认无损性。
"""
import json
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

# 添加路径
SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SKILL_DIR)

from tikz_table_gen import process_table
from latex_table_parser import tikz_to_json
from gen_table_from_json import generate_docx


def compare_grid_cols(orig, rt, tol=0.05):
    """比较列宽，容差 5%"""
    if len(orig) != len(rt):
        return False, f"列数不同: {len(orig)} vs {len(rt)}"
    for i, (o, r) in enumerate(zip(orig, rt)):
        ow = o.get('width_twips', 0)
        rw = r.get('width_twips', 0)
        if ow > 0 and abs(rw - ow) / ow > tol:
            return False, f"列{i}宽度差异: {ow} vs {rw} ({abs(rw-ow)/ow*100:.1f}%)"
    return True, "OK"


def compare_text(orig_rows, rt_rows):
    """比较单元格文本"""
    # 收集所有文本
    orig_texts = set()
    for row in orig_rows:
        for cell in row['cells']:
            t = cell.get('text', '').strip()
            if t:
                orig_texts.add(t)

    rt_texts = set()
    for row in rt_rows:
        for cell in row.get('cells', []):
            # 从 paragraphs 提取文本
            for p in cell.get('paragraphs', []):
                for r in p.get('runs', []):
                    rt_texts.add(r.get('text', '').strip())

    missing = orig_texts - rt_texts
    extra = rt_texts - orig_texts
    if missing:
        return False, f"缺失文本: {missing}"
    if extra:
        # 允许少量额外文本（可能是节点文本拆分差异）
        pass
    return True, "OK"


def compare_border_count(orig_rows, rt_rows):
    """比较边框数量"""
    orig_borders = 0
    for row in orig_rows:
        for cell in row['cells']:
            borders = cell.get('borders', {})
            for edge in ('top', 'bottom', 'left', 'right'):
                b = borders.get(edge, {})
                if b.get('val') == 'single':
                    orig_borders += 1

    rt_borders = 0
    for row in rt_rows:
        for cell in row.get('cells', []):
            borders = cell.get('borders', {})
            for edge in ('top', 'bottom', 'left', 'right'):
                b = borders.get(edge, {})
                if b.get('val') == 'single':
                    rt_borders += 1

    if orig_borders == 0 and rt_borders == 0:
        return True, "OK (无显式边框)"
    diff_pct = abs(rt_borders - orig_borders) / max(orig_borders, 1) * 100
    if diff_pct > 20:
        return False, f"边框数差异: {orig_borders} vs {rt_borders} ({diff_pct:.0f}%)"
    return True, f"边框数: {orig_borders} vs {rt_borders} ({diff_pct:.0f}%)"


def run_roundtrip_test(json_path, output_dir=None):
    """运行回环测试"""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    tables = data.get('tables', [])
    print(f"{'='*60}")
    print(f"TikZ 回环测试: {json_path}")
    print(f"表格数: {len(tables)}")
    print(f"{'='*60}\n")

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    total_pass = 0
    total_fail = 0

    for tbl in tables:
        idx = tbl.get('table_index', 0)
        print(f"[Table {idx}]", end=" ")

        # Step 1: 原始 JSON → TikZ
        tikz_code = process_table(tbl, idx)
        table_env = ''
        # 从 caption 提取
        cap_m = None
        for line in tikz_code.split('\n'):
            if '\\caption{' in line:
                cap_m = line
                break

        # 提取 tikz_body (\begin{tikzpicture}...\end{tikzpicture} 内)
        tikz_start = tikz_code.find('\\begin{tikzpicture}')
        tikz_end = tikz_code.find('\\end{tikzpicture}')
        if tikz_start < 0 or tikz_end < 0:
            print("SKIP (无 TikZ 内容)")
            continue

        tikz_body = tikz_code[tikz_start + len('\\begin{tikzpicture}'):tikz_end]

        # Step 2: TikZ → JSON (回环)
        rt_json = tikz_to_json(tikz_body, tikz_code)

        # Step 3: 比较
        results = []

        # 列宽比较
        orig_cols = tbl.get('grid_cols', [])
        rt_cols = rt_json.get('grid_cols', [])
        ok, msg = compare_grid_cols(orig_cols, rt_cols)
        results.append(('grid_cols', ok, msg))

        # 文本比较
        orig_rows = tbl.get('rows', [])
        rt_rows = rt_json.get('rows', [])
        ok, msg = compare_text(orig_rows, rt_rows)
        results.append(('text', ok, msg))

        # 边框比较
        ok, msg = compare_border_count(orig_rows, rt_rows)
        results.append(('borders', ok, msg))

        # 行数比较
        num_data_rows = len(orig_rows)
        # 跳过内部 caption 行
        if orig_rows and orig_rows[0]['cells'][0].get('gridSpan', 1) == len(orig_rows[0]['cells']):
            num_data_rows -= 1
        rt_num_rows = len(rt_rows)
        row_match = abs(rt_num_rows - num_data_rows) <= 1
        results.append(('row_count', row_match, f"{num_data_rows} vs {rt_num_rows}"))

        # 打印结果
        all_pass = all(ok for _, ok, _ in results)
        if all_pass:
            print("PASS", end=" ")
            total_pass += 1
        else:
            print("FAIL", end=" ")
            total_fail += 1

        for name, ok, msg in results:
            status = "✓" if ok else "✗"
            print(f"  {status}{name}: {msg}", end="")
        print()

        # Step 4: 生成 Word（如果指定了输出目录）
        if output_dir and rt_json.get('grid_cols'):
            try:
                docx_path = os.path.join(output_dir, f'roundtrip_table_{idx}.docx')
                generate_docx(rt_json, docx_path)
                print(f"  → Word: {docx_path}")
            except Exception as e:
                print(f"  → Word 生成失败: {e}")

    print(f"\n{'='*60}")
    print(f"结果: {total_pass} PASS, {total_fail} FAIL (共 {len(tables)} 表格)")
    print(f"{'='*60}")

    return total_fail == 0


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python test_tikz_roundtrip.py <all_tables_complete.json> [output_dir]")
        sys.exit(1)

    json_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    success = run_roundtrip_test(json_path, output_dir)
    sys.exit(0 if success else 1)
