#!/usr/bin/env python3
r"""
Word表格无损提取 -> LaTeX编译 完整Pipeline
用法: python run_pipeline.py <input.docx> [output_dir] [--no-pdf]

Pipeline:
  1. Word .docx -> all_tables_complete.json (extract_all_tables.py)
  2. JSON -> TikZ LaTeX .tex (tikz_table_gen.py)
  3. .tex -> .pdf (xelatex编译)
"""
import os, sys, json, argparse

sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def run_extraction(docx_path, output_dir):
    """Step 1: Word -> JSON"""
    from extract_all_tables import extract_tables
    result = extract_tables(docx_path)
    json_path = os.path.join(output_dir, 'all_tables_complete.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'[1/3] Extracted {result["total_tables"]} tables -> {json_path}')
    return json_path


def run_tikz_gen(json_path, output_dir, compile_pdf):
    """Step 2+3: JSON -> TikZ LaTeX -> PDF"""
    from tikz_table_gen import generate_tikz_document
    tikz_dir = os.path.join(output_dir, 'tikz_output')
    result = generate_tikz_document(json_path, tikz_dir, compile_pdf)
    print(f'[2/3] Generated TikZ LaTeX -> {result["tex_path"]}')
    if 'pdf_path' in result:
        print(f'[3/3] Compiled PDF -> {result["pdf_path"]}')
    elif 'compile_error' in result:
        print(f'[3/3] PDF compilation failed')
    else:
        print('[3/3] PDF compilation skipped')
    return result


def main():
    parser = argparse.ArgumentParser(description='Word表格无损提取 -> LaTeX编译 Pipeline')
    parser.add_argument('docx_path', help='输入的Word .docx文件路径')
    parser.add_argument('output_dir', nargs='?', default=None, help='输出目录(默认: docx同目录下)')
    parser.add_argument('--no-pdf', action='store_true', help='只生成.tex不编译PDF')
    args = parser.parse_args()

    docx_path = os.path.abspath(args.docx_path)
    if not os.path.exists(docx_path):
        print(f'Error: {docx_path} not found')
        sys.exit(1)

    output_dir = args.output_dir or os.path.join(os.path.dirname(docx_path), 'table_output')
    os.makedirs(output_dir, exist_ok=True)

    # 将skill目录加入path以便import
    if SCRIPT_DIR not in sys.path:
        sys.path.insert(0, SCRIPT_DIR)

    print(f'Input: {docx_path}')
    print(f'Output: {output_dir}')
    print()

    # Step 1: Word -> JSON
    json_path = run_extraction(docx_path, output_dir)

    # Step 2+3: JSON -> TikZ LaTeX -> PDF
    result = run_tikz_gen(json_path, output_dir, not args.no_pdf)

    print()
    print('Pipeline complete!')
    print(f'  JSON: {json_path}')
    print(f'  TeX:  {result["tex_path"]}')
    if 'pdf_path' in result:
        print(f'  PDF:  {result["pdf_path"]}')


if __name__ == '__main__':
    main()