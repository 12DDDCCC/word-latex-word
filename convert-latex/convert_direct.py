#!/usr/bin/env python3
"""Word→LaTeX 无损转换整合脚本

调用6个skill（text-extract, omml-to-latex, document-extract, table-lossless-extract,
citation-extract, journal-template-extract），将Word文档转换为可直接编译的LaTeX文件。

核心策略：
  - text-extract 为主索引（段落顺序、文本内容）
  - 图片/表格按上下文文本匹配到对应段落位置
  - 参考文献段落自动删除，由 bib+bibstyle 替代
  - 引用key用数字直接匹配 bib 文件

用法: python convert_direct.py <docx> <模板目录> <bib> <期刊名> [输出目录] [--no-pdf]
"""
import sys, json
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# ─── sys.path: 其他skill模块 + shared ───
SKILL_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(SKILL_DIR))  # shared模块
sys.path.insert(0, str(SKILL_DIR / 'text-extract'))
sys.path.insert(0, str(SKILL_DIR / 'omml-to-latex'))
sys.path.insert(0, str(SKILL_DIR / 'document-extract'))
sys.path.insert(0, str(SKILL_DIR / 'table-lossless-extract'))
sys.path.insert(0, str(SKILL_DIR / 'citation-extract'))
sys.path.insert(0, str(SKILL_DIR / 'journal-template-extract'))
sys.path.insert(0, str(SKILL_DIR / 'template-extract-lite'))

# ─── 其他skill模块导入 ───
from text_extract import extract_docx_text
from document_extract import extract_all_images_with_position
from extract_all_tables import extract_tables
from extract_citations import extract_citations
from verify_extract import generate_verify_report, collect_chem_items

# ─── 包内子模块 re-export ───
from spec_adapter import SpecAdapter
from numbering_system import (
    detect_template_numbering_mode, apply_numbering_system,
    detect_source_numbering_mode, convert_numbering_references,
    renumber_sectioned_to_simple, generate_equation_label,
    REF_KEYWORDS,
)
from skeleton_builder import (
    parse_template_skeleton, clean_preamble, extract_skeleton_commands,
    _build_skeleton_from_spec, _build_preamble_from_spec, _default_skeleton_info,
)
from template_spec_extract import (
    extract_template_spec, _extend_keywords_from_spec,
    INTRO_KEYWORDS, CONCLUSION_KEYWORDS, APPENDIX_KEYWORDS, DECL_KEYWORD_MAP,
)
from assemble_tex import (
    assemble_tex, _insert_images_and_tables, build_image_map,
    _strip_heading_number, _find_reference_start, _extract_abstract_keywords,
)
from tex_postprocess import postprocess_tex, copy_support_files, compile_tex


# ─── Pipeline 入口 ─────────────────────────────────────────

def run_pipeline(docx_path, template_dir, bib_path, journal, output_dir, compile_pdf=True, verify=True, config_mode=None):
    """完整转换 Pipeline

    Args:
        verify: 是否生成确认报告（默认 True）
        config_mode: 模板配置模式名 ('classic', 'manuscript', 'final', 'discussions')
    """
    docx_path = Path(docx_path).resolve()
    template_dir = Path(template_dir).resolve()
    bib_path = Path(bib_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    print('[1/6] 提取文本...')
    text_result = extract_docx_text(
        str(docx_path), str(output_dir / 'text_extract.json'), bib_path=str(bib_path))

    # -- 确认步骤：文本提取 --
    if verify:
        print('\n' + '=' * 60)
        print('【文本提取确认】请同时查看 Word 原文，确认以下内容：')
        print('=' * 60)
        report = generate_verify_report('text', text_result['paragraphs'], str(docx_path))
        print(report)
        report_path = output_dir / 'verify_text_report.txt'
        report_path.write_text(report, encoding='utf-8')
        print(f'\n确认报告已保存: {report_path}')
        print('如发现问题，请修复后再继续。')
        print('=' * 60 + '\n')

    print('[2/6] 提取图片...')
    image_result = extract_all_images_with_position(str(docx_path), str(output_dir))

    # -- 确认步骤：图片提取 --
    if verify:
        print('\n' + '=' * 60)
        print('【图片提取确认】')
        print('=' * 60)
        report = generate_verify_report('image', image_result, str(docx_path))
        print(report)
        report_path = output_dir / 'verify_image_report.txt'
        report_path.write_text(report, encoding='utf-8')
        print(f'确认报告已保存: {report_path}')
        print('=' * 60 + '\n')

    print('[3/6] 提取表格...')
    table_result = extract_tables(str(docx_path))
    with open(output_dir / 'all_tables_complete.json', 'w', encoding='utf-8') as f:
        json.dump(table_result, f, ensure_ascii=False, indent=2)

    # -- 确认步骤：表格提取 --
    if verify:
        print('\n' + '=' * 60)
        print('【表格提取确认】')
        print('=' * 60)
        report = generate_verify_report('table', table_result, str(docx_path))
        print(report)
        report_path = output_dir / 'verify_table_report.txt'
        report_path.write_text(report, encoding='utf-8')
        print(f'确认报告已保存: {report_path}')
        print('=' * 60 + '\n')

    print('[4/6] 提取引用（验证用）...')
    cite_result = extract_citations(str(docx_path), str(bib_path))
    with open(output_dir / 'citations.json', 'w', encoding='utf-8') as f:
        json.dump(cite_result, f, ensure_ascii=False, indent=2)

    print('[5/6] 提取期刊模板规格...')
    template_result = extract_template_spec(str(template_dir), journal, str(output_dir), config_mode=config_mode)

    print('[6/6] 整合生成完整 .tex...')
    doc_options = template_result.get('doc_options', [])
    tex_path, skeleton_info = assemble_tex(text_result, image_result, table_result,
                            template_result, bib_path, output_dir, str(docx_path),
                            doc_options=doc_options)

    # -- 确认步骤：化学式下标 --
    if verify:
        tex_content = tex_path.read_text(encoding='utf-8')
        chem_items = collect_chem_items(tex_content)
        print('\n' + '=' * 60)
        print('【化学式下标确认】')
        print('=' * 60)
        report = generate_verify_report('chem', chem_items, str(docx_path))
        print(report)
        report_path = output_dir / 'verify_chem_report.txt'
        report_path.write_text(report, encoding='utf-8')
        print(f'确认报告已保存: {report_path}')
        print('=' * 60 + '\n')

    copy_support_files(template_result, bib_path, output_dir, str(docx_path), skeleton_info=skeleton_info)

    print(f'\n整合完成!')
    print(f'  TeX:  {tex_path}')
    print(f'  图片: {output_dir / "fig"}')
    print(f'  Bib:  {output_dir / "references.bib"}')

    if compile_pdf:
        print('\n编译验证...')
        compile_result = compile_tex(tex_path, output_dir)
        if compile_result:
            print(f'  PDF: {compile_result}')
        else:
            print('  编译失败，请手动检查')

    return {
        'tex_path': str(tex_path),
        'output_dir': str(output_dir),
    }


# ─── CLI 入口 ─────────────────────────────────────────────

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Word→LaTeX 无损转换整合')
    parser.add_argument('docx_path', help='Word文档路径')
    parser.add_argument('template_dir', help='期刊模板目录(含.cls/.cfg/.bst)')
    parser.add_argument('bib_path', help='references.bib路径')
    parser.add_argument('journal', help='期刊缩写(如acp)')
    parser.add_argument('output_dir', nargs='?', default=None, help='输出目录')
    parser.add_argument('--no-pdf', action='store_true', help='不编译PDF')
    parser.add_argument('--no-verify', action='store_true', help='不生成确认报告')
    parser.add_argument('--config-mode', default=None,
                        help='模板配置模式 (classic/manuscript/final/discussions)')
    args = parser.parse_args()

    output_dir = args.output_dir or str(Path(args.docx_path).parent / 'convert_output')
    run_pipeline(args.docx_path, args.template_dir, args.bib_path,
                 args.journal, output_dir, not args.no_pdf, not args.no_verify,
                 config_mode=args.config_mode)
