#!/usr/bin/env python3
"""全局 Orchestrator - Word → LaTeX → Word 完整转换管道

整合所有 skill 的调用顺序，添加审查机制和错误处理。

调用顺序：
  Step 1: 提取文本 (text-extract)
  Step 2: 提取图片 (document-extract)
  Step 3: 图片顺序匹配 (image-rename) [可选]
  Step 4: 提取表格 (table-lossless-extract)
  Step 5: 提取引用 (citation-extract)
  Step 6: 提取期刊模板 (journal-template-extract)
  Step 7: 整合生成 .tex (assemble_tex)
  Step 8: LaTeX 编译 (compile_tex)
  Step 9: 转换为 Word (tex-to-word) [可选]

用法:
  python orchestrator.py <docx> <模板目录> <bib> <期刊名> [输出目录] [选项]

选项:
  --original-images <dir>  原始高清图片目录（启用图片顺序匹配）
  --no-pdf                 不编译PDF
  --no-word                不转换为Word
  --no-verify              不生成确认报告
  --continue-on-error      审查失败时继续执行（默认中止）
"""

import os
import sys
import json
import shutil
import re
import subprocess
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

SKILL_DIR = Path(__file__).resolve().parent

# 添加所有 skill 到 sys.path
sys.path.insert(0, str(SKILL_DIR / 'shared'))
sys.path.insert(0, str(SKILL_DIR / 'text-extract'))
sys.path.insert(0, str(SKILL_DIR / 'omml-to-latex'))
sys.path.insert(0, str(SKILL_DIR / 'document-extract'))
sys.path.insert(0, str(SKILL_DIR / 'table-lossless-extract'))
sys.path.insert(0, str(SKILL_DIR / 'citation-extract'))
sys.path.insert(0, str(SKILL_DIR / 'journal-template-extract'))
sys.path.insert(0, str(SKILL_DIR / 'image-rename'))
sys.path.insert(0, str(SKILL_DIR / 'tex-to-word'))
sys.path.insert(0, str(SKILL_DIR / 'convert-latex'))
sys.path.insert(0, str(SKILL_DIR / 'template-extract-lite'))

# 导入各 skill 模块
from text_extract import extract_docx_text
from document_extract import extract_all_images_with_position
from extract_all_tables import extract_tables
from extract_citations import extract_citations
from cross_ref_builder import insert_bib_cross_references
from extract_template import generate_latex_file
from verify_extract import generate_verify_report, collect_chem_items
from image_order_match import match_images, print_mapping_table
from tex_to_word import tex_to_word
from template_spec_extract import extract_template_spec

# 从 convert_direct.py 导入核心函数
from convert_direct import (
    assemble_tex, copy_support_files, compile_tex,
    REF_KEYWORDS, clean_preamble, extract_skeleton_commands, postprocess_tex
)


def _tex_to_word_cli(tex_path, output_path, bib_path, config_mode=None,
                     pdf_float_wrap=False, pdf_float_reflow=False):
    """Run tex-to-word through its standalone entrypoint for isolation."""
    script = SKILL_DIR / 'tex-to-word' / 'tex_to_word.py'
    cmd = [
        sys.executable,
        str(script),
        str(tex_path),
        '-o',
        str(output_path),
        '--bib',
        str(bib_path),
    ]
    if config_mode:
        cmd.extend(['--config-mode', str(config_mode)])
    if pdf_float_wrap:
        cmd.append('--pdf-float-wrap')
    if pdf_float_reflow:
        cmd.append('--pdf-float-reflow')
    try:
        proc = subprocess.run(
            cmd, capture_output=True, text=True, encoding='utf-8',
            errors='replace', timeout=600)
    except subprocess.TimeoutExpired:
        raise RuntimeError('tex-to-word CLI 超时（>600s）')
    if proc.stdout:
        print(proc.stdout, end='')
    if proc.stderr:
        print(proc.stderr, end='', file=sys.stderr)
    if proc.returncode != 0:
        raise RuntimeError(f'tex-to-word CLI failed: {proc.returncode}')
    return str(output_path) if Path(output_path).exists() else None


def _convert_tex_to_word(tex_path, output_path, bib_path, config_mode=None,
                         pdf_float_wrap=False, pdf_float_reflow=False):
    """Convert TeX to Word, isolating Pandoc discovery when needed."""
    try:
        return tex_to_word(
            tex_path,
            output_path=str(output_path),
            bib_file=str(bib_path),
            config_mode=config_mode,
            use_pdf_float_wrap=pdf_float_wrap,
            use_pdf_float_reflow=pdf_float_reflow,
        )
    except RuntimeError as exc:
        if 'Pandoc executable not found' not in str(exc):
            raise
        print('  [tex-to-word] imported converter could not find Pandoc; retrying standalone entrypoint')
        return _tex_to_word_cli(
            tex_path, output_path, bib_path, config_mode,
            pdf_float_wrap, pdf_float_reflow)


def _effective_pdf_float_wrap(pdf_float_wrap, config_mode):
    """Default final-mode Word conversion to PDF-guided square-wrap floats."""
    if pdf_float_wrap is not None:
        return bool(pdf_float_wrap)
    return str(config_mode or '').lower() == 'final'


class ReviewWarning:
    """审查警告"""
    def __init__(self, step_name, message, severity='warning'):
        self.step_name = step_name
        self.message = message
        self.severity = severity  # 'error', 'warning', 'info'
        self.timestamp = datetime.now()

    def __str__(self):
        return f"[{self.severity.upper()}] {self.step_name}: {self.message}"


class OrchestratorResult:
    """管道执行结果"""
    def __init__(self):
        self.success = True
        self.warnings = []
        self.outputs = {}
        self.tex_path = None
        self.pdf_path = None
        self.word_path = None
        self.visual_word_path = None

    def add_warning(self, step_name, message, severity='warning'):
        self.warnings.append(ReviewWarning(step_name, message, severity))
        if severity == 'error':
            self.success = False

    def summary(self):
        lines = ['=' * 60, '执行结果摘要', '=' * 60]
        lines.append(f"状态: {'成功' if self.success else '失败'}")
        lines.append(f"警告数: {len(self.warnings)}")

        if self.tex_path:
            lines.append(f"TeX: {self.tex_path}")
        if self.pdf_path:
            lines.append(f"PDF: {self.pdf_path}")
        if self.word_path:
            lines.append(f"Word: {self.word_path}")
        if self.visual_word_path:
            lines.append(f"Visual Word: {self.visual_word_path}")

        if self.warnings:
            lines.append('\n警告列表:')
            for w in self.warnings:
                lines.append(f"  {w}")

        lines.append('=' * 60)
        return '\n'.join(lines)


def review_step(step_name, result, expected_keys=None, min_count=None):
    """审查步骤结果

    Args:
        step_name: 步骤名称
        result: 步骤返回结果
        expected_keys: 期望的键列表
        min_count: 期望的最小条目数

    Returns:
        list of ReviewWarning
    """
    warnings = []

    if result is None:
        warnings.append(ReviewWarning(step_name, '步骤返回 None', 'error'))
        return warnings

    if expected_keys:
        for key in expected_keys:
            if key not in result:
                warnings.append(ReviewWarning(step_name, f'缺少键: {key}', 'warning'))

    if min_count is not None:
        if isinstance(result, list) and len(result) < min_count:
            warnings.append(ReviewWarning(step_name, f'条目数 {len(result)} < 期望 {min_count}', 'warning'))
        elif isinstance(result, dict):
            count_key = 'paragraphs' if 'paragraphs' in result else 'tables' if 'tables' in result else None
            if count_key and len(result.get(count_key, [])) < min_count:
                warnings.append(ReviewWarning(step_name, f'{count_key}条目数不足', 'warning'))

    return warnings


def run_image_matching(extracted_dir, original_dir, output_dir):
    """执行图片顺序匹配

    Args:
        extracted_dir: 从Word提取的图片目录
        original_dir: 原始高清图片目录
        output_dir: 输出目录

    Returns:
        list of (fig_name, original_name, match_type) 或 None
    """
    if not Path(original_dir).exists():
        print(f'  警告: 原始图片目录不存在: {original_dir}')
        return None

    if not Path(extracted_dir).exists():
        print(f'  警告: 提取图片目录不存在: {extracted_dir}')
        return None

    print(f'  提取图片目录: {extracted_dir}')
    print(f'  原始图片目录: {original_dir}')

    matches = match_images(extracted_dir, original_dir)
    if matches:
        print_mapping_table(matches, original_dir)

        # 保存映射结果
        mapping_path = Path(output_dir) / 'image_mapping.json'
        mapping_data = [
            {'fig': fig, 'original': orig, 'match_type': mtype}
            for fig, orig, mtype in matches
        ]
        with open(mapping_path, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        print(f'  映射已保存: {mapping_path}')

        # 复制原始高清图片到 fig 目录（重命名为 fig1.ext 等）
        fig_dir = Path(output_dir) / 'fig'
        fig_dir.mkdir(parents=True, exist_ok=True)

        for fig_name, orig_name, _ in matches:
            if orig_name:
                src = Path(original_dir) / orig_name
                dst = fig_dir / fig_name
                if src.exists() and not dst.exists():
                    shutil.copy2(str(src), str(dst))
                    print(f'  复制高清图片: {orig_name} -> {fig_name}')

    return matches


def run_pipeline(
    docx_path,
    template_dir,
    bib_path,
    journal,
    output_dir=None,
    compile_pdf=True,
    convert_word=True,
    verify=True,
    continue_on_error=False,
    original_images_dir=None,
    config_mode=None,
    pdf_float_wrap=None,
    pdf_float_reflow=False
):
    """完整转换 Pipeline

    Args:
        docx_path: Word文档路径
        template_dir: 期刊模板目录
        bib_path: BibTeX文件路径
        journal: 期刊缩写
        output_dir: 输出目录
        compile_pdf: 是否编译PDF
        convert_word: 是否转换为Word
        verify: 是否生成确认报告
        continue_on_error: 审查失败时是否继续
        original_images_dir: 原始高清图片目录
        config_mode: 模板配置模式 ('classic', 'manuscript', 'final', 'discussions')

    Returns:
        OrchestratorResult
    """
    result = OrchestratorResult()

    docx_path = Path(docx_path).resolve()
    template_dir = Path(template_dir).resolve()
    bib_path = Path(bib_path).resolve()
    output_dir = Path(output_dir or docx_path.parent / 'convert_output').resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f'\n{"=" * 60}')
    print('Word → LaTeX → Word 完整转换管道')
    print(f'{"=" * 60}')
    print(f'输入: {docx_path}')
    print(f'模板: {template_dir}')
    print(f'BibTeX: {bib_path}')
    print(f'期刊: {journal}')
    print(f'输出: {output_dir}')
    print(f'编译PDF: {compile_pdf}')
    print(f'转Word: {convert_word}')
    print(f'审查模式: {"继续" if continue_on_error else "中止"}')
    if original_images_dir:
        print(f'原始图片: {original_images_dir}')
    print(f'{"=" * 60}\n')

    # ========== Step 1: 提取文本 ==========
    print('[1/9] 提取文本...')
    try:
        text_result = extract_docx_text(
            str(docx_path), str(output_dir / 'text_extract.json'), bib_path=str(bib_path))
        result.outputs['text'] = text_result

        # 审查
        warnings = review_step('text-extract', text_result, expected_keys=['paragraphs'], min_count=1)
        result.warnings.extend(warnings)

        if verify:
            report = generate_verify_report('text', text_result['paragraphs'], str(docx_path))
            (output_dir / 'verify_text_report.txt').write_text(report, encoding='utf-8')
            print(f'  确认报告: verify_text_report.txt')

        if any(w.severity == 'error' for w in warnings) and not continue_on_error:
            print('  [中止] 文本提取失败')
            return result
    except Exception as e:
        result.add_warning('text-extract', f'异常: {e}', 'error')
        if not continue_on_error:
            return result

    # ========== Step 2: 提取图片 ==========
    print('[2/9] 提取图片...')
    try:
        image_result = extract_all_images_with_position(str(docx_path), str(output_dir))
        result.outputs['image'] = image_result

        if verify:
            report = generate_verify_report('image', image_result, str(docx_path))
            (output_dir / 'verify_image_report.txt').write_text(report, encoding='utf-8')
            print(f'  确认报告: verify_image_report.txt')
    except Exception as e:
        result.add_warning('document-extract', f'异常: {e}', 'error')
        if not continue_on_error:
            return result

    # ========== Step 3: 图片顺序匹配（可选） ==========
    print('[3/9] 图片顺序匹配...')
    if original_images_dir:
        extracted_fig_dir = output_dir / 'fig'
        matches = run_image_matching(str(extracted_fig_dir), original_images_dir, str(output_dir))
        if matches:
            result.outputs['image_matches'] = matches
            print(f'  匹配成功: {len(matches)} 张图片')
        else:
            result.add_warning('image-rename', '图片匹配失败或无匹配', 'warning')
    else:
        print('  跳过（未指定原始图片目录）')

    # ========== Step 4: 提取表格 ==========
    print('[4/9] 提取表格...')
    try:
        table_result = extract_tables(str(docx_path))
        result.outputs['table'] = table_result
        with open(output_dir / 'all_tables_complete.json', 'w', encoding='utf-8') as f:
            json.dump(table_result, f, ensure_ascii=False, indent=2)

        if verify:
            report = generate_verify_report('table', table_result, str(docx_path))
            (output_dir / 'verify_table_report.txt').write_text(report, encoding='utf-8')
            print(f'  确认报告: verify_table_report.txt')
    except Exception as e:
        result.add_warning('table-lossless-extract', f'异常: {e}', 'error')
        if not continue_on_error:
            return result

    # ========== Step 5: 提取引用 ==========
    print('[5/9] 提取引用...')
    try:
        cite_result = extract_citations(str(docx_path), str(bib_path))
        result.outputs['citation'] = cite_result
        with open(output_dir / 'citations.json', 'w', encoding='utf-8') as f:
            json.dump(cite_result, f, ensure_ascii=False, indent=2)
    except Exception as e:
        result.add_warning('citation-extract', f'异常: {e}', 'warning')

    # ========== Step 6: 提取期刊模板 ==========
    print('[6/9] 提取期刊模板...')
    try:
        template_result = extract_template_spec(str(template_dir), journal, str(output_dir), config_mode=config_mode)
        result.outputs['template'] = template_result
    except Exception as e:
        result.add_warning('journal-template-extract', f'异常: {e}', 'error')
        if not continue_on_error:
            return result

    # ========== Step 7: 整合生成 .tex ==========
    print('[7/9] 整合生成 .tex...')
    try:
        required = ('text', 'image', 'table', 'template')
        missing = [key for key in required if key not in result.outputs]
        if missing:
            raise RuntimeError(f'缺少前置输出: {", ".join(missing)}')

        doc_options = result.outputs['template'].get('doc_options', [])
        tex_result = assemble_tex(
            result.outputs['text'],
            result.outputs['image'],
            result.outputs['table'],
            result.outputs['template'],
            bib_path,
            output_dir,
            str(docx_path),
            doc_options=doc_options
        )
        if isinstance(tex_result, tuple):
            tex_path, skeleton_info = tex_result
        else:
            tex_path, skeleton_info = tex_result, None
        result.tex_path = str(tex_path)
        result.outputs['tex_path'] = tex_path
        result.outputs['skeleton_info'] = skeleton_info

        if verify:
            tex_content = tex_path.read_text(encoding='utf-8')
            chem_items = collect_chem_items(tex_content)
            report = generate_verify_report('chem', chem_items, str(docx_path))
            (output_dir / 'verify_chem_report.txt').write_text(report, encoding='utf-8')
            print(f'  确认报告: verify_chem_report.txt')
    except Exception as e:
        result.add_warning('assemble_tex', f'异常: {e}', 'error')
        if not continue_on_error:
            return result

    # 复制支撑文件
    if result.tex_path and 'template' in result.outputs:
        copy_support_files(
            result.outputs['template'],
            bib_path,
            output_dir,
            str(docx_path),
            skeleton_info=result.outputs.get('skeleton_info')
        )

    # ========== Step 8: LaTeX 编译 ==========
    if compile_pdf and result.tex_path:
        print('[8/9] LaTeX 编译...')
        try:
            cls_name = (result.outputs.get('template') or {}).get('document_class', {}).get('class_name')
            pdf_path = compile_tex(tex_path, output_dir, cls_name=cls_name)
            if pdf_path:
                result.pdf_path = pdf_path
                print(f'  PDF: {pdf_path}')
            else:
                result.add_warning('compile_tex', '编译失败', 'error')
        except Exception as e:
            result.add_warning('compile_tex', f'异常: {e}', 'error')
    else:
        print('[8/9] LaTeX 编译... 跳过')

    # ========== Step 9: 转换为 Word ==========
    if convert_word and result.tex_path:
        print('[9/9] 转换为 Word...')
        try:
            effective_pdf_float_wrap = _effective_pdf_float_wrap(pdf_float_wrap, config_mode)
            effective_pdf_float_reflow = bool(pdf_float_reflow)
            if effective_pdf_float_wrap:
                print('  [tex-to-word] PDF float wrap enabled')
            if effective_pdf_float_reflow:
                print('  [tex-to-word] PDF float reflow enabled')
            word_path = _convert_tex_to_word(
                result.tex_path,
                output_path=str(output_dir / f'{journal}_converted.docx'),
                bib_path=str(bib_path),
                config_mode=config_mode,
                pdf_float_wrap=effective_pdf_float_wrap,
                pdf_float_reflow=effective_pdf_float_reflow
            )
            if word_path:
                result.word_path = word_path
                print(f'  Word: {word_path}')
            else:
                result.add_warning('tex-to-word', '转换失败', 'error')
        except Exception as e:
            result.add_warning('tex-to-word', f'异常: {e}', 'error')
    else:
        print('[9/9] 转换为 Word... 跳过')

    # ========== 完成 ==========
    print(f'\n{result.summary()}')

    # 保存执行日志
    log_path = output_dir / 'orchestrator_log.json'
    log_data = {
        'success': result.success,
        'tex_path': result.tex_path,
        'pdf_path': result.pdf_path,
        'word_path': result.word_path,
        'visual_word_path': result.visual_word_path,
        'warnings': [
            {'step': w.step_name, 'message': w.message, 'severity': w.severity}
            for w in result.warnings
        ],
        'timestamp': datetime.now().isoformat(),
    }
    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump(log_data, f, ensure_ascii=False, indent=2)
    print(f'执行日志: {log_path}')

    return result


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description='Word → LaTeX → Word 完整转换管道',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 基本用法
  python orchestrator.py paper.docx template/ refs.bib acp

  # 启用图片顺序匹配
  python orchestrator.py paper.docx template/ refs.bib acp --original-images ./original_images

  # 不编译PDF，不转Word
  python orchestrator.py paper.docx template/ refs.bib acp --no-pdf --no-word

  # 审查失败时继续执行
  python orchestrator.py paper.docx template/ refs.bib acp --continue-on-error

  # 默认要求 Word 输出成功；若只需要 TeX/PDF，请显式使用 --no-word
        """
    )

    parser.add_argument('docx_path', help='Word文档路径')
    parser.add_argument('template_dir', help='期刊模板目录(含.cls/.cfg/.bst)')
    parser.add_argument('bib_path', help='references.bib路径')
    parser.add_argument('journal', help='期刊缩写(如acp)')
    parser.add_argument('output_dir', nargs='?', default=None, help='输出目录')

    parser.add_argument('--original-images', dest='original_images_dir',
                        help='原始高清图片目录（启用图片顺序匹配）')
    parser.add_argument('--no-pdf', action='store_true', help='不编译PDF')
    parser.add_argument('--no-word', action='store_true', help='不转换为Word')
    parser.add_argument('--no-verify', action='store_true', help='不生成确认报告')
    parser.add_argument('--continue-on-error', action='store_true',
                        help='审查失败时继续执行')
    parser.add_argument('--config-mode',
                        choices=['classic', 'manuscript', 'final', 'discussions'],
                        default=None,
                        help='模板配置模式')
    parser.add_argument('--pdf-float-reflow', action='store_true',
                        help='Word生成后渲染PDF并尝试前移跨栏图表组以减少大片空白')
    parser.add_argument('--pdf-float-wrap', dest='pdf_float_wrap', action='store_true',
                        default=None,
                        help='Word生成后渲染PDF并尝试把跨栏图表改为可编辑浮动环绕容器')
    parser.add_argument('--no-pdf-float-wrap', dest='pdf_float_wrap', action='store_false',
                        help='关闭final模式默认启用的PDF浮动环绕图表处理')

    args = parser.parse_args()

    run_pipeline(
        args.docx_path,
        args.template_dir,
        args.bib_path,
        args.journal,
        args.output_dir,
        compile_pdf=not args.no_pdf,
        convert_word=not args.no_word,
        verify=not args.no_verify,
        continue_on_error=args.continue_on_error,
        original_images_dir=args.original_images_dir,
        config_mode=args.config_mode,
        pdf_float_wrap=args.pdf_float_wrap,
        pdf_float_reflow=args.pdf_float_reflow
    )
