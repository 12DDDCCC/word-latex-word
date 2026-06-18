#!/usr/bin/env python3
"""提取确认模块 — 生成确认报告供人工审核

在转换流程的每个关键阶段生成确认报告，提示人工（Claude）同时查看
Word 原文和提取结果，确认一致性。

支持的确认阶段：
  - text: 文本提取（语义分类、标题/摘要/关键词/正文）
  - formula: 公式提取（OMML→LaTeX 转换正确性）
  - image: 图片提取（图例与图片对应）
  - table: 表格提取（表例与表格对应）
  - chem: 化学式下标（是否该加下标）

用法:
  from verify_extract import generate_verify_report
  report = generate_verify_report('text', extracted_data, docx_path)
  print(report)  # 输出确认报告，供人工审核
"""

import json, os, sys, re
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass


# ── 报告生成器 ──────────────────────────────────────────

def generate_verify_report(stage, data, docx_path=None, output_path=None):
    """生成确认报告，供人工审核

    Args:
        stage: 'text' | 'formula' | 'image' | 'table' | 'chem'
        data: 提取结果数据
        docx_path: 原始 Word 文档路径（用于提示人工查看原文）
        output_path: 报告输出路径（可选）

    Returns:
        str: 确认报告文本
    """
    generators = {
        'text': _generate_text_report,
        'formula': _generate_formula_report,
        'image': _generate_image_report,
        'table': _generate_table_report,
        'chem': _generate_chem_report,
    }

    if stage not in generators:
        return f'[错误] 未知阶段: {stage}'

    report_lines = []
    report_lines.append('=' * 60)
    report_lines.append(f'【{stage.upper()} 阶段确认报告】')
    report_lines.append('=' * 60)
    if docx_path:
        report_lines.append(f'原始文档: {docx_path}')
    report_lines.append('')
    report_lines.append('请同时查看 Word 原文和提取结果，确认以下内容是否一致：')
    report_lines.append('')

    # 调用对应阶段的报告生成器
    report_lines.extend(generators[stage](data))

    report_lines.append('')
    report_lines.append('=' * 60)
    report_lines.append('【确认要点】')
    report_lines.append('=' * 60)
    report_lines.extend(_get_checklist(stage))
    report_lines.append('')
    report_lines.append('如发现问题，请在转换流程中修复后再继续。')

    report = '\n'.join(report_lines)

    if output_path:
        Path(output_path).write_text(report, encoding='utf-8')

    return report


def _generate_text_report(paragraphs):
    """生成文本提取确认报告"""
    lines = []

    # 统计各语义类型
    type_counts = {}
    for p in paragraphs:
        t = p.get('semantic_type', 'unknown')
        type_counts[t] = type_counts.get(t, 0) + 1

    lines.append('【语义分类统计】')
    for t, cnt in sorted(type_counts.items()):
        lines.append(f'  {t}: {cnt} 个段落')
    lines.append('')

    # 列出关键语义类型的内容
    key_types = ['title', 'author', 'affiliation', 'abstract', 'keywords',
                 'heading', 'display_formula', 'figure_caption', 'table_caption', 'reference']

    for kt in key_types:
        items = [p for p in paragraphs if p.get('semantic_type') == kt]
        if not items:
            continue
        lines.append(f'【{kt}】({len(items)} 个)')
        for i, p in enumerate(items[:5]):  # 最多显示5个
            txt = p['text'][:80].replace('\n', ' ')
            pi = p.get('para_index', '?')
            hl = p.get('heading_level', '')
            extra = f' H{hl}' if hl else ''
            lines.append(f'  [{pi}{extra}] {txt}')
        if len(items) > 5:
            lines.append(f'  ... 还有 {len(items) - 5} 个')
        lines.append('')

    # 检查潜在问题
    issues = _check_text_issues(paragraphs)
    if issues:
        lines.append('【潜在问题】')
        for issue in issues:
            lines.append(f'  ⚠ {issue}')
        lines.append('')

    return lines


def _generate_formula_report(paragraphs):
    """生成公式提取确认报告"""
    lines = []

    formula_paras = [p for p in paragraphs if p.get('has_formula')]

    lines.append(f'【公式统计】共 {len(formula_paras)} 个段落含公式')
    lines.append('')

    for i, p in enumerate(formula_paras[:10]):  # 最多显示10个
        txt = p['text'][:60].replace('\n', ' ')
        latex = p.get('latex', '')[:100]
        pi = p.get('para_index', '?')
        lines.append(f'[{pi}] 原文: {txt}')
        lines.append(f'     LaTeX: {latex}')
        lines.append('')

    if len(formula_paras) > 10:
        lines.append(f'... 还有 {len(formula_paras) - 10} 个公式段落')
        lines.append('')

    return lines


def _generate_image_report(image_result):
    """生成图片提取确认报告"""
    lines = []

    images = image_result if isinstance(image_result, list) else image_result.get('images', [])

    lines.append(f'【图片统计】共 {len(images)} 张图片')
    lines.append('')

    for i, img in enumerate(images[:10]):
        img_file = img.get('image_file', '?')
        pi = img.get('para_index', '?')
        ctx = img.get('context', '')[:60]
        lines.append(f'[{i}] para_index={pi} | {img_file}')
        lines.append(f'     上下文: {ctx}')
        lines.append('')

    if len(images) > 10:
        lines.append(f'... 还有 {len(images) - 10} 张图片')
        lines.append('')

    return lines


def _generate_table_report(table_result):
    """生成表格提取确认报告"""
    lines = []

    tables = table_result.get('tables', [])

    lines.append(f'【表格统计】共 {len(tables)} 个表格')
    lines.append('')

    for i, tbl in enumerate(tables[:5]):
        rows = len(tbl.get('rows', []))
        cols = len(tbl.get('rows', [{}])[0].get('cells', [])) if tbl.get('rows') else 0
        lines.append(f'[{i}] {rows}行 × {cols}列')
        # 显示表头
        if tbl.get('rows'):
            header = tbl['rows'][0]
            header_text = ' | '.join(c.get('text', '')[:20] for c in header.get('cells', []))
            lines.append(f'     表头: {header_text[:80]}')
        lines.append('')

    if len(tables) > 5:
        lines.append(f'... 还有 {len(tables) - 5} 个表格')
        lines.append('')

    return lines


def _generate_chem_report(chem_items):
    """生成化学式下标确认报告"""
    lines = []

    lines.append(f'【化学式统计】共 {len(chem_items)} 项')
    lines.append('')

    # 已知排除项（不应加下标）
    _KNOWN_NON_CHEM = ['GCASv2', 'GOSAT2', 'OCO2', 'OCO3', 'MODIS', 'TCCON', 'MOPITT',
                       'CAMS', 'CMIP6', 'BEPS', 'FLUX']

    subscripted = []
    kept = []
    for item in chem_items:
        action = item.get('action', '?')
        if action == 'subscript':
            subscripted.append(item)
        else:
            kept.append(item)

    lines.append('【已加下标】')
    for item in subscripted[:10]:
        orig = item.get('original', '?')
        conv = item.get('converted', '?')
        lines.append(f'  {orig} → {conv}')
    if len(subscripted) > 10:
        lines.append(f'  ... 还有 {len(subscripted) - 10} 项')
    lines.append('')

    lines.append('【保持原样】')
    for item in kept[:10]:
        orig = item.get('original', '?')
        lines.append(f'  {orig}')
    if len(kept) > 10:
        lines.append(f'  ... 还有 {len(kept) - 10} 项')
    lines.append('')

    # 检查潜在问题
    issues = []
    for item in subscripted:
        orig = item.get('original', '')
        for exc in _KNOWN_NON_CHEM:
            if exc.lower() in orig.lower():
                issues.append(f'{orig} 被加下标，但可能是模型/卫星名')
                break

    if issues:
        lines.append('【潜在问题】')
        for issue in issues:
            lines.append(f'  ⚠ {issue}')
        lines.append('')

    return lines


def _check_text_issues(paragraphs):
    """检查文本提取的潜在问题"""
    issues = []

    # 检查是否有 title
    if not any(p.get('semantic_type') == 'title' for p in paragraphs):
        issues.append('未找到文章标题 (semantic_type=title)')

    # 检查是否有 abstract
    if not any(p.get('semantic_type') == 'abstract' for p in paragraphs):
        issues.append('未找到摘要 (semantic_type=abstract)')

    # 检查 figure_caption 是否以图/Figure 开头
    for p in paragraphs:
        st = p.get('semantic_type')
        txt = p.get('text', '')
        if st == 'figure_caption' and not re.match(r'^\s*(图|Figure|Fig\.?)', txt, re.IGNORECASE):
            issues.append(f'figure_caption 但文本不以图/Figure开头: {txt[:30]}')
            break

    # 检查 table_caption 是否以表/Table 开头
    for p in paragraphs:
        st = p.get('semantic_type')
        txt = p.get('text', '')
        if st == 'table_caption' and not re.match(r'^\s*(表|Table)', txt, re.IGNORECASE):
            issues.append(f'table_caption 但文本不以表/Table开头: {txt[:30]}')
            break

    return issues


def _get_checklist(stage):
    """获取各阶段的确认要点"""
    checklists = {
        'text': [
            '1. 文章标题是否正确识别（semantic_type=title）',
            '2. 作者姓名是否正确识别（semantic_type=author）',
            '3. 机构信息是否正确识别（semantic_type=affiliation）',
            '4. 摘要内容是否完整（semantic_type=abstract），Abstract后标点应为英文冒号":"',
            '5. 关键词是否正确识别（semantic_type=keywords）',
            '6. 章节标题编号是否剥离（如"1、研究背景"→"研究背景"，避免与LaTeX自动编号重复）',
            '7. 一级章节是否正确识别（如"3实验设置"应为H1而非H2）',
            '8. 图说明是否以"图/Figure"开头（semantic_type=figure_caption）',
            '9. 表说明是否以"表/Table"开头（semantic_type=table_caption）',
            '10. 正文段落是否正确分类（semantic_type=body）',
            '11. 独立公式段落是否识别为display_formula（不缩进，用equation环境）',
        ],
        'formula': [
            '1. 公式 LaTeX 转换是否正确',
            '2. 上下标是否正确',
            '3. 分数/积分/求和结构是否正确',
            '4. 希腊字母是否正确转换',
            '5. 独立公式是否用equation环境包裹（不应缩进）',
            '6. 公式编号是否正确提取',
        ],
        'image': [
            '1. 图片数量是否与 Word 一致',
            '2. 图片文件名是否正确',
            '3. 图片位置（para_index）是否与 Word 原文位置一致',
            '4. 图说明是否与图片对应（编号和内容）',
            '5. 图片插入顺序是否与 Word 原文一致',
        ],
        'table': [
            '1. 表格数量是否与 Word 一致',
            '2. 表格行列数是否正确',
            '3. 表格内容是否完整',
            '4. 表说明是否与表格对应（编号和内容）',
            '5. 表格插入位置是否与 Word 原文一致',
        ],
        'chem': [
            '1. CO₂、XCO₂ 等化学式是否正确加下标',
            '2. GCASv2、OCO-2 等模型/卫星名是否保持原样',
            '3. 没有原文不存在的下标被添加',
            '4. 原文无下标的文本不得被添加下标',
        ],
    }
    return checklists.get(stage, ['请确认提取结果与原文一致'])


def collect_chem_items(tex_content):
    """从 tex 内容中收集所有化学式相关项，供确认

    返回 list of {original, converted, action}
    """
    items = []
    # 已被转下标的
    for m in re.finditer(r'(\w+)\$_\{(\d+)\}$', tex_content):
        items.append({
            'original': m.group(1) + m.group(2),
            'converted': m.group(0),
            'action': 'subscript',
        })
    # OCO-2/OCO-3
    for m in re.finditer(r'OCO-(\d)', tex_content):
        items.append({
            'original': 'OCO' + m.group(1),
            'converted': m.group(0),
            'action': 'dash (satellite name)',
        })
    return items


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='提取确认报告生成器')
    parser.add_argument('stage', choices=['text', 'formula', 'image', 'table', 'chem'])
    parser.add_argument('json_file', help='提取结果 JSON 文件')
    parser.add_argument('--docx', help='原始 Word 文档路径')
    parser.add_argument('--output', '-o', help='报告输出路径')
    args = parser.parse_args()

    with open(args.json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    if args.stage == 'text':
        paragraphs = data.get('paragraphs', data) if isinstance(data, dict) else data
        result = generate_verify_report('text', paragraphs, args.docx, args.output)
    elif args.stage == 'chem':
        tex_path = Path(args.json_file)
        if tex_path.suffix == '.tex':
            items = collect_chem_items(tex_path.read_text(encoding='utf-8'))
        else:
            items = data if isinstance(data, list) else data.get('chem_items', [])
        result = generate_verify_report('chem', items, args.docx, args.output)
    else:
        result = generate_verify_report(args.stage, data, args.docx, args.output)

    print(result)
