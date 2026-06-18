"""Word 文档文本无损提取器

提取 Word (.docx) 中所有文本信息（段落、标题、格式、公式），不含表格和图片。
输出 JSON 结构，保留完整的格式信息用于后续 LaTeX 转换。
"""

import json
import sys
import io
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

# 确保 UTF-8 输出
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# OMML→LaTeX 引入
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'omml-to-latex'))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'citation-extract'))
from omml_to_latex import omml_to_latex, GREEK_MAP
from word_link_citations import make_citation_resolver

# 内部模块
from _semantic_classifier import classify_semantic_type
from _paragraph_extractor import extract_paragraph, _extract_run


# 颜色 → LaTeX 命令映射
COLOR_MAP = {
    'EE0000': 'cite',      # 红色 = 引用编号
    '000000': None,         # 黑色 = 默认
    'FF0000': '\\textcolor{red}',
    '0000FF': '\\textcolor{blue}',
    '008000': '\\textcolor{green}',
}


def heading_level(style_name):
    """从样式名推断标题级别"""
    if not style_name:
        return None
    if 'Heading 1' in style_name:
        return 1
    if 'Heading 2' in style_name:
        return 2
    if 'Heading 3' in style_name:
        return 3
    if 'Heading 4' in style_name:
        return 4
    return None


def alignment_name(align):
    """对齐方式转字符串"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    if align == WD_ALIGN_PARAGRAPH.LEFT:
        return 'left'
    elif align == WD_ALIGN_PARAGRAPH.CENTER:
        return 'center'
    elif align == WD_ALIGN_PARAGRAPH.RIGHT:
        return 'right'
    elif align == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return 'justify'
    return None


def extract_docx_text(docx_path, output_path=None, bib_path=None):
    """从 Word 文档无损提取所有文本信息（不含表格和图片）

    Args:
        docx_path: Word 文档路径
        output_path: JSON 输出路径（可选）

    Returns:
        dict: {
            'source': str,
            'total_paragraphs': int,
            'headings': list[dict],
            'paragraphs': list[dict],
            'statistics': dict,
        }
    """
    doc = Document(docx_path)
    citation_resolver = make_citation_resolver(bib_path)

    headings = []
    paragraphs = []
    cite_count = 0
    formula_count = 0
    bold_count = 0
    italic_count = 0

    for pi, para in enumerate(doc.paragraphs):
        # 跳过表格中的段落（python-docx 会把表格段落也包含在 paragraphs 中）
        parent = para._element.getparent()
        if parent is not None and parent.tag.split('}')[-1] == 'tc':
            continue

        info = extract_paragraph(
            para, pi, heading_level, alignment_name, omml_to_latex, COLOR_MAP,
            citation_resolver=citation_resolver,
        )
        paragraphs.append(info)

        if info['heading_level']:
            headings.append({
                'para_index': pi,
                'level': info['heading_level'],
                'text': info['text'],
                'style': info['style'],
                'latex': info['latex'],
            })

        for run in info['runs']:
            if run.get('is_cite'):
                cite_count += 1
            if run.get('type') == 'citation_link':
                cite_count += 1
            if run.get('type') == 'formula':
                formula_count += 1
            if run.get('bold'):
                bold_count += 1
            if run.get('italic'):
                italic_count += 1

    result = {
        'source': str(docx_path),
        'total_paragraphs': len(paragraphs),
        'headings': headings,
        'paragraphs': paragraphs,
        'statistics': {
            'headings': len(headings),
            'citations': cite_count,
            'formulas': formula_count,
            'bold_runs': bold_count,
            'italic_runs': italic_count,
        },
    }

    # 语义分类后处理（需要全局上下文）
    doc_meta = {'title_found': False, 'author_done': False, 'in_abstract': False}
    # 预推断正文字号：扫描所有非标题段落的主导字号，取众数
    from collections import Counter
    body_sizes = []
    for p in paragraphs:
        if p.get('heading_level') is not None:
            continue
        runs = p.get('runs', [])
        sizes = [r.get('size_pt') for r in runs
                 if isinstance(r, dict) and r.get('size_pt') is not None]
        if sizes and len(p.get('text', '').strip()) > 30:
            body_sizes.append(Counter(sizes).most_common(1)[0][0])
    if body_sizes:
        doc_meta['_body_font_size'] = Counter(body_sizes).most_common(1)[0][0]
    total = len(paragraphs)
    for i, p in enumerate(paragraphs):
        prev_type = paragraphs[i-1].get('semantic_type') if i > 0 else None
        p['semantic_type'] = classify_semantic_type(p, i, total, prev_type, doc_meta)

    # 标题级别修正：Word 中 Heading 2 但文本以单个数字开头 → 实为一级章节
    import re
    _RE_TOP_SECTION = re.compile(r'^\d+[\s、．.]|^\d+[^\d\s\.\、]')
    for p in paragraphs:
        if p.get('heading_level') == 2 and _RE_TOP_SECTION.match(p['text']):
            # 检查不是"2.1"这种带小数点的二级标题
            if not re.match(r'^\d+\.\d+', p['text']):
                p['heading_level'] = 1

    # 更新 headings 列表
    result['headings'] = []
    for p in paragraphs:
        if p.get('heading_level'):
            result['headings'].append({
                'para_index': p['para_index'],
                'level': p['heading_level'],
                'text': p['text'],
                'style': p['style'],
                'latex': p['latex'],
            })

    # 语义分类统计
    type_counts = {}
    for p in paragraphs:
        t = p.get('semantic_type', 'unknown')
        type_counts[t] = type_counts.get(t, 0) + 1
    result['statistics']['semantic_types'] = type_counts

    if output_path:
        Path(output_path).write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )

    return result


def generate_summary(result):
    """生成可读的提取摘要"""
    lines = []
    lines.append(f"来源: {result['source']}")
    lines.append(f"段落总数: {result['total_paragraphs']}")
    lines.append(f"统计: {result['statistics']}")
    lines.append("")
    lines.append("=== 标题结构 ===")
    for h in result['headings']:
        prefix = '#' * h['level']
        lines.append(f"  {prefix} {h['text']} (段落{h['para_index']})")
    lines.append("")
    lines.append("=== 段落格式概览 ===")
    for p in result['paragraphs'][:20]:
        level_str = f"H{p['heading_level']}" if p['heading_level'] else ''
        align_str = f", align={p['alignment']}" if p['alignment'] else ''
        indent_str = f", indent={p['first_line_indent_pt']}pt" if p['first_line_indent_pt'] else ''
        formula_str = ', has_formula' if p['has_formula'] else ''
        lines.append(f"  段落{p['para_index']}: {level_str}{align_str}{indent_str}{formula_str}")
        lines.append(f"    text: \"{p['text'][:40]}\"")
        if p['latex'] != p['text']:
            lines.append(f"    latex: \"{p['latex'][:40]}\"")
    lines.append("  ... (省略后续段落)")

    return '\n'.join(lines)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python text_extract.py <docx路径> [输出目录]")
        print("无损提取 Word 文档文本信息（标题/格式/公式/引用），不含表格图片")
        sys.exit(1)

    docx_path = sys.argv[1]
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path('.')
    bib_path = sys.argv[3] if len(sys.argv) > 3 else None
    out_dir.mkdir(parents=True, exist_ok=True)

    # 提取
    result = extract_docx_text(docx_path, out_dir / 'text_extract.json', bib_path=bib_path)
    print(f"提取完成: {result['total_paragraphs']} 个段落")
    print(f"统计: {result['statistics']}")

    # 保存摘要
    summary = generate_summary(result)
    Path(out_dir / 'text_extract_summary.txt').write_text(summary, encoding='utf-8')
    print(f"摘要: {out_dir / 'text_extract_summary.txt'}")
    print(f"JSON: {out_dir / 'text_extract.json'}")
