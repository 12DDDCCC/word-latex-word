"""Caption检测工具 — 跨skill使用的图例/表例识别逻辑

2个skill中重复的caption检测逻辑合并到此模块：
- document-extract: _is_caption_by_font_style, _is_caption_paragraph, _find_caption_and_context, _find_legend_paragraphs
- table-lossless-extract: _is_table_caption_para, _find_table_legend_paragraphs

注意：此模块处理caption的"检测"（判断是否为caption），
而 shared/caption_utils.py 处理caption的"格式化"（normalize/clean/strip）。
两者职责清晰分离。
"""

import re
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH


# caption文本模式正则
_CAPTION_TEXT_RE = re.compile(
    r'^[图表]\s*[\d\.]+|^Figure|^Fig\.|^Table|^Tab\.|^Talbe',
    re.IGNORECASE
)


_FIGURE_CAPTION_RE = re.compile(
    r'^\s*(?:图|Figure|Fig\.?)\s*\d+(?:\.\d+)*',
    re.IGNORECASE,
)
_TABLE_CAPTION_RE = re.compile(
    r'^\s*(?:表|Table|Tab\.?|Talbe)\s*\d+(?:\.\d+)*',
    re.IGNORECASE,
)


def is_figure_caption_text(text):
    return bool(_FIGURE_CAPTION_RE.match(text or ""))


def is_table_caption_text(text):
    return bool(_TABLE_CAPTION_RE.match(text or ""))


def get_para_font_size_xml(para):
    """从XML层获取段落首个非空run的sz值(half-pt)，比python-docx高层API更可靠"""
    elem = para._element
    for r in elem.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        if rPr is not None:
            sz = rPr.find(qn('w:sz'))
            if sz is not None:
                val = sz.get(qn('w:val'))
                if val:
                    return int(val)
    return None


def is_caption_by_font_style(para, body_font_halfpt=None):
    """通过字体样式判断是否为图例/表例(caption)

    核心规则：sz=18(9pt) = caption, sz=24(12pt) = 正文
    无论文本是否以"图X"/"表X"开头，只要字号比正文小1pt以上就是caption

    识别优先级：字号差异(最可靠) > 居中对齐+文本模式 > 文本模式+无缩进
    """
    text = para.text.strip()
    if not text:
        return False

    # 方法1：XML层字号检测（最可靠）
    sz_val = get_para_font_size_xml(para)
    if sz_val is not None and body_font_halfpt is not None:
        if sz_val <= body_font_halfpt - 4:
            return True

    # 方法2：python-docx高层字号（兼容旧逻辑）
    caption_sizes = set()
    for run in para.runs:
        if run.font.size and run.text.strip():
            caption_sizes.add(run.font.size.pt)
    if body_font_halfpt and caption_sizes:
        max_cap = max(caption_sizes)
        if max_cap < body_font_halfpt / 2 - 1:
            return True

    # 方法3：居中对齐 + 文本模式
    align = para.paragraph_format.alignment
    if align == WD_ALIGN_PARAGRAPH.CENTER:
        if _CAPTION_TEXT_RE.match(text):
            return True

    # 方法4：文本模式 + 无首行缩进（最弱规则）
    if _CAPTION_TEXT_RE.match(text):
        indent = para.paragraph_format.first_line_indent
        if indent is None or indent == 0:
            return True

    return False


def is_caption_paragraph(para, body_font_size=None):
    """判断段落是否为图例/表例(caption) - 委托给字体样式检测"""
    return is_caption_by_font_style(para, body_font_size)


def is_table_paragraph(para):
    """判断段落是否属于表格单元格"""
    parent = para._element.getparent()
    return parent is not None and parent.tag.split('}')[-1] == 'tc'


def get_body_font_size(paragraphs):
    """统计正文字号（返回half-pt值，最常见值），用于区分图例/表例

    XML层w:sz值为half-pt（9pt=18, 12pt=24）
    """
    from collections import Counter
    size_counter = Counter()
    for para in paragraphs:
        if not para.text.strip():
            continue
        sz = get_para_font_size_xml(para)
        if sz and sz > 0:
            size_counter[sz] += 1
    if size_counter:
        return size_counter.most_common(1)[0][0]
    return None


def find_caption_and_context(paragraphs, pi, body_font_size=None):
    """在图片段落后方查找图例，并获取正上下文

    Returns: (caption, context_above, context_below)
    """
    caption = ""
    context_above = ""
    context_below = ""

    # 上下文：向前找最近的正文段落（跳过图例/表例段落）
    for k in range(pi - 1, max(pi - 10, -1), -1):
        prev_para = paragraphs[k]
        prev_text = prev_para.text.strip() if prev_para.text else ""
        if not prev_text:
            continue
        if is_caption_paragraph(prev_para, body_font_size):
            continue
        context_above = prev_text[-50:] if prev_text else ""
        break

    # 向后查找图例
    for j in range(pi + 1, min(pi + 6, len(paragraphs))):
        next_para = paragraphs[j]
        next_text = next_para.text.strip()
        if not next_text:
            continue
        if is_caption_paragraph(next_para, body_font_size):
            caption = next_text
            break
        else:
            break

    # 上下文：向后找最近的正文段落
    start_j = pi + 1
    for j in range(start_j, min(start_j + 10, len(paragraphs))):
        next_para = paragraphs[j]
        next_text = next_para.text.strip()
        if not next_text:
            continue
        if is_caption_paragraph(next_para, body_font_size):
            continue
        context_below = next_text[:50] if next_text else ""
        break

    return caption, context_above, context_below


def find_legend_paragraphs(paragraphs, start_pi, body_font_size=None, expected_kind=None):
    """从 start_pi 开始向后扫描，收集图例段落

    核心判定：用字体样式差异（字号 < 正文2pt以上）区分图例与正文。
    停止条件：遇到正文段落或 heading 或下一个图片/表格
    """
    legends = []
    for j in range(start_pi, min(start_pi + 10, len(paragraphs))):
        para = paragraphs[j]
        text = para.text.strip() if para.text else ""
        if not text:
            continue
        style = para.style.name if para.style else ''
        if 'Heading' in style or 'heading' in style:
            break
        if expected_kind == 'figure' and is_table_caption_text(text):
            break
        if expected_kind == 'table' and is_figure_caption_text(text):
            break
        is_cap = is_caption_paragraph(para, body_font_size)
        if is_cap:
            legends.append(text)
        else:
            break
    return legends


def find_context_body_text(paragraphs, pi, direction, body_font_size=None, image_map=None):
    """查找图片附近的正文段落文本

    direction='above': 向前查找最近的正文段落
    direction='below': 向后查找最近的正文段落（跳过图例段落）
    """
    if direction == 'above':
        for k in range(pi - 1, max(pi - 15, -1), -1):
            prev_para = paragraphs[k]
            prev_text = prev_para.text.strip() if prev_para.text else ""
            if not prev_text:
                continue
            if is_caption_paragraph(prev_para, body_font_size):
                continue
            return prev_text
        return ""
    else:  # below
        for j in range(pi + 1, min(pi + 15, len(paragraphs))):
            next_para = paragraphs[j]
            next_text = next_para.text.strip()
            if not next_text:
                continue
            if is_caption_paragraph(next_para, body_font_size):
                continue
            if is_table_paragraph(next_para):
                continue
            return next_text
        return ""
