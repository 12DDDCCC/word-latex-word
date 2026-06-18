"""语义分类模块

基于多维度特征推断段落语义类型。
"""

import re


# 文本模式正则（优先级1：最可靠）
_RE_ABSTRACT_LABEL = re.compile(r'^\s*Abstract\s*[:：.]?\s*$', re.IGNORECASE)
_RE_ABSTRACT_INLINE = re.compile(r'^\s*Abstract\s*[:：.]\s*', re.IGNORECASE)
_RE_KEYWORDS = re.compile(r'^\s*Key\s*words?\s*[:：]|^关键词\s*[:：]', re.IGNORECASE)
_RE_FIGURE_CAP = re.compile(r'^\s*\u56fe\s*[\d\.]+|^Figure\s*[\d\.]+|^Fig\.?\s*[\d\.]+', re.IGNORECASE)
_RE_TABLE_CAP = re.compile(r'^\s*\u8868\s*[\d\.]+|^Table\s*[\d\.]+|^Talbe\s*[\d\.]+', re.IGNORECASE)
_RE_FIGURE_CAP_STRICT = re.compile(r'^\s*(?:\u56fe|Figure|Fig\.?)\s*\d+(?:\.\d+)*\s*[:\uff1a,\uff0c\.\u3002\u3001\s]+', re.IGNORECASE)
_RE_TABLE_CAP_STRICT = re.compile(r'^\s*(?:\u8868|Table|Talbe)\s*\d+(?:\.\d+)*\s*[:\uff1a,\uff0c\.\u3002\u3001\s]+', re.IGNORECASE)
_RE_REFERENCE = re.compile(r'^\s*References?\s*$|^参考文献\s*$', re.IGNORECASE)
_RE_ACKNOWLEDGEMENT = re.compile(r'^\s*Acknowledgements?|^致谢', re.IGNORECASE)
_RE_DECLARATION = re.compile(
    r'^\s*(Code\s*availability|Data\s*availability|Competing\s*interests|'
    r'Author\s*contribution|Sample\s*availability|Copyright)',
    re.IGNORECASE
)
_RE_AFFILIATION = re.compile(r'大学|学院|研究所|研究院|University|Institute|Laboratory|Lab\b', re.IGNORECASE)

# 机构关键词（用于识别 affiliation 段落）
_AFFIL_KEYWORDS = ('大学', '学院', '研究所', '研究院', 'University', 'Institute', 'Laboratory')


def _dominant_size(runs):
    """提取段落的主导字号(pt)。返回None表示无法判断。"""
    if not runs:
        return None
    sizes = [r['size_pt'] for r in runs
             if isinstance(r, dict) and r.get('size_pt') is not None]
    if not sizes:
        return None
    # 取众数作为主导字号
    from collections import Counter
    return Counter(sizes).most_common(1)[0][0]


def _body_font_size(doc_meta):
    """获取文档正文字号(pt)。通过已观察到的body段落推断，默认12pt。"""
    return doc_meta.get('_body_font_size', 12.0)


def classify_semantic_type(para_info, index, total, prev_type=None, doc_meta=None):
    """基于多维度特征推断段落语义类型

    Args:
        para_info: extract_paragraph() 返回的段落 dict
        index: 段落在列表中的序号(0-based)
        total: 段落总数
        prev_type: 前一个段落的 semantic_type
        doc_meta: 文档级元数据状态 dict，跨段落传递

    Returns:
        str: 语义类型标签
    """
    if doc_meta is None:
        doc_meta = {}

    text = para_info.get('text', '').strip()
    low = text.lower()
    h_level = para_info.get('heading_level')
    align = para_info.get('alignment')
    has_formula = para_info.get('has_formula', False)
    runs = para_info.get('runs', [])
    bold_count = sum(1 for r in runs if r.get('bold'))
    total_runs = len(runs)
    all_bold = total_runs > 0 and bold_count == total_runs
    has_cite = any(r.get('is_cite') for r in runs)
    indent = para_info.get('first_line_indent_pt')

    # ── 优先级1: Word 样式名 ──
    if h_level is not None:
        # heading 结束 abstract 区域
        if doc_meta.get('in_abstract'):
            doc_meta['in_abstract'] = False
        return 'heading'

    # ── 优先级1: 空段落 ──
    if not text and not has_formula:
        return 'empty'

    # ── 优先级1: 文本模式正则 ──
    if _RE_ABSTRACT_LABEL.match(text):
        return 'abstract_label'
    if _RE_KEYWORDS.match(text):
        doc_meta['in_abstract'] = False
        return 'keywords'
    # 以"图X.X"/"Figure X.X"/"Table X.X"开头，但内容很长（>100字）→ 正文，不是caption
    # 但如果包含子图标记（如a), b), c)等），则判定为图例说明
    fig_cap_m = _RE_FIGURE_CAP.match(text)
    tbl_cap_m = _RE_TABLE_CAP.match(text)
    if fig_cap_m or tbl_cap_m:
        cap_m = fig_cap_m or tbl_cap_m
        after_num = text[cap_m.end():].lstrip()
        strict_cap = _RE_FIGURE_CAP_STRICT.match(text) if fig_cap_m else _RE_TABLE_CAP_STRICT.match(text)

        # \u2500\u2500 \u7edf\u4e00\u5b57\u53f7\u4f18\u5148\u7b56\u7565 \u2500\u2500
        # caption\u7684\u5b57\u53f7\u4e00\u5b9a\u6bd4\u6b63\u6587\u5c0f\uff08\u59829pt vs 12pt, \u621610pt vs 11pt\uff09
        # \u8fd9\u662f\u53ef\u9760\u7684\u5224\u65ad\u4f9d\u636e\uff1a\u5b57\u53f7 < \u6b63\u6587 \u2192 caption\uff1b\u5b57\u53f7 = \u6b63\u6587 \u2192 body\u5f15\u7528
        size = _dominant_size(runs)
        body_size = _body_font_size(doc_meta)
        _is_small_font = (size is not None and size < body_size)

        if strict_cap:
            # \u53d9\u4e8b\u5f15\u7528\u68c0\u6d4b\uff1a\u56fe/\u8868\u53f7\u540e\u7d27\u8ddf"\u4e2d/\u662f/\u4e3a/\u8bf4\u660e/\u5c55\u793a/\u53ef\u4ee5\u770b\u5230/\u6240\u793a"
            if re.match(r'^(\u4e2d|\u662f|\u4e3a|\u8bf4\u660e|\u5c55\u793a|\u53ef\u4ee5\u770b\u5230|\u6240\u793a)', after_num):
                if _is_small_font:
                    pass  # \u5b57\u53f7\u5c0f\u2192caption\uff0c\u7ee7\u7eed
                else:
                    return 'body'  # \u5b57\u53f7=\u6b63\u6587\u2192\u53d9\u4e8b\u5f15\u7528
            # \u5b57\u53f7\u5224\u65ad\uff1a\u5b57\u53f7=\u6b63\u6587\u2192body\u5f15\u7528\uff08\u65e0\u8bba\u957f\u77ed\uff09
            if not _is_small_font and size is not None:
                return 'body'
            return 'figure_caption' if fig_cap_m else 'table_caption'

        # \u975estrict_cap\u8def\u5f84\uff08\u56fe/\u8868\u53f7\u540e\u65e0\u5206\u9694\u7b26\uff09
        if re.match(r'^(\u4e2d|\u662f|\u4e3a|\u8bf4\u660e|\u5c55\u793a|\u53ef\u4ee5\u770b\u5230|\u6240\u793a)', after_num):
            if _is_small_font:
                return 'figure_caption' if fig_cap_m else 'table_caption'
            return 'body'
        has_subfigure_markers = re.search(r'[a-zA-Z][\s]*[),]', text) or re.search(r'[\u4e00-\u9fff].*?[a-zA-Z][\s]*[),]', text)
        if len(text) > 100 and not has_subfigure_markers:
            if _is_small_font:
                return 'figure_caption' if fig_cap_m else 'table_caption'
            return 'body'
        return 'figure_caption' if fig_cap_m else 'table_caption'
    if _RE_REFERENCE.match(text):
        return 'reference'
    if _RE_ACKNOWLEDGEMENT.match(text):
        return 'acknowledgement'
    if _RE_DECLARATION.match(text):
        return 'declaration'

    # ── 优先级2: 上下文推断 ──
    # Abstract: 开头行（内容与标签合并在同一段落）
    if _RE_ABSTRACT_INLINE.match(text):
        doc_meta['in_abstract'] = True
        return 'abstract'

    # 在 abstract_label 之后、首个 heading 之前 → abstract
    if prev_type == 'abstract_label' or doc_meta.get('in_abstract'):
        # 遇到 heading 或 keywords 则结束 abstract 区
        if h_level is not None or _RE_KEYWORDS.match(text):
            doc_meta['in_abstract'] = False
        else:
            doc_meta['in_abstract'] = True
            return 'abstract'

    # ── 优先级3: 格式+位置启发式 ──
    # 标题：前5段 + 居中 + 全部加粗 + 长度>5
    if index < 5 and align == 'center' and all_bold and len(text) > 5:
        if not doc_meta.get('title_found'):
            doc_meta['title_found'] = True
            return 'title'

    # 作者：标题之后(前15段) + 居中 + 短文本(<50字) + 非加粗
    if (doc_meta.get('title_found') and not doc_meta.get('author_done')
            and index < 15 and align == 'center' and not all_bold
            and len(text) < 50 and len(text) > 1
            and not _RE_AFFILIATION.search(text)):
        return 'author'

    # 机构：居中 + 含机构关键词
    if align == 'center' and _RE_AFFILIATION.search(text):
        doc_meta['author_done'] = True
        return 'affiliation'

    # 作者区结束标记（遇到非居中段落）
    if doc_meta.get('title_found') and not doc_meta.get('author_done') and align != 'center' and index < 15:
        doc_meta['author_done'] = True

    # ── 优先级4: 独立公式段落 ──
    # 情况A: 只有公式和编号，无其他文字 → 独立公式
    # 情况B: 有公式，且非公式文本包含公式编号模式如 (2-1) → 也是独立公式
    if has_formula:
        formula_runs = [r for r in runs if r.get('type') == 'formula']
        text_runs = [r for r in runs if r.get('type') == 'text' and r.get('text', '').strip()]
        # 文字 run 只有编号如 (2-1) → 独立公式
        non_num_text = ''.join(r.get('text', '') for r in text_runs)
        non_num_text = re.sub(r'\([\d\-\.]+\)', '', non_num_text).strip()
        if formula_runs and not non_num_text:
            return 'display_formula'
        # 情况B: 段落包含公式 + 公式编号模式 → 独立公式
        # 但必须是短段落（<100字符），避免将包含公式的长段落正文误判
        # 例如: " , i = 1, 2, ... , N (2-1)" 包含公式和编号
        if formula_runs and len(text) < 100 and re.search(r'\([\d\-\.]+\)', text):
            return 'display_formula'

    # ── 优先级5: 默认分类 ──
    # 有首行缩进、长文本、含引用/公式 → 正文
    # 同时记录正文字号到doc_meta，供后续字号兜底使用
    if indent and indent > 0:
        _update_body_font_size(doc_meta, runs)
        return 'body'
    if len(text) > 50:
        _update_body_font_size(doc_meta, runs)
        return 'body'
    if has_cite or has_formula:
        _update_body_font_size(doc_meta, runs)
        return 'body'

    return 'unknown'


def _update_body_font_size(doc_meta, runs):
    """从body段落中记录正文字号，用于后续字号兜底判断。"""
    size = _dominant_size(runs)
    if size is not None and '_body_font_size' not in doc_meta:
        doc_meta['_body_font_size'] = size
