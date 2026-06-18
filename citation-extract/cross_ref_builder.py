#!/usr/bin/env python3
r"""
LaTeX文献引用 → Word Bookmark + REF域代码 无损转换

将LaTeX的 \citep/\citet/\cite 引用转为Word原生交叉引用:
  - 参考文献条目 → Word Bookmark (_Bib_key)
  - 引用文本    → Word HYPERLINK域代码 (HYPERLINK \l "_Bib_key") — 可点击跳转且保留显示文本

引用样式由 cite_style_config 动态配置，不硬编码。
提供预置模板，也可自定义。
"""
import re
from copy import deepcopy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import RGBColor, Pt

# 从shared模块导入W_NS（消除重复代码）
import sys
from pathlib import Path
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))
from shared.latex_text_utils import citation_marker
from shared.word_xml_utils import W_NS


# ============================================================
# 引用样式模板接口
# ============================================================

class CiteStyleConfig:
    """文献引用样式配置模板

    定义Word中交叉引用的视觉呈现，匹配LaTeX编译后的PDF效果。

    Attributes:
        name: 样式名称
        color: 字体颜色 RGBColor (None=继承段落默认色，通常黑色)
        underline: 是否下划线
        underline_color: 下划线颜色 (None=跟随字体色)
        bold: 是否加粗
        italic: 是否斜体
        font_size: 字号 Pt (None=继承默认)
        superscript: 是否上标 (如IEEE [1]上标格式)
        cite_format: 引用格式 'author_year' | 'numbered'
            author_year: (Author, Year) 或 Author (Year)
            numbered: [1], [2,3] 等
    """

    def __init__(self, name='default', *, color=None, underline=False,
                 underline_color=None, bold=False, italic=False,
                 font_size=None, superscript=False, cite_format='author_year'):
        self.name = name
        self.color = color
        self.underline = underline
        self.underline_color = underline_color
        self.bold = bold
        self.italic = italic
        self.font_size = font_size
        self.superscript = superscript
        self.cite_format = cite_format

    def apply_to_rPr(self, rPr_elem):
        """将样式配置写入 w:rPr XML元素

        Args:
            rPr_elem: OxmlElement('w:rPr')，已创建的rPr节点
        """
        # 字符样式名(保持Hyperlink使REF域可点击)
        rStyle = OxmlElement('w:rStyle')
        rStyle.set(qn('w:val'), 'Hyperlink')
        rPr_elem.append(rStyle)

        # 颜色 — 覆盖Hyperlink样式的默认蓝色
        if self.color is not None:
            color_elem = OxmlElement('w:color')
            color_elem.set(qn('w:val'), str(self.color))
            rPr_elem.append(color_elem)

        # 加粗
        if self.bold:
            b_elem = OxmlElement('w:b')
            rPr_elem.append(b_elem)

        # 斜体
        if self.italic:
            i_elem = OxmlElement('w:i')
            rPr_elem.append(i_elem)

        # 字号
        if self.font_size is not None:
            sz_elem = OxmlElement('w:sz')
            sz_elem.set(qn('w:val'), str(int(self.font_size.pt * 2)))  # Word用半磅单位
            rPr_elem.append(sz_elem)

        # 上标
        if self.superscript:
            vertAlign = OxmlElement('w:vertAlign')
            vertAlign.set(qn('w:val'), 'superscript')
            rPr_elem.append(vertAlign)

        # 下划线
        if self.underline:
            u_elem = OxmlElement('w:u')
            u_elem.set(qn('w:val'), 'single')
            if self.underline_color is not None:
                u_elem.set(qn('w:color'), str(self.underline_color))
            rPr_elem.append(u_elem)

    def __repr__(self):
        return (f'CiteStyleConfig(name={self.name!r}, color={self.color}, '
                f'underline={self.underline}, bold={self.bold}, italic={self.italic}, '
                f'superscript={self.superscript}, cite_format={self.cite_format!r})')


# ============================================================
# 预置模板
# ============================================================

CITE_STYLES = {
    # === author-year 格式 ===

    # APA / 自然科学通用: 纯黑色文本，无装饰 (LaTeX默认PDF效果)
    'apa': CiteStyleConfig(
        name='apa',
        color=RGBColor(0, 0, 0),    # 黑色，覆盖Hyperlink蓝色
        underline=False,
        cite_format='author_year',
    ),

    # hyperref默认: 蓝色可点击 (hyperref包不设置citecolor时的默认)
    'hyperref_default': CiteStyleConfig(
        name='hyperref_default',
        color=RGBColor(0x1F, 0x49, 0x7D),  # LaTeX hyperref默认蓝色
        underline=False,
        cite_format='author_year',
    ),

    # Nature/Science: 黑色，作者部分可能正常，年份可能斜体
    'nature': CiteStyleConfig(
        name='nature',
        color=RGBColor(0, 0, 0),
        italic=False,
        cite_format='author_year',
    ),

    # Copernicus期刊: 黑色文本，无装饰
    'copernicus': CiteStyleConfig(
        name='copernicus',
        color=RGBColor(0, 0, 0),
        cite_format='author_year',
    ),

    # Chicago: 黑色，无装饰
    'chicago': CiteStyleConfig(
        name='chicago',
        color=RGBColor(0, 0, 0),
        cite_format='author_year',
    ),

    # === numbered 格式 ===

    # IEEE: [1] 上标格式
    'ieee': CiteStyleConfig(
        name='ieee',
        color=RGBColor(0, 0, 0),
        superscript=True,
        cite_format='numbered',
    ),

    # numbered非上标: [1] 行内
    'numbered': CiteStyleConfig(
        name='numbered',
        color=RGBColor(0, 0, 0),
        cite_format='numbered',
    ),

    # === hyperref自定义颜色 ===

    # hyperref citecolor=green
    'hyperref_green': CiteStyleConfig(
        name='hyperref_green',
        color=RGBColor(0x00, 0x80, 0x00),
        cite_format='author_year',
    ),

    # hyperref citecolor=red
    'hyperref_red': CiteStyleConfig(
        name='hyperref_red',
        color=RGBColor(0xCC, 0x00, 0x00),
        cite_format='author_year',
    ),
}


def detect_cite_style_from_tex(tex_content):
    """从LaTeX源文件自动检测引用样式

    检测逻辑:
    1. 检查 hyperref 包的 citecolor 选项
    2. 检查 bibliographystyle 推断格式类型
    3. 检查 documentclass 推断期刊模板
    4. 默认返回 'apa'

    Args:
        tex_content: LaTeX源文件内容

    Returns:
        str: CITE_STYLES中的模板key
    """
    # 1. 检测hyperref citecolor
    hyperref_m = re.search(
        r'\\usepackage\s*(?:\[([^\]]*)\])?\s*\{hyperref\}',
        tex_content
    )
    if hyperref_m and hyperref_m.group(1):
        options = hyperref_m.group(1)
        # 提取citecolor
        cc_m = re.search(r'citecolor\s*=\s*(\w+)', options)
        if cc_m:
            color_name = cc_m.group(1).lower()
            if color_name == 'green':
                return 'hyperref_green'
            elif color_name == 'red':
                return 'hyperref_red'
            elif color_name == 'blue':
                return 'hyperref_default'
            elif color_name == 'black':
                return 'apa'

    # 2. 检测bibliographystyle
    bst_m = re.search(r'\\bibliographystyle\{([^}]+)\}', tex_content)
    if bst_m:
        bst = bst_m.group(1).lower()
        if 'ieee' in bst:
            return 'ieee'
        if 'nature' in bst:
            return 'nature'
        if 'apa' in bst or 'apacite' in bst:
            return 'apa'
        if 'chicago' in bst:
            return 'chicago'
        if 'copernicus' in bst:
            return 'copernicus'
        # plain/unsrt/abbrv → numbered
        if bst in ('plain', 'unsrt', 'abbrv', 'plainnat'):
            return 'numbered'

    # 3. 检测documentclass
    cls_m = re.search(r'\\documentclass(?:\[.*?\])?\{([^}]+)\}', tex_content)
    if cls_m:
        cls = cls_m.group(1).lower()
        if 'ieee' in cls:
            return 'ieee'
        if 'copernicus' in cls or 'gmd' in cls or 'acp' in cls:
            return 'copernicus'
        if 'nature' in cls:
            return 'nature'

    # 4. 检测是否有natbib(暗示author-year)
    if re.search(r'\\usepackage.*\{natbib\}', tex_content):
        return 'apa'

    # 默认: 黑色无装饰(author-year)
    return 'apa'


def get_cite_style(style_key):
    """获取引用样式配置

    Args:
        style_key: CITE_STYLES中的key, 或 CiteStyleConfig 实例, 或 None(自动检测)

    Returns:
        CiteStyleConfig
    """
    if isinstance(style_key, CiteStyleConfig):
        return style_key
    if style_key and style_key in CITE_STYLES:
        return CITE_STYLES[style_key]
    # 默认
    return CITE_STYLES['apa']


# ============================================================
# Word XML操作工具
# ============================================================

def _bib_key_to_bookmark(key):
    """bibitem key → Word bookmark名"""
    name = re.sub(r'[^a-zA-Z0-9_]', '_', key)
    return f'_Bib_{name}'


def _get_max_bookmark_id(doc):
    """获取文档中现有bookmark的最大ID"""
    max_id = 0
    for elem in doc.element.body.iter(qn('w:bookmarkStart')):
        bid = elem.get(qn('w:id'))
        if bid and bid.isdigit():
            max_id = max(max_id, int(bid))
    return max_id


def _add_bookmark_to_paragraph(para, bookmark_name, bookmark_id):
    """在段落中添加bookmark标记"""
    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bookmark_id))
    bm_start.set(qn('w:name'), bookmark_name)

    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bookmark_id))

    p_elem = para._element
    pPr = p_elem.find(qn('w:pPr'))
    if pPr is not None:
        pPr.addnext(bm_start)
    else:
        p_elem.insert(0, bm_start)
    p_elem.append(bm_end)


def _create_ref_field_runs(bookmark_name, display_text, cite_style):
    """创建内部 HYPERLINK 域代码的5段式XML run列表

    Args:
        bookmark_name: 目标bookmark名称
        display_text: 引用在正文中的显示文本
        cite_style: CiteStyleConfig 实例

    Returns:
        list[OxmlElement]: 5个w:r元素
    """
    runs = []
    for fld_type, text in [
        ('begin', None),
        ('instrText', f' HYPERLINK \\l "{bookmark_name}" '),
        ('separate', None),
        ('text', display_text),
        ('end', None),
    ]:
        r = OxmlElement('w:r')
        # 应用引用样式
        rPr = OxmlElement('w:rPr')
        cite_style.apply_to_rPr(rPr)
        r.append(rPr)

        if fld_type in ('begin', 'separate', 'end'):
            fldChar = OxmlElement('w:fldChar')
            fldChar.set(qn('w:fldCharType'), fld_type)
            r.append(fldChar)
        elif fld_type == 'instrText':
            instr = OxmlElement('w:instrText')
            instr.set(qn('xml:space'), 'preserve')
            instr.text = text
            r.append(instr)
        elif fld_type == 'text':
            # 标记此run为已替换的引用显示文本，避免被重复替换
            r.set(qn('w:rsidR'), 'CITE002')
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = text
            r.append(t)

        runs.append(r)
    return runs


def _insert_runs_after_element(prev_elem, new_runs):
    """在XML元素后逐个插入多个run(保持顺序)"""
    for r in new_runs:
        prev_elem.addnext(r)
        prev_elem = r
    return prev_elem


def _ensure_hyperlink_style(doc, cite_style):
    """确保文档中存在Hyperlink字符样式

    根据 cite_style 配置更新Hyperlink样式的默认属性，
    使后续引用跳转域的视觉呈现匹配LaTeX编译效果。
    """
    try:
        style = doc.styles['Hyperlink']
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        style = doc.styles.add_style('Hyperlink', WD_STYLE_TYPE.CHARACTER)

    # 用cite_style配置覆盖Hyperlink样式默认值
    if cite_style.color is not None:
        style.font.color.rgb = cite_style.color
    if cite_style.underline is not None:
        style.font.underline = cite_style.underline
    if cite_style.bold is not None:
        style.font.bold = cite_style.bold
    if cite_style.italic is not None:
        style.font.italic = cite_style.italic
    if cite_style.font_size is not None:
        style.font.size = cite_style.font_size


def set_update_fields_on_open(doc):
    """设置文档打开时自动更新域代码"""
    settings = doc.settings.element
    update = settings.find(qn('w:updateFields'))
    if update is None:
        update = OxmlElement('w:updateFields')
        update.set(qn('w:val'), 'true')
        settings.append(update)


def clear_update_fields_on_open(doc):
    """Remove Word's update-fields-on-open flag to avoid security prompts."""
    settings = doc.settings.element
    update = settings.find(qn('w:updateFields'))
    if update is not None:
        settings.remove(update)


# ============================================================
# 段落匹配与替换
# ============================================================

def _find_bibitem_paragraphs(doc, cite_map):
    """在Word文档中找到参考文献条目对应的段落

    Args:
        doc: python-docx Document
        cite_map: {key: author_year_str}

    Returns:
        dict: {key: Paragraph}
    """
    key_to_para = {}
    ay_to_key = {}
    for key, ay_str in cite_map.items():
        ay_to_key[ay_str] = key
        clean = re.sub(r'\s+', ' ', ay_str).strip()
        ay_to_key[clean] = key

    paras = doc.paragraphs
    bib_start = len(paras)
    for i in range(len(paras) - 1, -1, -1):
        text = paras[i].text.strip()
        if re.match(r'^(References|Bibliography|参考文献|REFERENCES)', text, re.I):
            bib_start = i + 1
            break

    for i in range(bib_start, len(paras)):
        para = paras[i]
        text = para.text.strip()
        if not text:
            continue

        for ay_str, key in ay_to_key.items():
            if key in key_to_para:
                continue
            ay_m = re.match(r'(.+?)\s*\((\d{4})', ay_str)
            if ay_m:
                author_part = ay_m.group(1).strip()
                year_part = ay_m.group(2)
                first_author = author_part.split(',')[0].split(' et')[0].strip()
                if first_author and year_part in text and first_author in text:
                    key_to_para[key] = para
                    break

    return key_to_para


def _is_hyperlink_run(run_elem):
    """检查run元素是否在HYPERLINK域代码内或w:hyperlink元素内

    HYPERLINK跳转有两种XML表示:
    1. w:hyperlink元素 — 外部链接
    2. w:fldChar域代码 — 内部链接(我们的5段式HYPERLINK \l "bookmark")
    此外，替换后的显示文本run会带有自定义标记属性 w:rsidR="CITE002"
    """
    from docx.oxml.ns import qn
    # 检查自定义标记(替换后的显示文本run)
    rsid = run_elem.get(qn('w:rsidR'))
    if rsid == 'CITE002':
        return True
    parent = run_elem.getparent()
    while parent is not None:
        # 方式1: w:hyperlink元素
        if parent.tag == qn('w:hyperlink'):
            return True
        # 方式2: fldChar域代码 — 域控制run
        if parent.tag == qn('w:r'):
            fldChar = parent.find(qn('w:fldChar'))
            if fldChar is not None:
                return True
            instrText = parent.find(qn('w:instrText'))
            if instrText is not None:
                return True
        parent = parent.getparent()
    return False


def _replace_cite_in_paragraph(
    para, cite_text, bookmark_name, display_text, cite_style, from_end=False
):
    """在段落中替换引用文本为内部 HYPERLINK 域代码

    只在普通run文本中搜索，跳过已替换的HYPERLINK域代码，
    避免同一位置被重复替换导致死循环。

    Args:
        para: python-docx Paragraph
        cite_text: 要替换的文本
        bookmark_name: 目标bookmark名
        display_text: 域显示文本
        cite_style: CiteStyleConfig 实例

    Returns:
        bool: 是否替换成功
    """
    # 只在非HYPERLINK域的普通run中搜索
    runs = list(para.runs)
    if not runs:
        return False

    # 构建偏移表，跳过HYPERLINK域代码内的run
    run_offsets = []
    plain_parts = []
    offset = 0
    for r in runs:
        if _is_hyperlink_run(r._element):
            # HYPERLINK域内的run，用空字符串占位(不搜索)
            run_offsets.append((offset, offset, r))
            continue
        run_offsets.append((offset, offset + len(r.text), r))
        plain_parts.append(r.text)
        offset += len(r.text)

    # 在普通run文本中搜索
    full_text = ''.join(plain_parts)
    pos = full_text.rfind(cite_text) if from_end else full_text.find(cite_text)
    if pos < 0:
        return False

    start_run = None
    end_run = None
    start_char = None
    end_char = None

    for i, (rstart, rend, r) in enumerate(run_offsets):
        # 跳过HYPERLINK域内的run(offset长度为0)
        if rstart == rend:
            continue
        if start_run is None and rend > pos:
            start_run = i
            start_char = pos - rstart
        if end_run is None and rend >= pos + len(cite_text):
            end_run = i
            end_char = pos + len(cite_text) - rstart
            break

    if start_run is None or end_run is None:
        return False

    field_runs = _create_ref_field_runs(bookmark_name, display_text, cite_style)
    p_elem = para._element

    if start_run == end_run:
        run = runs[start_run]
        before = run.text[:start_char]
        after = run.text[end_char:]

        if before:
            run.text = before
            _insert_runs_after_element(run._element, field_runs)
            if after:
                after_r = OxmlElement('w:r')
                t = OxmlElement('w:t')
                t.set(qn('xml:space'), 'preserve')
                t.text = after
                after_r.append(t)
                field_runs[-1].addnext(after_r)
        else:
            _insert_runs_after_element(run._element, field_runs)
            if after:
                run.text = after
            else:
                p_elem.remove(run._element)
    else:
        runs[start_run].text = runs[start_run].text[:start_char]
        _insert_runs_after_element(runs[start_run]._element, field_runs)

        end_text = runs[end_run].text[end_char:]
        if end_text:
            runs[end_run].text = end_text
        else:
            p_elem.remove(runs[end_run]._element)

        for i in range(start_run + 1, end_run):
            if i != end_run:
                p_elem.remove(runs[i]._element)

    return True


def _replace_marker_in_paragraph(para, marker, bookmark_name, display_text, cite_style):
    """Replace one stable citation marker in its exact Word text node."""
    for text_elem in para._element.iter(qn('w:t')):
        text = text_elem.text or ''
        pos = text.rfind(marker)
        if pos < 0:
            continue

        run_elem = text_elem.getparent()
        before = text[:pos]
        after = text[pos + len(marker):]
        text_elem.text = before

        field_runs = _create_ref_field_runs(bookmark_name, display_text, cite_style)
        last_elem = _insert_runs_after_element(run_elem, field_runs)
        if after:
            after_run = deepcopy(run_elem)
            for child in list(after_run):
                if child.tag != qn('w:rPr'):
                    after_run.remove(child)
            after_text = OxmlElement('w:t')
            after_text.set(qn('xml:space'), 'preserve')
            after_text.text = after
            after_run.append(after_text)
            last_elem.addnext(after_run)
        return True
    return False


# ============================================================
# 主入口
# ============================================================

def insert_bib_cross_references(doc, cite_map, cite_style='apa'):
    """在Word文档中插入文献交叉引用(bookmark + HYPERLINK域代码)

    主入口函数, 供 tex_to_word.py 调用。

    Args:
        doc: python-docx Document对象
        cite_map: {key: author_year_str} 来自_parse_bbl
        cite_style: 引用样式配置，支持:
            - str: CITE_STYLES中的模板key (如 'apa', 'ieee', 'copernicus')
            - CiteStyleConfig: 自定义样式实例
            - None: 使用默认 'apa'

    Returns:
        dict: {'bookmarks_added': int, 'refs_replaced': int}
    """
    if not cite_map:
        return {'bookmarks_added': 0, 'refs_replaced': 0}

    # 解析样式配置
    style = get_cite_style(cite_style)

    # 确保Hyperlink样式存在且属性匹配LaTeX效果
    _ensure_hyperlink_style(doc, style)

    next_bm_id = _get_max_bookmark_id(doc) + 1

    # === 第1步: 在参考文献条目添加bookmark ===
    key_to_bm = {}
    existing_bm_names = {
        elem.get(qn('w:name'))
        for elem in doc.element.body.iter(qn('w:bookmarkStart'))
        if elem.get(qn('w:name'))
    }
    for key in cite_map:
        bm_name = _bib_key_to_bookmark(key)
        if bm_name in existing_bm_names:
            key_to_bm[key] = bm_name
    existing_bookmarks = len(key_to_bm)

    key_to_para = _find_bibitem_paragraphs(doc, cite_map)
    bookmarks_added = 0

    for key, para in key_to_para.items():
        if key in key_to_bm:
            continue
        bm_name = _bib_key_to_bookmark(key)
        _add_bookmark_to_paragraph(para, bm_name, next_bm_id)
        key_to_bm[key] = bm_name
        next_bm_id += 1
        bookmarks_added += 1

    # === 第2步: 替换正文中的引用文本为内部 HYPERLINK 域代码 ===
    refs_replaced = 0

    # 构建引用匹配模式。优先使用预处理阶段写入的稳定key标记，
    # 文本模式仅用于兼容旧的prepared.tex或外部DOCX。
    cite_patterns = []
    for key, ay_str in cite_map.items():
        bm_name = key_to_bm.get(key, _bib_key_to_bookmark(key))

        if style.cite_format == 'numbered':
            cite_patterns.extend([
                {'bm_name': bm_name, 'display': ay_str, 'search_text': citation_marker(key, 'P'), 'marker': True},
                {'bm_name': bm_name, 'display': ay_str, 'search_text': citation_marker(key, 'M'), 'marker': True},
                {'bm_name': bm_name, 'display': ay_str, 'search_text': citation_marker(key, 'T'), 'marker': True},
                {'bm_name': bm_name, 'display': ay_str, 'search_text': ay_str, 'marker': False},
            ])
        else:
            # author-year格式
            ay_m = re.match(r'(.+?)\s*\((\d{4}[a-z]?)\)', ay_str)
            if ay_m:
                author = ay_m.group(1).strip()
                year = ay_m.group(2)
                citep_text = f'({author}, {year})'
                inner_text = f'{author}, {year}'
                cite_patterns.extend([
                    {'bm_name': bm_name, 'display': citep_text, 'search_text': citation_marker(key, 'P'), 'marker': True},
                    {'bm_name': bm_name, 'display': inner_text, 'search_text': citation_marker(key, 'M'), 'marker': True},
                    {'bm_name': bm_name, 'display': ay_str, 'search_text': citation_marker(key, 'T'), 'marker': True},
                    {'bm_name': bm_name, 'display': citep_text, 'search_text': citep_text, 'marker': False},
                    {'bm_name': bm_name, 'display': ay_str, 'search_text': ay_str, 'marker': False},
                ])
            else:
                cite_patterns.extend([
                    {'bm_name': bm_name, 'display': ay_str, 'search_text': citation_marker(key, 'P'), 'marker': True},
                    {'bm_name': bm_name, 'display': ay_str, 'search_text': ay_str, 'marker': False},
                ])

    # 遍历正文段落(参考文献区域之前)
    paras = doc.paragraphs
    bib_start = len(paras)
    for i in range(len(paras) - 1, -1, -1):
        text = paras[i].text.strip()
        if re.match(r'^(References|Bibliography|参考文献|REFERENCES)', text, re.I):
            bib_start = i
            break

    for i in range(bib_start):
        para = paras[i]
        para_text = para.text
        if not para_text:
            continue

        for cp in (item for item in cite_patterns if item.get('marker')):
            while _replace_marker_in_paragraph(
                para, cp['search_text'], cp['bm_name'], cp['display'], style,
            ):
                refs_replaced += 1

        for cp in (item for item in cite_patterns if not item.get('marker')):
            search_text = cp['search_text']
            if not search_text or len(search_text) < 8:
                continue
            count = 0
            while count < 10:
                if _replace_cite_in_paragraph(para, search_text, cp['bm_name'], cp['display'], style):
                    refs_replaced += 1
                    count += 1
                    para_text = para.text
                else:
                    break

    # === 第3步: 内部HYPERLINK域不需要打开时自动更新 ===
    clear_update_fields_on_open(doc)

    total_bookmarks = existing_bookmarks + bookmarks_added
    print(f'  [文献交叉引用] 样式={style.name}, Bookmarks: {total_bookmarks}/{len(cite_map)} '
          f'(existing={existing_bookmarks}, added={bookmarks_added}), citation links: {refs_replaced}')
    return {'bookmarks_added': bookmarks_added, 'refs_replaced': refs_replaced}
