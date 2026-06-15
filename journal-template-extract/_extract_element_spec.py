#!/usr/bin/env python3
r"""
元素规格提取模块: 标题/作者/摘要/章节/正文/caption/表格/图片/脚注/文献
包含方法:
  - _extract_title
  - _extract_author
  - _extract_abstract
  - _find_balanced_braces (静态方法)
  - _extract_headings
  - _extract_body_text
  - _extract_caption
  - _extract_table_spec
  - _extract_figure_spec
  - _extract_footnote
  - _extract_bibliography
"""
import re
from collections import OrderedDict
from pathlib import Path
import sys
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.latex_parse_utils import (
    BS, _cmd, _la_size_to_pt, _size_style, LATEX_SIZE_PT
)


def _extract_title(self):
    """提取论文标题规格"""
    spec = OrderedDict()
    # 从\maketitle中搜索标题格式
    mt = re.search(_cmd('def') + _cmd('@?maketitle') + r'\b(.{0,4000})', self.content, re.DOTALL)
    if mt:
        body = mt.group(1)
        # 搜索 \LARGE\bfseries\@title 模式
        tl = re.search(r'(' + BS + r'(?:LARGE|Large|large|huge|Huge|normalsize|small)\s*(' + BS + r'(?:bfseries|mdseries|itshape|upshape|sffamily|rmfamily|mathversion)\s*)*)' + BS + r'@?title', body)
        if tl:
            spec.update(_size_style(tl.group(1), self.base_size))
        # 对齐和颜色
        if BS + r'centering' in body or BS + r'center{' in body:
            spec.setdefault('alignment', 'center')
        cm = re.search(BS + r'color\{(\w+)\}', body[:2000])
        if cm:
            spec['color'] = cm.group(1)
        if BS + r'sffamily' in body[:2000]:
            spec.setdefault('font_family', 'sans-serif')

    # 从\titlefont获取
    tf = re.search(_cmd('titlefont') + r'\s*\{([^}]+)\}', self.content)
    if tf:
        spec.update(_size_style(tf.group(1), self.base_size))

    spec.setdefault('size_name', 'LARGE')
    spec.setdefault('size_pt', _la_size_to_pt('LARGE', self.base_size))
    spec.setdefault('weight', 'bold')
    spec.setdefault('shape', 'normal')
    spec.setdefault('font_family', 'serif')
    spec.setdefault('alignment', 'center')
    return spec


def _extract_author(self):
    """提取作者规格"""
    spec = OrderedDict()
    # \Authfont — 取最后一个定义（\renewcommand覆盖\newcommand）
    authfont_body = None
    for af in re.finditer(_cmd('Authfont') + r'\s*\{([^}]+)\}', self.content):
        authfont_body = af.group(1).strip()
    if authfont_body:
        spec.update(_size_style(authfont_body, self.base_size))
        spec['font_declaration'] = authfont_body
        # \normalfont 隐含 \normalsize
        if 'normalfont' in authfont_body and spec.get('size_pt') is None:
            spec['size_name'] = 'normalsize'
            spec['size_pt'] = self.base_size

    # \Affilfont — 取最后一个定义
    affilfont_body = None
    for aff in re.finditer(_cmd('Affilfont') + r'\s*\{([^}]+)\}', self.content):
        affilfont_body = aff.group(1).strip()
    if affilfont_body:
        spec['affilfont'] = affilfont_body
        # \normalsize\normalfont → normalsize
        if 'normalsize' in affilfont_body:
            spec['size_name'] = 'normalsize'
            spec['size_pt'] = self.base_size
        elif 'normalfont' in affilfont_body and spec.get('size_pt') is None:
            spec['size_name'] = 'normalsize'
            spec['size_pt'] = self.base_size

    # \maketitle中的重定义（优先级最高）
    mt = re.search(_cmd('def') + _cmd('@?maketitle') + r'\b(.{0,4000})', self.content, re.DOTALL)
    if mt:
        body = mt.group(1)
        # 取最后一个 \renewcommand\Authfont
        last_af = None
        for af2 in re.finditer(BS + r'renewcommand' + BS + r'Authfont\s*\{([^}]+)\}', body):
            last_af = af2.group(1).strip()
        if last_af:
            spec.update(_size_style(last_af, self.base_size))
            spec['font_declaration'] = last_af
            if 'normalfont' in last_af and spec.get('size_pt') is None:
                spec['size_name'] = 'normalsize'
                spec['size_pt'] = self.base_size
        al = re.search(BS + r'(normalsize|large|Large|small|footnotesize)\b[^@\n]{0,40}' + BS + r'@?author', body)
        if al:
            spec['size_name'] = al.group(1)
            spec['size_pt'] = _la_size_to_pt(al.group(1), self.base_size)

    spec.setdefault('size_name', 'normalsize')
    spec.setdefault('size_pt', self.base_size)
    spec.setdefault('weight', 'normal')
    return spec


def _extract_abstract(self):
    """提取摘要规格"""
    spec = OrderedDict([('label', 'Abstract'), ('label_weight', 'bold')])
    for m in re.finditer(
        BS + r'renewenvironment\{abstract\}\[?\d?\]?\s*\{([^}]+)\}\s*\{([^}]+)\}',
        self.content, re.DOTALL
    ):
        begin_code = m.group(1)
        spec.update(_size_style(begin_code, self.base_size))
        spec['label_weight'] = 'bold' if 'bfseries' in begin_code or 'textbf' in begin_code else 'normal'
        if 'quote' in begin_code or 'quotation' in begin_code:
            spec['indent'] = True
        wm = re.search(BS + r'begin\{(?:minipage|lrbox)\}\{([^}]+)\}', begin_code)
        if wm:
            spec['width'] = wm.group(1)

    an = re.search(BS + r'renewcommand\{' + BS + r'abstractname\}\{([^}]+)\}', self.content)
    if an:
        spec['label'] = an.group(1)
    return spec


def _find_balanced_braces(text, start):
    """从start位置(必须是'{')开始，找到匹配的'}'的索引。返回内容不含外层花括号。"""
    if start >= len(text) or text[start] != '{':
        return None, start
    depth = 0
    i = start
    while i < len(text):
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                return text[start+1:i], i
        i += 1
    return None, start


def _extract_headings(self):
    """提取章节标题规格"""
    headings = OrderedDict()
    flat = re.sub(r'\n\s+', ' ', self.content)

    for level_name in ('section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'):
        # 收集所有 \def\level_name{...} 中的 \@startsection 定义
        candidates = []
        for dm in re.finditer(_cmd('def') + _cmd(level_name) + r'\s*\{', flat):
            inner, end = self._find_balanced_braces(flat, dm.end() - 1)
            if not inner:
                continue
            ss_idx = inner.find('@startsection')
            if ss_idx < 0:
                continue
            brace_start = inner.find('{', ss_idx)
            if brace_start < 0:
                continue
            args = []
            pos = brace_start
            for _ in range(6):
                content, pos = self._find_balanced_braces(inner, pos)
                if content is None:
                    break
                args.append(content)
                pos += 1
                while pos < len(inner) and inner[pos] in ' \t':
                    pos += 1
            if len(args) >= 6:
                style_code = args[5]
                h = _size_style(style_code, self.base_size)
                h['level'] = int(args[1]) if args[1].isdigit() else 1
                h['indent'] = args[2] if args[2] else '0'
                h['before_skip'] = args[3].strip()
                h['after_skip'] = args[4].strip()
                if 'sffamily' in style_code:
                    h['font_family'] = 'sans-serif'
                if 'color{textcol}' in style_code:
                    h['color'] = 'textcol'
                if 'raggedright' in style_code:
                    h['alignment'] = 'left'
                # 标记分支类型
                before = flat[max(0, dm.start()-80):dm.start()]
                if 'classical' in before:
                    h['_branch'] = 'classical'
                elif 'sffamily' in style_code or 'sansserifface' in before:
                    h['_branch'] = 'sansserifface'
                elif 'discussions' in before:
                    h['_branch'] = 'discussions'
                else:
                    h['_branch'] = 'default'
                candidates.append(h)

        # 优先选择 classical 版本（默认风格），其次选不含sffamily的
        chosen = None
        for c in candidates:
            if c.get('_branch') == 'classical':
                chosen = c
                break
        if not chosen:
            for c in candidates:
                if c.get('_branch') != 'sansserifface':
                    chosen = c
                    break
        if not chosen and candidates:
            chosen = candidates[0]
        if chosen:
            chosen.pop('_branch', None)
            headings[level_name] = chosen
            continue

        # 方法2: 标准 \@startsection{section}{1}... 不在\def内
        for sm in re.finditer(_cmd('@startsection') + r'\{' + level_name + r'\}', flat):
            brace_start = flat.find('{', sm.end())
            if brace_start < 0:
                continue
            args = []
            pos = sm.end() - 1
            for _ in range(6):
                content, pos = self._find_balanced_braces(flat, pos)
                if content is None:
                    break
                args.append(content)
                pos += 1
                while pos < len(flat) and flat[pos] in ' \t':
                    pos += 1
            if len(args) >= 6:
                style_code = args[5]
                h = _size_style(style_code, self.base_size)
                h['level'] = int(args[1]) if args[1].isdigit() else 1
                h['indent'] = args[2] if args[2] else '0'
                h['before_skip'] = args[3].strip()
                h['after_skip'] = args[4].strip()
                headings[level_name] = h
                break

    # 如果前面没找到，尝试从@sect解析
    if not headings:
        sect_match = re.search(_cmd('def') + _cmd('@sect') + r'[^{]*\{(.{0,800})', self.content, re.DOTALL)
        if sect_match:
            body = sect_match.group(1)
            h = _size_style(body, self.base_size)
            if 'bfseries' in body:
                h['weight'] = 'bold'
            headings['section'] = h

    # 从\section的重定义推断
    for level_name in ('section', 'subsection', 'subsubsection'):
        if level_name not in headings:
            pat = BS + r'renewcommand\s*' + BS + level_name + r'\s*(?:\[[^\]]*\])?\s*\{(.{0,300})'
            m = re.search(pat, flat)
            if m:
                headings[level_name] = _size_style(m.group(1), self.base_size)

    # \let别名解析: \let\paragraph=\subsubsection → paragraph继承subsubsection的格式
    for level_name in ('paragraph', 'subparagraph'):
        if level_name not in headings:
            let_m = re.search(BS + r'let\s*' + BS + level_name + r'\s*=?\s*' + BS + r'(\w+)', flat)
            if let_m:
                alias = let_m.group(1)
                if alias in headings:
                    headings[level_name] = OrderedDict(headings[alias])
                    headings[level_name]['alias_of'] = alias

    # 默认值填充
    defaults = {
        'section': {'size_name': 'normalsize', 'size_pt': 10, 'weight': 'bold', 'shape': 'normal',
                    'before_skip': '-2em', 'after_skip': '1em'},
        'subsection': {'size_name': 'normalsize', 'size_pt': 10, 'weight': 'bold', 'shape': 'normal'},
        'subsubsection': {'size_name': 'normalsize', 'size_pt': 10, 'weight': 'bold', 'shape': 'italic'},
    }
    for key, default in defaults.items():
        if key not in headings:
            headings[key] = default
    return headings


def _extract_body_text(self):
    """提取正文规格"""
    spec = OrderedDict()
    spec['font_size_pt'] = self.base_size
    spec['font_family'] = 'serif (Times New Roman)'
    bs = re.search(BS + r'baselinestretch\s*=?\s*([\d.]+)', self.content)
    ls = re.search(BS + r'linespread\{([\d.]+)\}', self.content)
    if bs:
        spec['line_spacing'] = f'{bs.group(1)}× baselineskip'
    elif ls:
        spec['line_spacing'] = f'{ls.group(1)}× baselineskip'
    else:
        spec['line_spacing'] = 'single (1.0)'
    pi = re.search(BS + r'parindent\s*=?\s*([\d.]+\w+)', self.content)
    spec['first_line_indent'] = pi.group(1) if pi else '1em'
    ps = re.search(BS + r'parskip\s*=?\s*([\d.]+\w+)', self.content)
    spec['paragraph_skip'] = ps.group(1) if ps else '0pt'
    return spec


def _extract_caption(self):
    """提取Caption规格"""
    spec = OrderedDict([
        ('font_size', 'small'), ('font_size_pt', 9),
        ('weight', 'normal'), ('label_weight', 'bold'),
        ('separator', '.'), ('figure_position', 'below'), ('table_position', 'above'),
    ])
    cap_match = re.search(
        BS + r'long' + BS + r'def' + BS + r'@makecaption\{[^}]*\}\s*\{(.{0,1200})\}', self.content, re.DOTALL
    )
    if cap_match:
        body = cap_match.group(1)
        spec.update(_size_style(body, self.base_size))
        if 'bfseries' in body:
            spec['label_weight'] = 'bold'
        if ': ' in body or BS + r'colon' in body:
            spec['separator'] = ':'
        elif '. ' in body:
            spec['separator'] = '.'
        if BS + r'small' in body:
            spec['font_size'] = 'small'; spec['font_size_pt'] = 9
        elif BS + r'footnotesize' in body:
            spec['font_size'] = 'footnotesize'; spec['font_size_pt'] = 8
        # 解析\ifx\@captype条件分支推断figure/table位置
        fig_branch = re.search(BS + r'def' + BS + r'@tempa\{figure\}.*?' + BS + r'ifx' + BS + r'@captype' + BS + r'@tempa(.{0,200}?)(?:' + BS + r'fi|' + BS + r'else)', body, re.DOTALL)
        if fig_branch:
            fig_body = fig_branch.group(1)
            if BS + r'abovecaptionskip' in fig_body and BS + r'belowcaptionskip' not in fig_body:
                spec['figure_position'] = 'below'
            elif BS + r'belowcaptionskip' in fig_body and BS + r'abovecaptionskip' not in fig_body:
                spec['figure_position'] = 'above'
        tbl_branch = re.search(BS + r'def' + BS + r'@tempa\{table\}.*?' + BS + r'ifx' + BS + r'@captype' + BS + r'@tempa(.{0,200}?)(?:' + BS + r'fi|' + BS + r'else)', body, re.DOTALL)
        if tbl_branch:
            tbl_body = tbl_branch.group(1)
            if BS + r'abovecaptionskip' in tbl_body and BS + r'belowcaptionskip' not in tbl_body:
                spec['table_position'] = 'above'
            elif BS + r'belowcaptionskip' in tbl_body and BS + r'abovecaptionskip' not in tbl_body:
                spec['table_position'] = 'below'

    cf = re.search(BS + r'(?:caption(?:text)?font|cfsf)\s*\{([^}]+)\}', self.content)
    if cf:
        spec['font_declaration'] = cf.group(1)

    for name, key in [('abovecaptionskip', 'above_skip'), ('belowcaptionskip', 'below_skip')]:
        m = re.search(BS + name + r'\s*=?\s*([^\\\n]+)', self.content)
        if m:
            spec[key] = m.group(1).strip()
    return spec


def _extract_table_spec(self):
    """提取表格规格"""
    spec = OrderedDict([
        ('header_weight', 'bold'), ('header_size', 'small'), ('header_size_pt', 9),
        ('body_size', 'small'), ('body_size_pt', 9),
        ('caption_position', 'above'), ('rule_style', 'booktabs'),
    ])
    tbl_size = re.search(BS + r'begin\{table\*?\}[^}]*' + BS + r'(small|footnotesize|normalsize)\b', self.content)
    if tbl_size:
        spec['body_size'] = tbl_size.group(1)
        spec['body_size_pt'] = LATEX_SIZE_PT.get(tbl_size.group(1), 9)
    return spec


def _extract_figure_spec(self):
    """提取图片规格"""
    spec = OrderedDict([('caption_position', 'below')])
    fig_size = re.search(BS + r'begin\{figure\*?\}[^}]*' + BS + r'(small|footnotesize|normalsize)\b', self.content)
    if fig_size:
        spec['caption_size'] = fig_size.group(1)
    return spec


def _extract_footnote(self):
    """提取脚注规格"""
    spec = OrderedDict([('font_size', 'footnotesize'), ('font_size_pt', 8)])
    fnt = re.search(BS + r'def' + BS + r'@makefntext[^{]*\{(.{0,400})', self.content, re.DOTALL)
    if fnt:
        body = fnt.group(1)
        if BS + r'footnotesize' in body:
            spec['font_size'] = 'footnotesize'; spec['font_size_pt'] = 8
        elif BS + r'small' in body:
            spec['font_size'] = 'small'; spec['font_size_pt'] = 9
        if 'textsuperscript' in body:
            spec['mark_style'] = 'superscript'
        elif '@thefnmark' in body:
            spec['mark_style'] = 'thefnmark'

    tfn = re.search(BS + r'def' + BS + r'thefootnote\s*\{([^}]+)\}', self.content)
    if tfn:
        spec['numbering_format'] = tfn.group(1)

    fr = re.search(BS + r'footnoterule', self.content)
    spec['has_rule'] = bool(fr)
    return spec


def _extract_bibliography(self):
    """提取参考文献规格"""
    spec = OrderedDict([('font_size', 'small'), ('font_size_pt', 9), ('style', 'author-year')])
    if 'natbib' in self.content:
        spec['style'] = 'author-year (natbib)'; spec['natbib'] = True
    if 'biblatex' in self.content:
        spec['style'] = 'biblatex'; spec['biblatex'] = True

    bib_size = re.search(BS + r'thebibliography.*?' + BS + r'(small|footnotesize|normalsize)', self.content, re.DOTALL)
    if bib_size:
        spec['font_size'] = bib_size.group(1)
        spec['font_size_pt'] = LATEX_SIZE_PT.get(bib_size.group(1), 9)

    bbs = re.search(BS + r'bibliographystyle\{([^}]+)\}', self.content)
    if bbs:
        spec['bst_file'] = bbs.group(1)

    bl = re.search(BS + r'def' + BS + r'@?biblabel\s*\{([^}]+)\}', self.content)
    if bl:
        spec['label_format'] = bl.group(1)
    return spec
