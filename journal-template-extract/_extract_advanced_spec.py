#!/usr/bin/env python3
r"""
高级规格提取模块: 14个v3.1提取方法

包含方法:
  - _extract_math_display_skip
  - _extract_footnote_detail
  - _extract_hyperref
  - _extract_global_typography
  - _extract_title_page_layout
  - _extract_date_declarations
  - _extract_abstract_detail
  - _extract_keywords
  - _extract_author_detail
  - _extract_custom_typo_commands
  - _extract_subfig_settings
  - _extract_name_redefinitions
  - _extract_page_style_detail
  - _extract_geometry_config

v3.2方法已移至 _extract_config_spec.py，通过 re-export 保持兼容
"""
import re
from collections import OrderedDict
from shared.latex_parse_utils import BS, _la_size_to_pt, _size_style, _len_to_mm

# Re-export v3.2 配置规格提取方法（保持向后兼容）
from _extract_config_spec import (
    _extract_geometry_config,
    _extract_fontspec_config,
    _extract_titlesec_config,
    _extract_caption_package_config,
    _extract_biblatex_config,
    _extract_theorem_envs,
    _extract_algorithm_envs,
    _extract_setspace_config,
    _extract_enumerate_styles,
    _extract_table_detail,
    _extract_marginpar,
    _extract_doi_url_format,
    _extract_cleveref_config,
    _extract_mode_options,
)


# ═══════════════════════════════════════════════════════════════
# v3.1 新增: 14个遗漏的提取类别
# ═══════════════════════════════════════════════════════════════

def _extract_math_display_skip(self):
    """提取数学间距+amsmath选项"""
    spec = OrderedDict()
    for name in ('abovedisplayskip', 'belowdisplayskip',
                  'abovedisplayshortskip', 'belowdisplayshortskip'):
        patterns = [
            BS + name + r'\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex)(?:\s*[+-]\s*[\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))*)',
            BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}',
        ]
        for pat in patterns:
            m = re.search(pat, self.content)
            if m:
                spec[name] = m.group(1).strip()
                break
    m = re.search(BS + r'mathindent\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))', self.content)
    if m:
        spec['mathindent'] = m.group(1).strip()
    m = re.search(BS + r'RequirePackage\[([^\]]+)\]\{amsmath\}', self.content)
    if m:
        spec['amsmath_options'] = m.group(1).strip()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{amsmath\}', self.content)
    if m:
        spec['amsmath_options'] = m.group(1).strip()
    return spec


def _extract_footnote_detail(self):
    """提取脚注详细格式"""
    spec = OrderedDict()
    patterns = [
        BS + r'long' + BS + r'def' + BS + r'@makefntext\s*#\d\s*\{(.{0,800})',
        BS + r'def' + BS + r'@makefntext\s*#\d\s*\{(.{0,800})',
        BS + r'renewcommand' + BS + r'@makefntext\s*\[\d\]\s*\{(.{0,800})',
    ]
    for pat in patterns:
        m = re.search(pat, self.content, re.DOTALL)
        if m:
            body = m.group(1)
            hm = re.search(r'to\s+([\d.]+(?:em|ex|pt|mm))', body)
            if hm:
                spec['mark_width'] = hm.group(1)
            pm = re.search(BS + r'parindent\s*([\d.]+(?:em|ex|pt|mm))', body)
            if pm:
                spec['text_indent'] = pm.group(1)
            pm2 = re.search(BS + r'parindent\s*=\s*([\d.]+(?:em|ex|pt|mm))', body)
            if pm2 and 'text_indent' not in spec:
                spec['text_indent'] = pm2.group(1)
            if BS + r'hss' in body:
                spec['mark_alignment'] = 'right'
            break
    m = re.search(BS + r'renewcommand\{' + BS + r'thefootnote\}\{([^}]+)\}', self.content)
    if m:
        spec['numbering_renew'] = m.group(1)
    m = re.search(BS + r'def' + BS + r'thefootnote\s*\{([^}]+)\}', self.content)
    if m and 'numbering_renew' not in spec:
        spec['numbering_renew'] = m.group(1)
    return spec


def _extract_hyperref(self):
    """提取hyperref配置"""
    spec = OrderedDict()
    m = re.search(BS + r'hypersetup\{([^}]+)\}', self.content, re.DOTALL)
    if not m:
        m = re.search(BS + r'hypersetup\{([^}]+)\}', self.content, re.DOTALL)
    if m:
        setup = m.group(1)
        for key in ('linkcolor', 'citecolor', 'urlcolor', 'filecolor',
                     'menucolor', 'pagecolor', 'anchorcolor',
                     'colorlinks', 'breaklinks', 'linktocpage',
                     'unicode', 'plainpages', 'draft', 'final',
                     'bookmarks', 'bookmarksopen', 'bookmarksnumbered',
                     'pdfencoding', 'pdfpagelabels',
                     'backref', 'pagebackref', 'hyperindex',
                     'linktoc', 'pageanchor', 'raiselinks'):
            km = re.search(key + r'\s*=\s*([^,\s}]+)', setup)
            if km:
                spec[key] = km.group(1).strip()
        for key in ('pdfauthor', 'pdftitle', 'pdfsubject', 'pdfkeywords',
                     'pdfcreator', 'pdfproducer'):
            km = re.search(key + r'\s*=\s*\{([^}]*)\}', setup)
            if km:
                spec[key] = km.group(1).strip()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{hyperref\}', self.content)
    if m:
        for key in ('colorlinks', 'bookmarks', 'pdfencoding', 'unicode', 'draft',
                     'backref', 'pagebackref', 'plainpages', 'linkcolor',
                     'citecolor', 'urlcolor'):
            km = re.search(key + r'\s*=?\s*\{?([^,\s}]*)\}?', m.group(1))
            if km and key not in spec:
                spec[key] = km.group(1).strip() or True
    return spec


def _extract_global_typography(self):
    """提取全局排版设置"""
    spec = OrderedDict()
    if re.search(BS + r'frenchspacing\b', self.content):
        spec['frenchspacing'] = True
    if re.search(BS + r'sloppy\b', self.content):
        spec['sloppy'] = True
    if re.search(BS + r'flushbottom\b', self.content):
        spec['flushbottom'] = True
    elif re.search(BS + r'raggedbottom\b', self.content):
        spec['raggedbottom'] = True
    for m in re.finditer(BS + r'setlength\{' + BS + r'(parskip|parindent)\}\{([^}]+)\}', self.content):
        spec[m.group(1)] = m.group(2).strip()
    m = re.search(BS + r'baselinestretch\s*=?\s*([\d.]+)', self.content)
    if m:
        spec['baselinestretch'] = m.group(1)
    m = re.search(BS + r'linespread\{([\d.]+)\}', self.content)
    if m:
        spec['linespread'] = m.group(1)
    return spec


def _extract_title_page_layout(self):
    """提取标题页布局"""
    spec = OrderedDict()
    mt = re.search(BS + r'def' + BS + r'@maketitle' + r'\b(.{0,6000})', self.content, re.DOTALL)
    if not mt:
        mt = re.search(BS + r'def' + BS + r'maketitle' + r'\b(.{0,6000})', self.content, re.DOTALL)
    if mt:
        body = mt.group(1)
        elements = []
        if BS + r'journalname' in body:
            elements.append({'name': 'journal_name', 'size': 'Large/bfseries', 'description': '期刊名称'})
        if BS + r'journalinfo' in body:
            elements.append({'name': 'journal_info', 'size': 'normalsize', 'description': '期刊信息'})
        if BS + r'doi' in body:
            elements.append({'name': 'doi', 'size': 'normalsize', 'description': 'DOI信息'})
        if BS + r'@received' in body or BS + r'received' in body:
            elements.append({'name': 'date_line', 'size': 'normalsize', 'description': '日期行: Received – Accepted – Published'})
        title_m = re.search(r'(' + BS + r'(?:LARGE|Large|large|huge|Huge)\s*(?:' + BS + r'(?:bfseries|itshape|sffamily)\s*)*)' + BS + r'@?title', body)
        if title_m:
            title_style = _size_style(title_m.group(1), self.base_size)
            elements.append({'name': 'title', 'size': f'{title_style.get("size_name","?")}({title_style.get("size_pt","?")}pt)', 'weight': title_style.get('weight','bold'), 'description': '论文标题'})
        else:
            elements.append({'name': 'title', 'size': 'LARGE(12pt)', 'weight': 'bold', 'description': '论文标题'})
        if BS + r'@author' in body or BS + r'author' in body:
            elements.append({'name': 'author', 'size': 'normalsize', 'description': '作者列表'})
        if BS + r'@affiliation' in body or BS + r'affiliation' in body:
            elements.append({'name': 'affiliation', 'size': 'small+itshape', 'description': '单位'})
        if BS + r'correspondence' in body or 'Correspondence' in body:
            elements.append({'name': 'correspondence', 'size': 'normalsize', 'description': 'Correspondence to: ...'})
        spec['elements'] = elements
        vs = re.findall(BS + r'vspace\{(\d+(?:\.\d+)?(?:pt|mm|em|ex))\}', body)
        if vs:
            spec['vspace_values'] = vs
    if not spec.get('elements'):
        for variant in ('@@maketitlemanuscript', '@@maketitlefinal', '@@maketitle'):
            m2 = re.search(BS + r'def' + BS + re.escape(variant) + r'\b(.{0,6000})', self.content, re.DOTALL)
            if m2:
                body = m2.group(1)
                elements = []
                if BS + r'journalname' in body:
                    elements.append({'name': 'journal_name', 'size': 'Large/bfseries', 'description': '期刊名称'})
                if BS + r'@title' in body:
                    elements.append({'name': 'title', 'size': 'LARGE(12pt)', 'weight': 'bold', 'description': '论文标题'})
                if BS + r'@author' in body:
                    elements.append({'name': 'author', 'size': 'normalsize', 'description': '作者列表'})
                if BS + r'@affiliation' in body:
                    elements.append({'name': 'affiliation', 'size': 'small+itshape', 'description': '单位'})
                if elements:
                    spec['elements'] = elements
                    break
    return spec


def _extract_date_declarations(self):
    """提取日期声明格式"""
    spec = OrderedDict()
    for cmd in ('received', 'pubdiscuss', 'revised', 'accepted', 'published'):
        m = re.search(BS + r'def' + BS + cmd + r'#?\d?\s*\{(.{0,100})', self.content)
        if m:
            spec[cmd] = m.group(1).strip()[:80]
    m = re.search(r'Received.*?Accepted.*?Published', self.content, re.DOTALL)
    if m:
        spec['date_line_pattern'] = 'Received: ... – Accepted: ... – Published: ...'
    if re.search(r'\bendash\b|–', self.content[:5000]):
        spec['separator'] = 'en-dash (–)'
    return spec


def _extract_abstract_detail(self):
    """提取摘要详细格式"""
    spec = OrderedDict()
    for m in re.finditer(
        BS + r'renewenvironment\{abstract\}\s*(?:\[\d?\])?\s*\{(.{0,600})\}\s*\{(.{0,200})\}',
        self.content, re.DOTALL
    ):
        begin_code = m.group(1)
        if BS + r'large' in begin_code:
            spec['label_size'] = 'large'
            spec['label_size_pt'] = _la_size_to_pt('large', self.base_size)
        elif BS + r'Large' in begin_code:
            spec['label_size'] = 'Large'
            spec['label_size_pt'] = _la_size_to_pt('Large', self.base_size)
        elif BS + r'normalsize' in begin_code:
            spec['label_size'] = 'normalsize'
            spec['label_size_pt'] = self.base_size
        if BS + r'bfseries' in begin_code or BS + r'textbf' in begin_code:
            spec['label_weight'] = 'bold'
        if BS + r'parindent' in begin_code:
            pi_m = re.search(BS + r'parindent\s*([\d.]+\w+)', begin_code)
            if pi_m:
                spec['internal_parindent'] = pi_m.group(1)
            elif re.search(BS + r'parindent\s*0\s*pt', begin_code) or re.search(BS + r'parindent=0pt', begin_code):
                spec['internal_parindent'] = '0pt'
        if BS + r'sffamily' in begin_code:
            spec['label_font_family'] = 'sans-serif'
        break  # 取最后一个
    m = re.search(BS + r'def' + BS + r'secondabstract', self.content)
    if m:
        spec['second_abstract'] = True
    m = re.search(BS + r'secabstractname', self.content)
    if m:
        spec['second_abstract_name_cmd'] = True
    return spec


def _extract_keywords(self):
    """提取关键词格式"""
    spec = OrderedDict()
    m = re.search(BS + r'def' + BS + r'keywords\s*#?\d?\s*\{(.{0,300})', self.content, re.DOTALL)
    if m:
        body = m.group(1)
        if BS + r'bfseries' in body or BS + r'textbf' in body:
            spec['prefix_weight'] = 'bold'
        km = re.search(r'Keywords:\s*', body)
        if km:
            spec['prefix_text'] = 'Keywords: '
        else:
            km = re.search(r'Search\s*keywords:\s*', body, re.IGNORECASE)
            if km:
                spec['prefix_text'] = 'Search keywords: '
        if BS + r'textcol' in body or BS + r'color\{textcol\}' in body:
            spec['prefix_color'] = 'textcol'
        vm = re.search(BS + r'vspace\{([^}]+)\}', body)
        if vm:
            spec['prefix_vspace'] = vm.group(1)
        spec.update(_size_style(body, self.base_size))
    m = re.search(BS + r'def' + BS + r'@keywords\s*#?\d?\s*\{(.{0,200})', self.content)
    if m:
        spec['storage_cmd'] = '@keywords'
    if 'prefix_weight' not in spec:
        if re.search(BS + r'textbf\{Keywords', self.content) or re.search(BS + r'bfseries.*Keywords', self.content):
            spec['prefix_weight'] = 'bold'
    return spec


def _extract_author_detail(self):
    """提取作者详细格式"""
    spec = OrderedDict()
    for af in re.finditer(BS + r'(?:renewcommand)?' + BS + r'Affilfont\s*\{([^}]+)\}', self.content):
        body = af.group(1).strip()
        spec['affil_font_declaration'] = body
        if BS + r'itshape' in body:
            spec['affil_italic'] = True
        if BS + r'small' in body:
            spec['affil_size'] = 'small'
            spec['affil_size_pt'] = 9
    for cmd in ('Authsep', 'Authand', 'Authands'):
        m = re.search(BS + cmd + r'\s*\{([^}]*)\}', self.content)
        if m:
            spec[f'{cmd}_value'] = m.group(1).strip()
    if re.search(BS + r'correspondauthor', self.content):
        spec['correspondauthor_mark'] = True
    if re.search(BS + r'equalcontrib', self.content):
        spec['equalcontrib_mark'] = True
    if re.search(BS + r'deco\b', self.content):
        spec['deceased_mark'] = True
    if re.search(BS + r'presentaddress', self.content):
        spec['presentaddress'] = True
    return spec


def _extract_custom_typo_commands(self):
    """提取自定义排版命令"""
    spec = OrderedDict()
    cmd_patterns = [
        (BS + r'(?:newcommand|DeclareRobustCommand|providecommand)\{?' + BS + r'(%s)\}?\s*(?:\[(\d+)\])?\s*\{(.{0,150})'),
        (BS + r'def' + BS + r'(%s)\s*(?:#?\d)?\s*\{(.{0,150})'),
    ]
    for cmd_name in ('ce', 'unit', 'degre', 'permil', 'sun', 'earth'):
        found = False
        for pat_template in cmd_patterns:
            pat = pat_template % cmd_name
            m = re.search(pat, self.content)
            if m:
                groups = m.groups()
                body = groups[-1].strip() if groups else ''
                nargs = groups[1] if len(groups) > 2 else None
                if body:
                    spec[cmd_name] = body
                    if nargs:
                        spec[f'{cmd_name}_nargs'] = int(nargs)
                    found = True
                    break
    return spec


def _extract_subfig_settings(self):
    """提取子图subfig设置"""
    spec = OrderedDict()
    if re.search(BS + r'RequirePackage\{subfig\}', self.content) or re.search(BS + r'usepackage\{subfig\}', self.content):
        spec['subfig_package'] = True
    if re.search(BS + r'RequirePackage\{subfloat\}', self.content) or re.search(BS + r'usepackage\{subfloat\}', self.content):
        spec['subfloat_package'] = True
    for m in re.finditer(BS + r'g@addto@macro' + BS + r'(subfiguresbegin|subtablesbegin|subfiguresend|subtablesend)\s*\{(.{0,200})', self.content):
        spec[m.group(1)] = m.group(2).strip()[:100]
    for name in ('subfigtopskip', 'subfigcapskip', 'subfigbottomskip',
                  'subcapskip', 'sublabelskip',
                  'subfiglabelfont', 'subfigcapfont'):
        for pat in [
            BS + r'renewcommand\{' + BS + name + r'\}\{([^}]+)\}',
            BS + r'def' + BS + name + r'\{([^}]+)\}',
            BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}',
        ]:
            m = re.search(pat, self.content)
            if m:
                spec[name] = m.group(1).strip()
                break
    return spec


def _extract_name_redefinitions(self):
    """提取名称重定义"""
    spec = OrderedDict()
    for name in ('figurename', 'tablename', 'abstractname', 'refname',
                  'appendixname', 'bibname', 'contentsname',
                  'listfigurename', 'listtablename', 'indexname',
                  'equationname', 'sectionname', 'schemename',
                  'platesname', 'listingsname', 'boxesname',
                  'introductionname', 'conclusionname',
                  'secabstractname', 'authorcontributionname',
                  'competinginterestsname', 'copyrightstatementname',
                  'codeavailabilityname', 'dataavailabilityname',
                  'acknowledgementname', 'acknowledgementsname'):
        m = re.search(BS + r'renewcommand\{?' + BS + name + r'\}?\s*\{([^}]+)\}', self.content)
        if m:
            spec[name] = m.group(1).strip()
            continue
        m = re.search(BS + r'def' + BS + name + r'\{([^}]+)\}', self.content)
        if m:
            val = m.group(1).strip()
            val = re.sub(r'%.*', '', val).strip()
            if val:
                spec[name] = val
    return spec


def _extract_page_style_detail(self):
    """提取页面样式详细"""
    spec = OrderedDict()
    for name in ('headrulewidth', 'footrulewidth', 'headwidth'):
        m = re.search(BS + r'renewcommand\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            spec[name] = m.group(1).strip()
    for m in re.finditer(BS + r'fancyhead\[([^\]]+)\]\{([^}]*)\}', self.content):
        positions = m.group(1).strip()
        content = m.group(2).strip()
        spec[f'header_{positions}'] = content
    for m in re.finditer(BS + r'fancyfoot\[([^\]]+)\]\{([^}]*)\}', self.content):
        positions = m.group(1).strip()
        content = m.group(2).strip()
        spec[f'footer_{positions}'] = content
    m = re.search(BS + r'fancyhf\{([^}]*)\}', self.content)
    if m:
        spec['fancyhf'] = m.group(1).strip()
    for m in re.finditer(BS + r'def' + BS + r'ps@(\w+)\s*\{(.{0,3000})', self.content, re.DOTALL):
        style_name = m.group(1)
        body = m.group(2)
        items = OrderedDict()
        for sub in re.finditer(BS + r'def' + BS + r'@oddhead\s*\{([^}]+)\}', body):
            items['odd_head'] = sub.group(1).strip()[:120]
        for sub in re.finditer(BS + r'def' + BS + r'@evenhead\s*\{([^}]+)\}', body):
            items['even_head'] = sub.group(1).strip()[:120]
        for sub in re.finditer(BS + r'def' + BS + r'@oddfoot\s*\{([^}]+)\}', body):
            items['odd_foot'] = sub.group(1).strip()[:120]
        for sub in re.finditer(BS + r'def' + BS + r'@evenfoot\s*\{([^}]+)\}', body):
            items['even_foot'] = sub.group(1).strip()[:120]
        if BS + r'let' + BS + r'@evenhead' + BS + r'@oddhead' in body:
            items['even_head_equals_odd_head'] = True
        if BS + r'let' + BS + r'@evenfoot' + BS + r'@oddfoot' in body:
            items['even_foot_equals_odd_foot'] = True
        if items:
            spec[f'ps@{style_name}'] = items
    return spec