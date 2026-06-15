#!/usr/bin/env python3
r"""
配置规格提取模块: 14个v3.2普适性增强提取类别

包含方法:
  - _extract_geometry_config
  - _extract_fontspec_config
  - _extract_titlesec_config
  - _extract_caption_package_config
  - _extract_biblatex_config
  - _extract_theorem_envs
  - _extract_algorithm_envs
  - _extract_setspace_config
  - _extract_enumerate_styles
  - _extract_table_detail
  - _extract_marginpar
  - _extract_doi_url_format
  - _extract_cleveref_config
  - _extract_mode_options
"""
import re
from collections import OrderedDict
from shared.latex_parse_utils import BS, _size_style, _len_to_mm


# ═══════════════════════════════════════════════════════════════
# v3.2 新增: 14个普适性增强提取类别
# ═══════════════════════════════════════════════════════════════

def _extract_geometry_config(self):
    """提取geometry包配置"""
    spec = OrderedDict()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{geometry\}', self.content)
    if m:
        opts = m.group(1)
        for key in ('paperwidth', 'paperheight', 'left', 'right', 'top', 'bottom',
                     'width', 'height', 'textwidth', 'textheight',
                     'margin', 'hmargin', 'vmargin', 'includehead', 'includefoot',
                     'headheight', 'headsep', 'footskip', 'columnsep',
                     'twoside', 'asymmetric', 'heightrounded', 'ignorehead',
                     'ignorefoot', 'ignoremp', 'bindingoffset',
                     'landscape', 'portrait', 'a4paper', 'letterpaper',
                     'a5paper', 'b5paper', 'legalpaper', 'executivepaper'):
            km = re.search(key + r'\s*=\s*([^,\s}]+)', opts)
            if km:
                spec[key] = km.group(1).strip()
    gm = re.search(BS + r'geometry\{([^}]+)\}', self.content)
    if gm:
        for key_val in gm.group(1).split(','):
            kv = key_val.strip()
            if '=' in kv:
                k, v = kv.split('=', 1)
                spec[k.strip()] = v.strip()
            elif kv:
                spec[kv] = True
    return spec


def _extract_fontspec_config(self):
    """提取fontspec/XeLaTeX字体配置"""
    spec = OrderedDict()
    m = re.search(BS + r'setmainfont\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content)
    if m:
        spec['main_font'] = m.group(2).strip()
        if m.group(1):
            spec['main_font_options'] = m.group(1).strip()
    m = re.search(BS + r'setsansfont\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content)
    if m:
        spec['sans_font'] = m.group(2).strip()
        if m.group(1):
            spec['sans_font_options'] = m.group(1).strip()
    m = re.search(BS + r'setmonofont\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content)
    if m:
        spec['mono_font'] = m.group(2).strip()
        if m.group(1):
            spec['mono_font_options'] = m.group(1).strip()
    m = re.search(BS + r'setmathfont\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content)
    if m:
        spec['math_font'] = m.group(2).strip()
        if m.group(1):
            spec['math_font_options'] = m.group(1).strip()
    for m in re.finditer(BS + r'newfontfamily\s*' + BS + r'(\w+)\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content):
        spec[f'newfontfamily_{m.group(1)}'] = m.group(3).strip()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{fontspec\}', self.content)
    if m:
        spec['fontspec_options'] = m.group(1).strip()
    for pkg in ('xeCJK', 'ctex', 'CJKutf8'):
        m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{' + pkg + r'\}', self.content)
        if m:
            spec[f'{pkg}_loaded'] = True
            if m.group(1):
                spec[f'{pkg}_options'] = m.group(1).strip()
    for cmd in ('setCJKmainfont', 'setCJKsansfont', 'setCJKmonofont'):
        m = re.search(BS + cmd + r'\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content)
        if m:
            spec[cmd] = m.group(2).strip()
    return spec


def _extract_titlesec_config(self):
    """提取titlesec包章节配置"""
    spec = OrderedDict()
    for level in ('section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'):
        m = re.search(
            BS + r'titleformat\s*\{' + BS + level + r'\}'
            r'(?:\s*\[(\w+)\])?'
            r'\s*\{(.{0,200})\}'
            r'\s*\{(.{0,100})\}'
            r'\s*\{(.{0,50})\}'
            r'\s*\{(.{0,200})\}',
            self.content, re.DOTALL
        )
        if m:
            h = OrderedDict()
            if m.group(1):
                h['shape'] = m.group(1)
            fmt = m.group(2)
            h.update(_size_style(fmt, self.base_size))
            h['label'] = m.group(3).strip()
            h['sep'] = m.group(4).strip()
            if m.group(5).strip():
                h['before_code'] = m.group(5).strip()[:80]
            spec[level] = h
    for level in ('section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph'):
        m = re.search(
            BS + r'titlespacing\s*\*?\s*\{' + BS + level + r'\}'
            r'\s*\{(.{0,50})\}'
            r'\s*\{(.{0,50})\}'
            r'\s*\{(.{0,50})\}',
            self.content, re.DOTALL
        )
        if m:
            if level not in spec:
                spec[level] = OrderedDict()
            spec[level]['titlesec_left'] = m.group(1).strip()
            spec[level]['titlesec_before_sep'] = m.group(2).strip()
            spec[level]['titlesec_after_sep'] = m.group(3).strip()
    m = re.search(BS + r'titlelabel\{([^}]+)\}', self.content)
    if m:
        spec['titlelabel'] = m.group(1).strip()
    if re.search(BS + r'titlesec', self.content):
        spec['titlesec_loaded'] = True
    return spec


def _extract_caption_package_config(self):
    """提取caption包配置"""
    spec = OrderedDict()
    m = re.search(BS + r'captionsetup\s*\{([^}]+)\}', self.content)
    if m:
        opts = m.group(1)
        for key in ('font', 'labelfont', 'labelsep', 'format', 'margin',
                     'width', 'indentation', 'parindent', 'hangindent',
                     'hypcap', 'skip', 'position', 'justification',
                     'singlelinecheck', 'aboveskip', 'belowskip',
                     'tableposition', 'figureposition'):
            km = re.search(key + r'\s*=\s*\{([^}]*)\}', opts)
            if not km:
                km = re.search(key + r'\s*=\s*([^,\s}]+)', opts)
            if km:
                spec[key] = km.group(1).strip()
    for env_type in ('figure', 'table', 'subfigure', 'subtable'):
        m = re.search(BS + r'captionsetup\[' + env_type + r'\]\s*\{([^}]+)\}', self.content)
        if m:
            spec[f'{env_type}_setup'] = m.group(1).strip()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{caption\}', self.content)
    if m:
        spec['package_options'] = m.group(1).strip()
    m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{subcaption\}', self.content)
    if m:
        spec['subcaption_loaded'] = True
        if m.group(1):
            spec['subcaption_options'] = m.group(1).strip()
    return spec


def _extract_biblatex_config(self):
    """提取biblatex配置"""
    spec = OrderedDict()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{biblatex\}', self.content)
    if m:
        opts = m.group(1)
        for key in ('style', 'bibstyle', 'citestyle', 'backend', 'natbib',
                     'sorting', 'sortcites', 'maxnames', 'minnames',
                     'maxbibnames', 'minbibnames', 'maxcitenames', 'mincitenames',
                     'dashed', 'isbn', 'url', 'doi', 'eprint', 'pagetracker',
                     'backref', 'giveninits', 'uniquename', 'uniquelist',
                     'date', 'urldate', 'autocite', 'citereset'):
            km = re.search(key + r'\s*=\s*\{([^}]*)\}', opts)
            if not km:
                km = re.search(key + r'\s*=\s*([^,\s}]+)', opts)
            if km:
                spec[key] = km.group(1).strip()
    for m in re.finditer(BS + r'addbibresource\{([^}]+)\}', self.content):
        spec.setdefault('bibresources', []).append(m.group(1).strip())
    for m in re.finditer(BS + r'DeclareFieldFormat\{(\w+)\}\{([^}]{0,80})', self.content):
        spec[f'fieldformat_{m.group(1)}'] = m.group(2).strip()
    m = re.search(BS + r'usepackage\{biblatex\}', self.content)
    if m and 'style' not in spec:
        spec['loaded'] = True
    return spec


def _extract_theorem_envs(self):
    """提取定理环境"""
    spec = OrderedDict()
    m = re.search(BS + r'theoremstyle\{(\w+)\}', self.content)
    if m:
        spec['theorem_style'] = m.group(1)
    for m in re.finditer(
        BS + r'newtheorem\{(\w+)\}(?:\[(\w+)\])?\{([^}]+)\}(?:\[(\w+)\])?',
        self.content
    ):
        name = m.group(1)
        info = OrderedDict()
        info['display_name'] = m.group(3).strip()
        if m.group(2):
            info['shared_counter'] = m.group(2)
        if m.group(4):
            info['number_within'] = m.group(4)
        spec[name] = info
    for m in re.finditer(
        BS + r'newtheoremstyle\{(\w+)\}'
        r'\s*\{([^}]*)\}'
        r'\s*\{([^}]*)\}'
        r'\s*\{([^}]*)\}'
        r'\s*\{([^}]*)\}'
        r'\s*\{([^}]*)\}'
        r'\s*\{([^}]*)\}',
        self.content
    ):
        name = m.group(1)
        info = OrderedDict()
        info['aboveskip'] = m.group(2).strip()
        info['belowskip'] = m.group(3).strip()
        info['bodyfont'] = m.group(4).strip()
        info['indent'] = m.group(5).strip()
        info['headfont'] = m.group(6).strip()
        info['headpunct'] = m.group(7).strip()
        spec[f'style_{name}'] = info
    for m in re.finditer(
        BS + r'declaretheorem\s*(?:\[([^\]]+)\])?\s*\{(\w+)\}',
        self.content
    ):
        name = m.group(2)
        info = OrderedDict()
        if m.group(1):
            for opt in m.group(1).split(','):
                opt = opt.strip()
                if '=' in opt:
                    k, v = opt.split('=', 1)
                    info[k.strip()] = v.strip()
                else:
                    info[opt] = True
        spec[f'declare_{name}'] = info
    return spec


def _extract_algorithm_envs(self):
    """提取算法环境"""
    spec = OrderedDict()
    for pkg in ('algorithm', 'algorithmic', 'algorithm2e', 'algorithmicx',
                 'algpseudocode', 'algcompatible', 'algpascal', 'algc'):
        m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{' + pkg + r'\}', self.content)
        if m:
            spec[f'{pkg}_loaded'] = True
            if m.group(1):
                spec[f'{pkg}_options'] = m.group(1).strip()
    m = re.search(BS + r'RequirePackage\[([^\]]+)\]\{algorithm2e\}', self.content)
    if m:
        spec['algorithm2e_loaded'] = True
        spec['algorithm2e_options'] = m.group(1).strip()
    m = re.search(BS + r'floatstyle\{([^}]+)\}', self.content)
    if m:
        spec['float_style'] = m.group(1)
    for m in re.finditer(BS + r'renewcommand\{' + BS + r'algorithm(?:name|caption)\w*\}\{([^}]+)\}', self.content):
        spec['caption_format'] = m.group(1).strip()
    if re.search(BS + r'SetAlgoNlRelativeSize' + r'|' + BS + r'SetAlgoNoEnd' + r'|' + BS + r'SetAlgoNoLine', self.content):
        spec['algorithm2e_line_style'] = True
    return spec


def _extract_setspace_config(self):
    """提取行距设置(setspace等)"""
    spec = OrderedDict()
    m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{setspace\}', self.content)
    if m:
        spec['setspace_loaded'] = True
        if m.group(1):
            spec['setspace_options'] = m.group(1).strip()
    if re.search(BS + r'doublespacing\b', self.content):
        spec['default_spacing'] = 'double'
    elif re.search(BS + r'onehalfspacing\b', self.content):
        spec['default_spacing'] = 'onehalf'
    elif re.search(BS + r'singlespacing\b', self.content):
        spec['default_spacing'] = 'single'
    m = re.search(BS + r'setstretch\{([\d.]+)\}', self.content)
    if m:
        spec['stretch'] = float(m.group(1))
    m = re.search(BS + r'linespread\{([\d.]+)\}', self.content)
    if m:
        spec['linespread'] = float(m.group(1))
    m = re.search(BS + r'baselinestretch\s*=?\s*([\d.]+)', self.content)
    if m:
        spec['baselinestretch'] = float(m.group(1))
    m = re.search(BS + r'renewcommand\{' + BS + r'baselinestretch\}\{([\d.]+)\}', self.content)
    if m:
        spec['baselinestretch_renew'] = float(m.group(1))
    m = re.search(BS + r'(?:def|gdef|xdef|edef)' + BS + r'baselinestretch\s*\{([\d.]+)\}', self.content)
    if m:
        spec['baselinestretch_def'] = float(m.group(1))
    return spec


def _extract_enumerate_styles(self):
    """提取enumerate样式"""
    spec = OrderedDict()
    for i, level in enumerate(['i', 'ii', 'iii', 'iv'], 1):
        m = re.search(BS + r'labelenum' + level + r'\s*\{([^}]+)\}', self.content)
        if m:
            spec[f'level{level}_label'] = m.group(1).strip()
        m = re.search(BS + r'def' + BS + r'labelenum' + level + r'\s*\{([^}]+)\}', self.content)
        if m and f'level{level}_label' not in spec:
            spec[f'level{level}_label'] = m.group(1).strip()
    for i, level in enumerate(['i', 'ii', 'iii', 'iv'], 1):
        m = re.search(BS + r'theenum' + level + r'\s*\{([^}]+)\}', self.content)
        if m:
            spec[f'level{level}_counter'] = m.group(1).strip()
        m = re.search(BS + r'def' + BS + r'theenum' + level + r'\s*\{([^}]+)\}', self.content)
        if m and f'level{level}_counter' not in spec:
            spec[f'level{level}_counter'] = m.group(1).strip()
    m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{enumitem\}', self.content)
    if m:
        spec['enumitem_loaded'] = True
        if m.group(1):
            spec['enumitem_options'] = m.group(1).strip()
    for m in re.finditer(BS + r'setlist\s*(?:\[([^\]]*)\])?\s*\{([^}]+)\}', self.content):
        key = m.group(1) or 'default'
        spec[f'setlist_{key}'] = m.group(2).strip()
    m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{enumerate\}', self.content)
    if m:
        spec['enumerate_pkg_loaded'] = True
    return spec


def _extract_table_detail(self):
    """提取表格详细格式"""
    spec = OrderedDict()
    m = re.search(BS + r'tabcolsep\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))', self.content)
    if m:
        spec['tabcolsep'] = m.group(1).strip()
    m = re.search(BS + r'setlength\{' + BS + r'tabcolsep\}\{([^}]+)\}', self.content)
    if m and 'tabcolsep' not in spec:
        spec['tabcolsep'] = m.group(1).strip()
    m = re.search(BS + r'arraystretch\s*=?\s*([\d.]+)', self.content)
    if m:
        spec['arraystretch'] = float(m.group(1))
    m = re.search(BS + r'renewcommand\{' + BS + r'arraystretch\}\{([\d.]+)\}', self.content)
    if m:
        spec['arraystretch_renew'] = float(m.group(1))
    m = re.search(BS + r'setlength\{' + BS + r'extrarowheight\}\{([^}]+)\}', self.content)
    if m:
        spec['extrarowheight'] = m.group(1).strip()
    for name in ('abovetopsep', 'belowbottomsep', 'aboverulesep', 'belowrulesep',
                  'cmidrulekern', 'heavyrulewidth', 'lightrulewidth'):
        m = re.search(BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            spec[f'booktabs_{name}'] = m.group(1).strip()
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:pt|mm|cm|em|ex))', self.content)
        if m and f'booktabs_{name}' not in spec:
            spec[f'booktabs_{name}'] = m.group(1).strip()
    for name in ('textfraction', 'topfraction', 'bottomfraction',
                  'floatpagefraction', 'topnumber', 'bottomnumber', 'totalnumber'):
        m = re.search(BS + r'setcounter\{' + name + r'\}\{(\d+)\}', self.content)
        if m:
            spec[f'counter_{name}'] = int(m.group(1))
    for cmd in ('tophline', 'middlehline', 'bottomhline', 'hhline', 'midrule', 'cmidrule'):
        m = re.search(BS + r'(?:newcommand|def)' + BS + cmd + r'\s*(?:\[[^\]]*\])?\s*\{(.{0,80})', self.content)
        if m:
            spec[f'custom_rule_{cmd}'] = m.group(1).strip()[:60]
    for pkg in ('tabularx', 'longtable', 'supertabular', 'xtab', 'ltablex', 'ltxtable'):
        if re.search(BS + r'(?:RequirePackage|usepackage)\{?' + pkg, self.content):
            spec[f'{pkg}_loaded'] = True
    return spec


def _extract_marginpar(self):
    """提取边注设置"""
    spec = OrderedDict()
    for name in ('marginparwidth', 'marginparsep', 'marginparpush'):
        m = re.search(BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            spec[name] = _len_to_mm(m.group(1).strip())
        else:
            m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))', self.content)
            if m:
                spec[name] = _len_to_mm(m.group(1).strip())
    if re.search(BS + r'reversemarginpar\b', self.content):
        spec['reverse'] = True
    m = re.search(BS + r'usepackage\{marginnote\}', self.content)
    if m:
        spec['marginnote_pkg'] = True
    return spec


def _extract_doi_url_format(self):
    """提取DOI/URL格式"""
    spec = OrderedDict()
    for cmd in ('doi', 'url', 'href'):
        m = re.search(BS + r'(?:newcommand|def|DeclareRobustCommand)' + BS + cmd + r'\s*(?:\[[^\]]*\])?\s*\{(.{0,150})', self.content)
        if m:
            spec[f'{cmd}_definition'] = m.group(1).strip()[:100]
    m = re.search(BS + r'doiurl\s*\{([^}]*)\}', self.content)
    if m:
        spec['doiurl_prefix'] = m.group(1).strip()
    for cmd in ('doitext', 'urlprefix'):
        m = re.search(BS + cmd + r'\s*\{([^}]*)\}', self.content)
        if m:
            spec[cmd] = m.group(1).strip()
    m = re.search(BS + r'bibentry\s*\{([^}]*)\}', self.content)
    if m:
        spec['bibentry'] = m.group(1).strip()
    for key in ('urlcolor', 'citecolor', 'linkcolor'):
        m = re.search(BS + key + r'\s*=\s*\{?(\w+)\}?', self.content)
        if m:
            spec[f'hyperref_{key}'] = m.group(1)
    return spec


def _extract_cleveref_config(self):
    """提取cleveref交叉引用配置"""
    spec = OrderedDict()
    m = re.search(BS + r'usepackage\[([^\]]+)\]\{cleveref\}', self.content)
    if m:
        opts = m.group(1)
        for key in ('capitalise', 'nameinlink', 'noabbrev', 'compress',
                     'sort', 'sort&compress', 'roman', 'subsection',
                     'position', 'backend'):
            if re.search(key, opts):
                spec[key] = True
    m = re.search(BS + r'usepackage\{cleveref\}', self.content)
    if m and 'capitalise' not in spec:
        spec['loaded'] = True
    for m in re.finditer(BS + r'Cre?frefname\{(\w+)\}\{([^}]+)\}\{([^}]+)\}', self.content):
        spec[f'refname_{m.group(1)}'] = {'singular': m.group(2).strip(), 'plural': m.group(3).strip()}
    for m in re.finditer(BS + r'creflabelformat\{(\w+)\}\{([^}]{0,100})', self.content):
        spec[f'labelformat_{m.group(1)}'] = m.group(2).strip()
    return spec


def _extract_mode_options(self):
    """提取文档类模式/选项"""
    spec = OrderedDict()
    declared_options = []
    for m in re.finditer(BS + r'DeclareOption\{(\w+)\}', self.content):
        declared_options.append(m.group(1))
    for m in re.finditer(BS + r'DeclareOption\*\{', self.content):
        declared_options.append('*')
    if declared_options:
        spec['declared_options'] = declared_options
    m = re.search(BS + r'ExecuteOptions\{([^}]+)\}', self.content)
    if m:
        spec['default_options'] = [o.strip() for o in m.group(1).split(',')]
    for m in re.finditer(BS + r'OptionNotUsed', self.content):
        pass
    for mode in ('draft', 'final', 'manuscript', 'preprint', 'review',
                  'twoside', 'oneside', 'onecolumn', 'twocolumn',
                  'titlepage', 'notitlepage', 'openright', 'openany',
                  'leqno', 'fleqn', 'openbib'):
        if re.search(BS + r'DeclareOption\{' + mode + r'\}', self.content):
            spec[f'option_{mode}'] = True
    journal_opts = []
    for m in re.finditer(BS + r'DeclareOption\{(\w+)\}' + r'(.{0,200})', self.content):
        opt_name = m.group(1)
        body = m.group(2)
        if re.search(BS + r'journalname|journalurl|@journal', body):
            journal_opts.append(opt_name)
    if journal_opts:
        spec['journal_options'] = journal_opts
    for m in re.finditer(BS + r'PassOptionsToClass\{([^}]+)\}\{([^}]+)\}', self.content):
        spec.setdefault('pass_to_class', []).append({
            'options': m.group(1).strip(),
            'class': m.group(2).strip()
        })
    m = re.search(BS + r'LoadClass(?:\[([^\]]*)\])?\{(\w+)\}', self.content)
    if m:
        spec['base_class'] = m.group(2)
        if m.group(1):
            spec['base_class_options'] = m.group(1).strip()
    return spec
