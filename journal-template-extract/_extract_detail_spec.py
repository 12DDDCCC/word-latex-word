#!/usr/bin/env python3
r"""
详细规格提取模块: 编号/间距/页眉页脚/颜色/列表/特殊环境/自定义命令/包/浮动
包含方法:
  - _extract_numbering
  - _extract_spacing
  - _extract_header_footer
  - _extract_colors
  - _extract_lists
  - _extract_special_envs
  - _extract_custom_commands
  - _extract_packages
  - _extract_float_settings
"""
import re
from collections import OrderedDict
from pathlib import Path
import sys
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.latex_parse_utils import BS, _size_style


def _extract_numbering(self):
    """提取编号格式"""
    spec = OrderedDict()
    for name in ('section', 'subsection', 'subsubsection', 'paragraph',
                  'figure', 'table', 'equation', 'footnote',
                  'reaction', 'algorithm', 'listing', 'subfigure', 'subtable',
                  'scheme', 'plate', 'box'):
        m = re.search(BS + r'def' + BS + r'the' + name + r'\s*\{([^}]+)\}', self.content)
        if m:
            spec[f'{name}_format'] = m.group(1)
        # \renewcommand{\theX}
        m = re.search(BS + r'renewcommand\{' + BS + r'the' + name + r'\}\{([^}]+)\}', self.content)
        if m and f'{name}_format' not in spec:
            spec[f'{name}_format'] = m.group(1)

    for m in re.finditer(BS + r'numberwithin\{(\w+)\}\{(\w+)\}', self.content):
        spec[f'{m.group(1)}_within'] = m.group(2)

    # \@addtoreset{counter}{section}
    for m in re.finditer(BS + r'@addtoreset\{(\w+)\}\{(\w+)\}', self.content):
        key = f'{m.group(1)}_reset_by'
        if key not in spec:
            spec[key] = m.group(2)

    snd = re.search(BS + r'secnumdepth\s*=?\s*(\d+)', self.content)
    if snd: spec['secnumdepth'] = int(snd.group(1))

    td = re.search(BS + r'tocdepth\s*=?\s*(\d+)', self.content)
    if td: spec['tocdepth'] = int(td.group(1))

    sep = re.search(BS + r'the(?!page|footnote)\w+\s*\{' + BS + r'the\w+\.([^}]*)\}', self.content)
    if sep: spec['number_separator'] = sep.group(1)
    return spec


def _extract_spacing(self):
    """提取间距设置"""
    spec = OrderedDict()
    for name in ('baselinestretch', 'linespread', 'parskip', 'parindent',
                  'abovedisplayskip', 'belowdisplayskip', 'abovedisplayshortskip',
                  'belowdisplayshortskip', 'floatsep', 'textfloatsep',
                  'intextsep', 'dbltextfloatsep', 'dblfloatsep'):
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\w*)', self.content)
        if m: spec[name] = m.group(1).strip()

    for m in re.finditer(BS + r'setlength\{' + BS + r'(\w+)\}\{([^}]+)\}', self.content):
        name = m.group(1)
        if name in ('parskip', 'parindent', 'abovedisplayskip', 'belowdisplayskip',
                    'floatsep', 'textfloatsep', 'intextsep', 'itemsep', 'labelsep'):
            spec[name] = m.group(2).strip()
    return spec


def _extract_header_footer(self):
    """提取页眉页脚"""
    spec = OrderedDict()
    ps = re.search(BS + r'pagestyle\{(\w+)\}', self.content)
    if ps: spec['pagestyle'] = ps.group(1)

    for m in re.finditer(BS + r'fancyhead\[(\w)\]\{([^}]*)\}', self.content):
        pos = {'L': 'left', 'C': 'center', 'R': 'right'}.get(m.group(1), m.group(1))
        spec.setdefault('header', OrderedDict())[pos] = m.group(2)

    for m in re.finditer(BS + r'fancyfoot\[(\w)\]\{([^}]*)\}', self.content):
        pos = {'L': 'left', 'C': 'center', 'R': 'right'}.get(m.group(1), m.group(1))
        spec.setdefault('footer', OrderedDict())[pos] = m.group(2)

    for m in re.finditer(BS + r'def' + BS + r'ps@(\w+)\b(.{0,500})', self.content, re.DOTALL):
        style_name = m.group(1)
        body = m.group(2)
        items = OrderedDict()
        for sub in re.finditer(BS + r'(let' + BS + r'@)evenhead\s*=?\s*([^\\\n]+)', body):
            items['even_head'] = sub.group(2).strip()
        for sub in re.finditer(BS + r'(let' + BS + r'@)oddhead\s*=?\s*([^\\\n]+)', body):
            items['odd_head'] = sub.group(2).strip()
        for sub in re.finditer(BS + r'(let' + BS + r'@)evenfoot\s*=?\s*([^\\\n]+)', body):
            items['even_foot'] = sub.group(2).strip()
        for sub in re.finditer(BS + r'(let' + BS + r'@)oddfoot\s*=?\s*([^\\\n]+)', body):
            items['odd_foot'] = sub.group(2).strip()
        if items: spec[style_name] = items

    pfn = re.search(BS + r'def' + BS + r'thepage\s*\{([^}]+)\}', self.content)
    if pfn: spec['page_number_format'] = pfn.group(1)
    return spec


def _extract_colors(self):
    """提取颜色定义"""
    colors = OrderedDict()
    for m in re.finditer(BS + r'definecolor\{(\w+)\}\{([^}]*)\}\{([^}]*)\}', self.content):
        colors[m.group(1)] = {'model': m.group(2), 'value': m.group(3)}
    cm = re.search(BS + r'colormark\s*=?\s*(true|false|1|0)', self.content, re.IGNORECASE)
    if cm: colors['colormark'] = cm.group(1).lower() in ('true', '1')
    for name in ('urlcolor', 'linkcolor', 'citecolor', 'menucolor', 'filecolor'):
        m = re.search(BS + name + r'\s*=\s*\{?(\w+)\}?', self.content)
        if m: colors[name] = m.group(1)
    return colors


def _extract_lists(self):
    """提取列表样式"""
    spec = OrderedDict()
    for m in re.finditer(BS + r'setlength\{' + BS + r'(itemsep|labelsep|leftmargin|rightmargin|listparindent|topsep)\}\{([^}]+)\}', self.content):
        spec[m.group(1)] = m.group(2)
    li = re.search(BS + r'def' + BS + r'labelitemi\s*\{([^}]+)\}', self.content)
    if li: spec['bullet_style'] = li.group(1)
    ei = re.search(BS + r'def' + BS + r'labelenumi\s*\{([^}]+)\}', self.content)
    if ei: spec['enumerate_style'] = ei.group(1)
    if 'enumitem' in self.content: spec['enumitem_package'] = True
    return spec


def _extract_special_envs(self):
    """提取特殊环境"""
    envs = OrderedDict()
    env_names = ['acknowledgements', 'acknowledgment', 'appendix', 'dataavailability',
                  'codeavailability', 'authorcontribution', 'competinginterests',
                  'supplement', 'supplementary', 'correspondence', 'introduction',
                  'conclusions', 'methods', 'results', 'discussion']

    for env_name in env_names:
        # \newenvironment
        pat_new = (BS + r'newenvironment\{' + env_name + r'\}'
                   r'\s*(?:\[\d?\])?\s*\{(.{0,300})')
        for m in re.finditer(pat_new, self.content, re.DOTALL):
            body = m.group(1)
            e = _size_style(body, self.base_size)
            if BS + r'section' in body or BS + r'subsection' in body:
                e['maps_to'] = 'section' if BS + r'section' in body else 'subsection'
            if e:
                envs[env_name] = e
                break

        # \renewenvironment
        pat_renew = (BS + r'renewenvironment\{' + env_name + r'\}'
                     r'\s*(?:\[\d?\])?\s*\{(.{0,300})')
        for m in re.finditer(pat_renew, self.content, re.DOTALL):
            body = m.group(1)
            e = _size_style(body, self.base_size)
            if e:
                envs[env_name] = e
                break

        # \def\envname{\section{...}}
        pat_def = BS + r'def' + BS + env_name + r'\s*\{(.{0,300})'
        m = re.search(pat_def, self.content, re.DOTALL)
        if m:
            body = m.group(1)
            e = OrderedDict()
            if BS + r'section' in body:
                e['maps_to'] = 'section'
                sec_title = re.search(BS + r'section\{([^}]*)\}', body)
                if sec_title: e['section_title'] = sec_title.group(1)
            elif BS + r'subsection' in body:
                e['maps_to'] = 'subsection'
            elif BS + r'paragraph' in body:
                e['maps_to'] = 'paragraph'
            e.update(_size_style(body, self.base_size))
            envs[env_name] = e

        # \newcommand{\envname}{...}
        pat_cmd = BS + r'newcommand\{?' + BS + env_name + r'\}?\s*(?:\[\d\])?\s*\{(.{0,300})'
        m = re.search(pat_cmd, self.content, re.DOTALL)
        if m and env_name not in envs:
            body = m.group(1)
            e = OrderedDict()
            if BS + r'section' in body:
                e['maps_to'] = 'section'
            elif BS + r'subsection' in body:
                e['maps_to'] = 'subsection'
            e.update(_size_style(body, self.base_size))
            envs[env_name] = e

    # \newenvironment 扫描所有自定义环境
    for m in re.finditer(BS + r'newenvironment\{(\w+)\}', self.content):
        name = m.group(1)
        std_envs = ('document', 'abstract', 'itemize', 'enumerate',
                    'description', 'figure', 'table', 'equation',
                    'tabular', 'minipage', 'array', 'thebibliography',
                    'quote', 'quotation', 'verse', 'center', 'flushleft',
                    'flushright')
        if name not in envs and name not in std_envs:
            envs[name] = OrderedDict([('defined', True)])
    return envs


def _extract_custom_commands(self):
    """提取自定义命令"""
    cmds = OrderedDict()
    for m in re.finditer(BS + r'newcommand\{?' + BS + r'(\w+)\}?\s*(?:\[(\d)\])?\s*\{(.{0,100})', self.content):
        name = m.group(1)
        nargs = m.group(2) or '0'
        definition = m.group(3).strip()
        if name.startswith('@') or name in ('bfseries', 'itshape', 'rmfamily', 'sffamily'):
            continue
        cmds[name] = {'nargs': int(nargs), 'definition_preview': definition}

    for m in re.finditer(BS + r'def' + BS + r'(\w+)\s*(?:#\d)?\s*\{(.{0,100})', self.content):
        name = m.group(1)
        if name.startswith('@') or name in cmds or len(name) < 3:
            continue
        cmds[name] = {'definition_preview': m.group(2).strip()}

    for name in ('Author', 'affil', 'runningtitle', 'runningauthor',
                  'correspondence', 'received', 'accepted', 'published',
                  'firstpage', 'lastpage'):
        if re.search(BS + r'def' + BS + name + r'\b', self.content) or re.search(BS + r'newcommand\{?' + BS + name + r'\}?\b', self.content):
            if name not in cmds:
                cmds[name] = {'defined_in_template': True}
    return cmds


def _extract_packages(self):
    """提取依赖宏包"""
    pkgs = []
    for m in re.finditer(BS + r'RequirePackage(?:\[.*?\])?\{([^}]+)\}', self.content):
        for pkg in m.group(1).split(','):
            pkg = pkg.strip()
            if pkg and pkg not in pkgs: pkgs.append(pkg)
    for m in re.finditer(BS + r'usepackage(?:\[.*?\])?\{([^}]+)\}', self.content):
        for pkg in m.group(1).split(','):
            pkg = pkg.strip()
            if pkg and pkg not in pkgs: pkgs.append(pkg)
    return pkgs


def _extract_float_settings(self):
    """提取浮动体设置"""
    spec = OrderedDict()
    # 浮动比例 — 支持 \renewcommand{\name}{val}, \def\name{val}, \name=val
    for name in ('textfraction', 'topfraction', 'bottomfraction',
                  'floatpagefraction', 'dblfloatpagefraction', 'dbltopfraction'):
        # \renewcommand{\name}{val}
        m = re.search(BS + r'renewcommand\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            spec[name] = m.group(1).strip()
            continue
        # \def\name{val}
        m = re.search(BS + r'def' + BS + name + r'\{([^}]+)\}', self.content)
        if m:
            val = m.group(1).strip()
            val = re.sub(r'%.*', '', val).strip()  # 去掉尾部注释
            if val:
                spec[name] = val
                continue
        # \name=val 或 \name\s*=val
        m = re.search(BS + name + r'\s*=?\s*([\d.]+)', self.content)
        if m:
            spec[name] = m.group(1).strip()
    # 浮动计数器
    for name in ('topnumber', 'bottomnumber', 'totalnumber'):
        m = re.search(BS + r'setcounter\{' + name + r'\}\{(\d+)\}', self.content)
        if m:
            spec[name] = int(m.group(1))
    # 浮动间距
    for name in ('textfloatsep', 'floatsep', 'intextsep',
                  'dbltextfloatsep', 'dblfloatsep'):
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))', self.content)
        if m:
            spec[name] = m.group(1).strip()
        else:
            m = re.search(BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}', self.content)
            if m:
                spec[name] = m.group(1).strip()
    return spec
