#!/usr/bin/env python3
r"""
.cls/.cfg/.tex 文件解析器
从LaTeX类文件、配置文件和模板tex中提取期刊信息和论文结构
"""
import re

from _template_scanner import _SKIP_ENVS, _STD_OPTS, _FORMAT_OPTS


def parse_cls_file(cls_path):
    """从.cls文件中提取关键信息"""
    with open(cls_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    info = {
        'class_name': '',
        'version': '',
        'date': '',
        'journals': [],
        'custom_commands': [],
        'custom_environments': [],
        'required_packages': [],
    }

    # 类名和版本
    m = re.search(r'\\ProvidesClass\{(\w+)\}\s*\n?\s*\[([^\]]+)\]', content)
    if m:
        info['class_name'] = m.group(1)
        version_info = m.group(2)
        vm = re.match(r'(\d{4}/\d{2}/\d{2})\s*(.+)', version_info)
        if vm:
            info['date'] = vm.group(1)
            info['version'] = vm.group(2).strip()
        else:
            info['version'] = version_info.strip()

    # 期刊选项 - 只保留真正的期刊缩写
    for m in re.finditer(r'\\DeclareOption\{(\w+)\}', content):
        opt = m.group(1)
        if opt not in _STD_OPTS and opt not in _FORMAT_OPTS and not opt.startswith('cop'):
            info['journals'].append(opt)

    # 自定义命令(用户可用的, 排除内部命令)
    seen = set()
    for m in re.finditer(r'\\newcommand\{\\(\w+)\}', content):
        cmd = m.group(1)
        if not cmd.startswith('@') and not cmd.startswith('cop@') and cmd not in seen:
            info['custom_commands'].append(cmd)
            seen.add(cmd)
    for m in re.finditer(r'\\newcommand\\(\w+)', content):
        cmd = m.group(1)
        if not cmd.startswith('@') and not cmd.startswith('cop@') and cmd not in seen:
            info['custom_commands'].append(cmd)
            seen.add(cmd)
    # \DeclareMathOperator 等也提取
    for m in re.finditer(r'\\DeclareMathOperator\{\\(\w+)\}', content):
        cmd = m.group(1)
        if cmd not in seen:
            info['custom_commands'].append(cmd)
            seen.add(cmd)

    # 自定义环境
    for m in re.finditer(r'\\newenvironment\{(\w+)}', content):
        env = m.group(1)
        if not env.startswith('@') and env not in _SKIP_ENVS:
            info['custom_environments'].append(env)

    # 需要的包
    for m in re.finditer(r'\\RequirePackage\{([^}]+)\}', content):
        pkg = m.group(1).strip()
        info['required_packages'].append(pkg)

    return info


def parse_cfg_file(cfg_path):
    """从.cfg文件中提取期刊配置"""
    with open(cfg_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    journals = {}
    journal_options = []

    # 从DeclareOption中提取期刊缩写(排除格式选项)
    for m in re.finditer(r'\\DeclareOption\{(\w+)\}', content):
        opt = m.group(1)
        # 排除"d"后缀的draft选项(如acpd是acp的draft版)
        base = opt.rstrip('d') if opt.endswith('d') and len(opt) > 3 else opt
        if opt not in _STD_OPTS and opt not in _FORMAT_OPTS and not opt.startswith('cop'):
            journal_options.append(opt)

    # Copernicus风格: \def\cop@journal@ACP@abbreviation{acp}
    for m in re.finditer(r'\\def\\cop@journal@(\w+)@abbreviation\{(\w+)\}', content):
        journals[m.group(2)] = {'long_name_key': m.group(1)}

    for m in re.finditer(r'\\def\\cop@journal@(\w+)@name\{([^}]+)\}', content):
        key = m.group(1)
        name = m.group(2).strip()
        for abbr, data in journals.items():
            if data.get('long_name_key') == key:
                data['name'] = name

    # 通用风格: 直接定义期刊名
    for m in re.finditer(r'\\def\\@journalnameabbreviation\{(\w+)\}', content):
        if m.group(1) not in journals:
            journals[m.group(1)] = {}
    for m in re.finditer(r'\\def\\@journalname\{([^}]+)\}', content):
        name = m.group(1).strip()

    return {'journals': journals, 'journal_options': journal_options}


def parse_template_tex(tex_path):
    """从template.tex中提取论文结构"""
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    structure = {
        'documentclass': '',
        'preamble_extra': [],      # \usepackage 等导言区命令
        'body_sections': [],       # 正文结构命令/环境
    }

    # documentclass
    m = re.search(r'\\documentclass\[([^\]]*)\]\{(\w+)\}', content)
    if m:
        structure['documentclass'] = f'[{m.group(1)}]{{{m.group(2)}}}'
    else:
        m = re.search(r'\\documentclass\{(\w+)\}', content)
        if m:
            structure['documentclass'] = f'{{{m.group(1)}}}'

    # 提取导言区内容(\begin{document}之前)
    doc_begin = content.find('\\begin{document}')
    if doc_begin > 0:
        preamble = content[:doc_begin]
        # 提取\usepackage
        for m in re.finditer(r'\\usepackage(\[([^\]]*)\])?\{([^}]+)\}', preamble):
            options = m.group(2) or ''
            packages = m.group(3).strip()
            structure['preamble_extra'].append(
                f'\\usepackage[{options}]{{{packages}}}' if options else f'\\usepackage{{{packages}}}'
            )

    # 提取正文结构
    in_document = False
    for line in content.split('\n'):
        stripped = line.strip()
        if not in_document:
            if '\\begin{document}' in stripped:
                in_document = True
            continue
        if '\\end{document}' in stripped:
            break

        # 跳过纯注释行(但保留%%开头的模板说明行)
        if stripped.startswith('%') and not stripped.startswith('%%'):
            continue
        if not stripped:
            continue

        # 提取关键结构命令(以\开头, 后跟字母)
        m = re.match(r'\\(\w+)', stripped)
        if m:
            cmd = m.group(1)
            # 论文结构命令(有参数或无参数的都捕获)
            if cmd in ('title', 'Author', 'author', 'affil', 'affiliation',
                       'runningtitle', 'runningauthor', 'runninghead',
                       'correspondence', 'received', 'pubdiscuss', 'revised',
                       'accepted', 'published', 'firstpage', 'maketitle',
                       'copyrightstatement', 'codeavailability',
                       'dataavailability', 'codedataavailability',
                       'sampleavailability', 'videosupplement',
                       'authorcontribution', 'competinginterests', 'disclaimer',
                       'introduction', 'conclusions', 'appendix', 'noappendix',
                       'appendixfigures', 'appendixtables',
                       'section', 'subsection', 'subsubsection',
                       'tableofcontents', 'listoffigures', 'listoftables',
                       'label', 'caption', 'bibliographystyle', 'bibliography',
                       'clearpage', 'newpage', 'tableofcontents',
                       'nohyperlinks', 'keywords'):
                structure['body_sections'].append(cmd)
                continue

        # \begin{...} 环境(行内可能出现, 不一定要行首)
        m = re.search(r'\\begin\{(\w+)\}', stripped)
        if m:
            env = m.group(1)
            if env not in _SKIP_ENVS and env != 'document':
                structure['body_sections'].append(f'begin_{env}')

        # \end{...} 环境
        m = re.search(r'\\end\{(\w+)\}', stripped)
        if m:
            env = m.group(1)
            if env not in _SKIP_ENVS and env not in ('document',):
                structure['body_sections'].append(f'end_{env}')

    return structure
