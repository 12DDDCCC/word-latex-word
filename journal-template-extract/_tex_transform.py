#!/usr/bin/env python3
r"""
LaTeX模板变换函数

包含:
- _transform_template_tex: 根据模板和排版规格变换template.tex中的内容
- _is_user_cmd: 判断一个命令是否是用户内容命令
- _write_full_spec: 将完整的排版规格写入 .tex 文件
"""

import re
from collections import OrderedDict


def _transform_template_tex(lines, spec, headings, special_envs,
                            doc_class_options, required_packages,
                            document_format):
    """根据模板和排版规格变换template.tex中的内容

    Args:
        lines: 原始template.tex的行列表
        spec: 排版规格字典
        headings: 章节标题格式字典
        special_envs: 特殊环境映射
        doc_class_options: 文档类选项
        required_packages: 必需包列表
        document_format: 文档格式字典

    Returns:
        变换后的行列表
    """
    out = []
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # -- documentclass 选项替换 --
        if stripped.startswith('\\documentclass'):
            if doc_class_options:
                m = re.match(r'(\\documentclass\s*)(\[.*?\])?(\{.*?\})', stripped)
                if m:
                    line = f'{m.group(1)}[{doc_class_options}]{m.group(3)}\n'
            out.append(line)
            i += 1
            continue

        # -- \usepackage 替换/注释 --
        if stripped.startswith('\\usepackage'):
            m = re.match(r'\\usepackage\s*(?:\[([^\]]*)\])?\{([^}]+)\}', stripped)
            if m:
                pkg_names = [p.strip() for p in m.group(2).split(',')]
                opts = m.group(1) or ''
                new_pkgs = []
                for pkg in pkg_names:
                    if pkg in required_packages:
                        new_pkgs.append(pkg)
                    elif pkg in ('amsmath', 'amssymb', 'amsfonts'):
                        new_pkgs.append(pkg)
                    elif pkg in ('graphicx', 'graphics'):
                        new_pkgs.append(pkg)
                    elif pkg in ('hyperref',):
                        new_pkgs.append(pkg)
                    elif pkg in ('natbib',):
                        new_pkgs.append(pkg)
                    elif pkg in ('url',):
                        new_pkgs.append(pkg)
                    else:
                        out.append(f'% REMOVED: {stripped}\n')
                        continue
                if new_pkgs:
                    pkg_str = ', '.join(new_pkgs)
                    if opts:
                        out.append(f'\\usepackage[{opts}]{{{pkg_str}}}\n')
                    else:
                        out.append(f'\\usepackage{{{pkg_str}}}\n')
            else:
                out.append(line)
            i += 1
            continue

        # -- 标题占位符 --
        if '\\title{' in stripped:
            m = re.match(r'(.*\\title\s*)(\[.*?\])?\{.*?\}(.*)', stripped)
            if m:
                prefix = m.group(1)
                short = m.group(2) or ''
                suffix = m.group(3)
                line = f'{prefix}{short}{{TITLE}}{suffix}\n'
            out.append(line)
            i += 1
            continue

        # -- 作者占位符 --
        if stripped.startswith('\\Author') or stripped.startswith('\\author'):
            m = re.match(r'(.*\\(?:Author|author)\s*)(\[.*?\])?\{.*?\}(.*)', stripped)
            if m:
                prefix = m.group(1)
                short = m.group(2) or ''
                suffix = m.group(3)
                line = f'{prefix}{short}{{AUTHOR NAME}}{suffix}\n'
            out.append(line)
            i += 1
            continue

        # -- \affil / \affiliation 占位符 --
        if stripped.startswith('\\affil') or stripped.startswith('\\affiliation'):
            m = re.match(r'(.*\\(?:affil|affiliation)\s*)(\[.*?\])?\{.*?\}(.*)', stripped)
            if m:
                prefix = m.group(1)
                short = m.group(2) or ''
                suffix = m.group(3)
                line = f'{prefix}{short}{{AFFILIATION}}{suffix}\n'
            out.append(line)
            i += 1
            continue

        # -- abstract 环境 --
        if '\\begin{abstract}' in stripped:
            out.append('\\begin{abstract}\n')
            i += 1
            # 跳过原始摘要内容
            while i < n and '\\end{abstract}' not in lines[i]:
                i += 1
            out.append('ABSTRACT CONTENT\n')
            out.append('\\end{abstract}\n')
            i += 1
            continue

        # -- keywords 命令 --
        if '\\keywords{' in stripped or '\\keywords ' in stripped:
            out.append('\\keywords{KEYWORDS}\n')
            i += 1
            continue

        # -- section 命令 -->
        sec_m = re.match(r'\\(section|subsection|subsubsection|paragraph)\{', stripped)
        if sec_m:
            cmd = sec_m.group(1)
            title_m = re.match(r'\\' + cmd + r'\{([^}]*)\}', stripped)
            title = title_m.group(1) if title_m else cmd.upper()
            # 跳过该段内容直到下一个 \section 或空行
            h = headings.get(cmd, {})
            notes = []
            if h.get('size_name'):
                notes.append(f'{h["size_name"]}({h.get("size_pt","?")}pt)')
            if h.get('weight') and h['weight'] != 'normal':
                notes.append(h['weight'])
            note_str = f' ({", ".join(notes)})' if notes else ''
            out.append(f'% -- \\{cmd} {note_str} --\n')
            out.append(f'\\{cmd}{{{title}}}\n')
            i += 1
            continue

        # -- 特殊环境 -->
        for env_name, env_info in special_envs.items():
            if isinstance(env_info, dict) and env_info.get('maps_to'):
                env_cmd = f'\\{env_name}'
                if stripped.startswith(env_cmd):
                    out.append(f'% -- \\{env_name} -> \\{env_info["maps_to"]} --\n')
                    out.append(f'\\{env_name}\n')
                    i += 1
                    break
        else:
            out.append(line)
            i += 1
            continue

    return out


def _is_user_cmd(cmd_name):
    """判断一个命令是否是用户内容命令（如 \\section, \\textbf 等）

    Args:
        cmd_name: LaTeX 命令名（不含反斜杠）

    Returns:
        True 如果是用户内容命令
    """
    # LaTeX 内核命令
    tex_primitives = {
        'begin', 'end', 'documentclass', 'usepackage', 'input', 'include',
        'def', 'gdef', 'edef', 'xdef', 'let', 'newcommand', 'renewcommand',
        'providecommand', 'DeclareRobustCommand', 'setlength', 'addtolength',
        'newlength', 'newcounter', 'setcounter', 'stepcounter', 'addtocounter',
        'value', 'arabic', 'roman', 'Roman', 'alph', 'Alph', 'fnsymbol',
        'if', 'ifx', 'ifnum', 'ifdim', 'ifcat', 'else', 'fi', 'or',
        'relax', 'space', 'par', 'newline', 'hfill', 'vfill', 'hrule', 'vrule',
        'kern', 'hskip', 'vskip', 'hbox', 'vbox', 'vtop',
        'makeatletter', 'makeatother', 'expandafter', 'noexpand',
        'the', 'string', 'csname', 'endcsname', 'catcode', 'char', 'jobname',
    }

    # LaTeX 格式命令（不应在 .tex 正文中出现）
    format_cmds = {
        'pagestyle', 'thispagestyle', 'pagenumbering',
        'headheight', 'headsep', 'footskip',
        'evensidemargin', 'oddsidemargin', 'topmargin',
        'columnsep', 'columnseprule',
        'textheight', 'textwidth', 'paperwidth', 'paperheight',
        'marginparwidth', 'marginparsep', 'marginparpush',
        'floatsep', 'textfloatsep', 'intextsep', 'dblfloatsep',
        'floatpagefraction', 'textfraction', 'topfraction', 'bottomfraction',
        'abovecaptionskip', 'belowcaptionskip',
    }

    if cmd_name in tex_primitives or cmd_name in format_cmds:
        return False

    # 以 @ 开头的命令是内部命令
    if cmd_name.startswith('@'):
        return False

    return True


def _write_full_spec(spec, output_path):
    """将完整的排版规格写入 .tex 文件（注释格式）

    Args:
        spec: 完整排版规格字典
        output_path: 输出 .tex 文件路径
    """
    L = []
    L.append('% ═══════════════════════════════════════════════════════════════')
    L.append('% 自动提取的排版规格（由 journal-template-extract 生成）')
    L.append('% ═══════════════════════════════════════════════════════════════')
    L.append('')

    # 按类别逐个输出
    _write_spec_section(L, '文档类', spec.get('document', {}))
    _write_spec_section(L, '页面布局', spec.get('page_layout', {}))
    _write_spec_section(L, '字体', spec.get('fonts', {}))
    _write_spec_section(L, '正文', spec.get('body_text', {}))
    _write_spec_section(L, '章节标题', spec.get('headings', {}))
    _write_spec_section(L, '摘要', spec.get('abstract', {}))
    _write_spec_section(L, '关键词', spec.get('keywords', {}))
    _write_spec_section(L, '图表', spec.get('figure', {}), spec.get('table', {}))
    _write_spec_section(L, '参考文献', spec.get('bibliography', {}))
    _write_spec_section(L, '脚注', spec.get('footnote', {}))
    _write_spec_section(L, '页眉页脚', spec.get('page_style', {}))
    _write_spec_section(L, '数学', spec.get('math', {}))
    _write_spec_section(L, '超链接', spec.get('hyperref', {}))

    # 高级规格
    advanced = spec.get('advanced', {})
    if advanced:
        L.append('% ─── 高级规格 ───')
        for key, val in advanced.items():
            if isinstance(val, dict):
                L.append(f'% {key}:')
                for k2, v2 in val.items():
                    L.append(f'%   {k2}: {v2}')
            elif isinstance(val, list):
                L.append(f'% {key}: [{", ".join(str(v) for v in val)}]')
            else:
                L.append(f'% {key}: {val}')
        L.append('')

    L.append('% ═══════════════════════════════════════════════════════════════')
    L.append('')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')


def _write_spec_section(L, title, *dicts):
    """辅助函数：写入一个规格区段"""
    merged = OrderedDict()
    for d in dicts:
        if d:
            merged.update(d)
    if not merged:
        return
    L.append(f'% ─── {title} ───')
    for k, v in merged.items():
        if isinstance(v, dict):
            L.append(f'% {k}:')
            for k2, v2 in v.items():
                L.append(f'%   {k2}: {v2}')
        elif isinstance(v, list):
            L.append(f'% {k}: [{", ".join(str(x) for x in v)}]')
        else:
            L.append(f'% {k}: {v}')
    L.append('')
