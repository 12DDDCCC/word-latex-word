#!/usr/bin/env python3
"""模板骨架解析/构建/preamble处理

包含：
- parse_template_skeleton: 从模板骨架 tex 中提取格式映射
- clean_preamble: 清理preamble中的伪命令行
- extract_skeleton_commands: 从模板骨架中提取结构命令
- _build_skeleton_from_spec: 从spec和skeleton_info生成骨架tex
- _build_preamble_from_spec: 从spec生成preamble
"""

import re


def normalize_config_mode(config_mode=None):
    """Return the template mode used consistently by TeX and layout probing."""
    if config_mode in (None, '', 'classic'):
        return 'manuscript'
    return config_mode


def class_options_from_spec(spec, journal, config_mode=None):
    """Return class options accepted by the extracted target class."""
    dc = spec.get('document_class', {}) if spec else {}
    class_name = dc.get('class_name', 'article')
    declared = set(dc.get('declared_options', []))
    accepts_any = '*' in declared

    def accepted(option):
        return bool(option) and (accepts_any or option in declared)

    options = list(dc.get('default_options', []))
    if journal != class_name and accepted(journal):
        options.append(journal)
    mode = normalize_config_mode(config_mode)
    if accepted(mode):
        options.append(mode)
    return list(dict.fromkeys(options))


def parse_template_skeleton(template_content):
    """从模板骨架 tex 中提取格式映射

    Returns:
        dict: {
            'metadata_commands':   {cmd_name: full_line, ...},  # e.g. {'title': '\\title{TITLE}', ...}
            'introduction_cmd':    '\\introduction' or None,
            'conclusions_cmd':     '\\conclusions' or None,
            'statement_cmds':      {cmd_name: placeholder_text, ...},  # e.g. {'codeavailability': 'CODE AVAILABILITY TEXT'}
            'ack_env':             'acknowledgements' or 'acknowledgment' or None,
            'bib_style':           'copernicus' or None,
            'abstract_after_maketitle': True/False,
            'metadata_block':      str,  # \begin{document} 到第一个正文之间的全部内容
        }
    """
    doc_begin_match = re.search(r'\\begin\{document\}', template_content)
    doc_end_match = re.search(r'\\end\{document\}', template_content)
    if not doc_begin_match or not doc_end_match:
        return _default_skeleton_info()

    skeleton = template_content[doc_begin_match.end():doc_end_match.start()]

    info = {
        'metadata_commands': {},
        'introduction_cmd': None,
        'conclusions_cmd': None,
        'statement_cmds': {},
        'ack_env': None,
        'bib_style': None,
        'abstract_env': 'abstract',
        'abstract_cmd_optional': '',
        'keywords_cmd': None,
        'bib_filename': 'references',
        'abstract_after_maketitle': True,
        'metadata_block': '',
    }

    # 1. 提取元数据命令
    meta_patterns = [
        (r'(\\title\{[^}]*\})', 'title'),
        (r'(\\runningtitle\{[^}]*\})', 'runningtitle'),
        (r'(\\runningauthor\{[^}]*\})', 'runningauthor'),
        (r'(\\correspondence\{[^}]*\})', 'correspondence'),
        (r'(\\Author\[[^\]]*\]\{[^}]*\}\{[^}]*\})', 'author'),
        (r'(\\affil\[[^\]]*\]\{[^}]*\})', 'affil'),
    ]
    for pattern, name in meta_patterns:
        m = re.search(pattern, skeleton)
        if m:
            info['metadata_commands'][name] = m.group(1)

    # 2. 提取特殊章节命令（\introduction, \conclusions 等）
    #    特征：独占一行的 \command 不带花括号，且不是 \begin/\end/\section/\subsection
    section_cmds = re.findall(r'^\s*(\\[a-zA-Z]+)\s*$', skeleton, re.MULTILINE)
    for cmd in section_cmds:
        cmd_name = cmd[1:].lower()  # 去掉反斜杠
        if cmd_name in ('maketitle', 'clearpage', 'newpage', 'tableofcontents',
                        'listoffigures', 'listoftables', 'appendix', 'bibliography'):
            continue
        if 'intro' in cmd_name:
            info['introduction_cmd'] = cmd
        elif 'conclu' in cmd_name:
            info['conclusions_cmd'] = cmd

    # 3. 提取声明命令（\codeavailability{TEXT} 等）
    #    特征：\commandname{PLACEHOLDER TEXT}
    decl_pattern = r'\\([a-zA-Z]+)\{([^}]*)\}'
    for m in re.finditer(decl_pattern, skeleton):
        cmd_name = m.group(1).lower()
        placeholder = m.group(2).strip()
        # 排除非声明的命令
        non_decls = {'title', 'author', 'affil', 'section', 'subsection', 'subsubsection',
                     'paragraph', 'includegraphics', 'caption', 'label', 'ref', 'cite',
                     'bibliographystyle', 'bibliography', 'runningtitle', 'runningauthor',
                     'correspondence', 'maketitle', 'begin', 'end', 'footnote', 'emph',
                     'textbf', 'textit', 'textrm', 'textsc', 'centering', 'hfill'}
        if cmd_name in non_decls:
            continue
        # 判断是否是声明命令：占位文本是大写英文或较长
        if placeholder and (placeholder.isupper() or len(placeholder) > 10):
            info['statement_cmds'][cmd_name] = placeholder

    # 4. 提取致谢环境名
    ack_match = re.search(r'\\begin\{(acknowledgement?s?)\}', skeleton, re.IGNORECASE)
    if ack_match:
        info['ack_env'] = ack_match.group(1)

    # 5. 提取参考文献样式和文件名
    bib_match = re.search(r'\\bibliographystyle\{([^}]+)\}', skeleton)
    if bib_match:
        info['bib_style'] = bib_match.group(1)
    bib_file_match = re.search(r'\\bibliography\{([^}]+)\}', skeleton)
    if bib_file_match:
        info['bib_filename'] = bib_file_match.group(1)

    # 6. 提取 abstract 环境名（可能不是 'abstract'）
    abs_env_match = re.search(r'\\begin\{(abstract\*?)\}', skeleton, re.IGNORECASE)
    if abs_env_match:
        info['abstract_env'] = abs_env_match.group(1)

    # 7. 提取 keywords 命令（如 \keywords{...}）
    kw_match = re.search(r'\\(keywords)\{', skeleton)
    if kw_match:
        info['keywords_cmd'] = '\\' + kw_match.group(1)

    # 8. 判断摘要位置（在 \maketitle 之前还是之后）
    maketitle_pos = skeleton.find('\\maketitle')
    abstract_pos = skeleton.find('\\begin{abstract}')
    if maketitle_pos >= 0 and abstract_pos >= 0:
        info['abstract_after_maketitle'] = abstract_pos > maketitle_pos

    # 9. 提取元数据区（\begin{document} 到第一个正文内容之间）
    #    找到 \maketitle 后面的位置作为元数据区结束
    if maketitle_pos >= 0:
        # 从 skeleton 开头到 \maketitle 后一行
        end_of_meta = skeleton.find('\n', maketitle_pos)
        if end_of_meta < 0:
            end_of_meta = maketitle_pos + len('\\maketitle')
        info['metadata_block'] = skeleton[:end_of_meta].strip()
    else:
        # 没有 \maketitle，取 \begin{abstract} 之前
        if abstract_pos >= 0:
            info['metadata_block'] = skeleton[:abstract_pos].strip()
        else:
            info['metadata_block'] = skeleton[:200].strip()

    return info


def _default_skeleton_info():
    """返回默认的模板格式信息（无模板时使用标准 LaTeX 格式）"""
    return {
        'metadata_commands': {},
        'introduction_cmd': None,
        'conclusions_cmd': None,
        'statement_cmds': {},
        'ack_env': None,
        'bib_style': 'plain',
        'abstract_env': 'abstract',
        'keywords_cmd': None,
        'bib_filename': 'references',
        'abstract_after_maketitle': True,
        'metadata_block': '',
    }


def _build_skeleton_from_spec(spec, skeleton_info, cls_name, opts, metadata_block):
    """从spec和skeleton_info生成骨架tex"""
    L = []
    L.append(f'\\documentclass[{opts}]{{{cls_name}}}')
    L.append('')
    L.append('\\begin{document}')
    L.append('')
    L.append(metadata_block)

    # abstract
    abs_env = skeleton_info.get('abstract_env', 'abstract')
    L.append(f'\\begin{{{abs_env}}}')
    L.append('ABSTRACT TEXT')
    L.append(f'\\end{{{abs_env}}}')
    L.append('')

    # introduction
    if skeleton_info.get('introduction_cmd'):
        L.append(skeleton_info['introduction_cmd'])
    else:
        L.append('\\section{Introduction}')
    L.append('INTRODUCTION TEXT')
    L.append('')

    # conclusions
    if skeleton_info.get('conclusions_cmd'):
        L.append(skeleton_info['conclusions_cmd'])
    else:
        L.append('\\section{Conclusions}')
    L.append('CONCLUSIONS TEXT')
    L.append('')

    # 声明区
    for cmd_name in skeleton_info.get('statement_cmds', {}):
        L.append(f'\\{cmd_name}{{PLACEHOLDER}}')
        L.append('')

    # acknowledgements
    ack_env = skeleton_info.get('ack_env')
    if ack_env:
        L.append(f'\\begin{{{ack_env}}}')
        L.append('ACKNOWLEDGEMENTS TEXT')
        L.append(f'\\end{{{ack_env}}}')
        L.append('')

    # bibliography
    bib_style = skeleton_info.get('bib_style') or cls_name
    L.append(f'\\bibliographystyle{{{bib_style}}}')
    L.append('\\bibliography{references}')
    L.append('')

    L.append('\\end{document}')
    return '\n'.join(L)


def _build_preamble_from_spec(spec, journal):
    """从spec生成preamble（替代从骨架tex中提取）

    只生成\\documentclass，不添加\\usepackage（cls已通过RequirePackage加载）。
    manuscript选项只在cls声明了该选项时添加。
    """
    dc = spec.get('document_class', {})
    cls_name = dc.get('class_name', 'article')
    opts_list = class_options_from_spec(spec, journal)
    opts_str = ','.join(opts_list)

    lines = []
    option_block = f'[{opts_str}]' if opts_str else ''
    lines.append(f'\\documentclass{option_block}{{{cls_name}}}')
    lines.append('')
    return '\n'.join(lines)


def clean_preamble(preamble):
    """清理preamble中的伪命令行、无效行和自动生成的注释块

    过滤如 "\\usepackage commands included in the copernicus.cls:" 之类的伪命令。
    这些行以 \\ 开头但后面跟的是普通文字而非LaTeX语法。
    保留有效的 \\usepackage[options]{package} 命令。
    同时过滤自动生成的排版规格注释块（以 % === 开头）。
    """
    import re as _re
    # 伪命令特征：\cmd 后跟空格和小写字母，且没有后续的 { 或 [
    # 有效命令如 \usepackage[autolanguage]{numcompress} 应保留
    _pseudo_cmd = _re.compile(
        r'^\s*\\[a-zA-Z]+\s+[a-z]'  # \cmd 后跟空格和小写字母开头
    )
    _valid_with_option = _re.compile(
        r'^\s*\\[a-zA-Z]+\s*\['       # \cmd[option] — 有效
        r'|^\s*\\[a-zA-Z]+\s*\{'       # \cmd{arg} — 有效
    )
    # 自动生成的注释块（排版规格等）
    _auto_comment_block = _re.compile(r'^\s*% [=]{3,}')  # % ===== 开头

    lines = preamble.split('\n')
    cleaned = []
    in_auto_block = False
    for l in lines:
        # 检测自动注释块开始
        if _auto_comment_block.match(l):
            in_auto_block = True
            continue
        # 检测自动注释块结束（空行或非注释行）
        if in_auto_block:
            if l.strip() == '' or not l.strip().startswith('%'):
                in_auto_block = False
            else:
                continue  # 跳过块内注释行
        # 过滤伪命令
        if _pseudo_cmd.match(l) and not _valid_with_option.match(l):
            continue
        cleaned.append(l)
    return '\n'.join(cleaned)


def extract_skeleton_commands(skeleton_body):
    """从模板骨架中提取标题/作者/声明等结构命令

    保留 \\begin{document} 之后、第一个正文 \\section 之前的所有内容。
    只保留有效的LaTeX命令行和空行/注释行，过滤伪命令如
    \\usepackage commands included in the copernicus.cls:
    同时过滤自动生成的排版规格注释块（以 % === 开头）。
    """
    import re as _re
    # 严格匹配：命令后必须紧跟 { 或 [ 或为独立命令（行尾/注释）
    # 排除 "\usepackage commands included..." 这类伪命令
    _valid_line = _re.compile(
        r'^\s*%'                       # 注释行
        r'|^\s*$'                      # 空行
        r'|^\\usepackage\s*[{[]'       # \usepackage{...} 或 \usepackage[...]{...}
        r'|^\\(?:begin|end)\{[^}]+\}'  # \begin{...} / \end{...}
        r'|^\\(?:maketitle|newcommand|renewcommand|providecommand|def|let)\b'
        r'|^\\(?:title|Author|affil|correspondence|runningauthor)\s*[{[]'
        r'|^\\(?:codeavailability|dataavailability|competinginterests)\s*[{[]'
        r'|^\\(?:acknowledgements|bibliographystyle|bibliography|citestyle)\s*[{[]'
        r'|^\\(?:runningtitle|runningauthor|firstpage|lastpage|pubdoi)\s*[{[]'
        r'|^\\(?:input|include|label|ref|cite)\s*[{[]'
        r'|^\\(?:setcounter|setlength|renewenvironment|newenvironment)\s*[{[]'
        r'|^\\(?:pagestyle|thispagestyle|noindent|par|vspace|hspace|smallskip|medskip|bigskip)\b'
    )
    # 自动生成的注释块（排版规格等）
    _auto_comment_block = _re.compile(r'^\s*% [=]{3,}')

    body_start_patterns = ('\\section', '\\subsection', '\\subsubsection',
                           '\\introduction', '\\conclusions', '\\paragraph')

    lines = skeleton_body.split('\n')
    result = []
    in_auto_block = False
    for line in lines:
        stripped = line.strip()
        # 检测自动注释块开始
        if _auto_comment_block.match(stripped):
            in_auto_block = True
            continue
        # 检测自动注释块结束（空行或非注释行）
        if in_auto_block:
            if stripped == '' or not stripped.startswith('%'):
                in_auto_block = False
            else:
                continue  # 跳过块内注释行
        if any(stripped.startswith(p) for p in body_start_patterns):
            break
        if _valid_line.match(stripped):
            result.append(line)

    return '\n'.join(result)
