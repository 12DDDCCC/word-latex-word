"""
精简版LaTeX模板提取器 — 骨架推导与指南生成
包含 derive_skeleton_info(), derive_metadata_block(), 以及 to_guide() 方法
"""
from collections import OrderedDict
from pathlib import Path


class _GuideMixin:
    """混入类：为 TemplateExtractLite 提供 to_guide 方法"""

    def to_guide(self, output_path):
        spec = self.extract_all()
        lines = [f'# {self.journal} 模板使用指南\n']

        # 文档类
        dc = spec.get('document_class', {})
        if dc:
            cls_name = dc.get('class_name', self.journal)
            opts = ','.join(dc.get('default_options', []))
            lines.append(f'## 文档类\n')
            lines.append(f'```latex')
            lines.append(f'\\documentclass[{opts}]{{{cls_name}}}')
            lines.append(f'```\n')

        # 必需包
        pkgs = spec.get('required_packages', {})
        if pkgs:
            lines.append(f'## 必需包\n```latex')
            for pkg, opts in pkgs.items():
                if opts is True:
                    lines.append(f'\\usepackage{{{pkg}}}')
                else:
                    lines.append(f'\\usepackage[{opts}]{{{pkg}}}')
            lines.append('```\n')

        # 标题
        tf = spec.get('title_format', {})
        if tf:
            lines.append(f'## 标题\n')
            args = tf.get('title_args', 0)
            if tf.get('has_short_title') or args > 0:
                lines.append('```latex\\title[短标题]{标题}```\n')
            else:
                lines.append('```latex\\title{标题}```\n')

        # 摘要
        af = spec.get('abstract_format', {})
        if af:
            lines.append(f'## 摘要\n')
            if af.get('type') == 'command':
                lines.append(f'```latex\\abstract{{摘要内容}}```\n')
            else:
                lines.append(f'```latex\n\\begin{{abstract}}\n摘要内容\n\\end{{abstract}}\n```\n')

        # 关键词
        kf = spec.get('keywords_format', {})
        if kf:
            lines.append('## 关键词\n```latex\\keywords{关键词1, 关键词2}```\n')

        # 特殊环境
        se = spec.get('special_envs', {})
        if se:
            lines.append(f'## 声明段落\n')
            for name, detail in se.items():
                tp = detail.get('type', 'environment')
                maps = detail.get('maps_to', '')
                title = detail.get('section_title', '')
                if tp == 'environment':
                    lines.append(f'- `\\begin{{{name}}}...\\end{{{name}}}`' + (f' → {maps}' if maps else ''))
                else:
                    lines.append(f'- `\\{name}{{...}}`' + (f' → {maps}' if maps else ''))
                    if title:
                        lines.append(f'  默认标题: {title}')
            lines.append('')

        # 参考文献
        bf = spec.get('bibliography_format', {})
        if bf:
            lines.append(f'## 参考文献\n')
            if bf.get('style') == 'natbib':
                opts = bf.get("natbib_options", "")
                lines.append(f'```latex\n\\usepackage[{opts}]{{natbib}}\n\\bibliography{{references}}\n```\n')
            elif bf.get('style') == 'biblatex':
                opts = bf.get("biblatex_options", "")
                lines.append(f'```latex\n\\usepackage[{opts}]{{biblatex}}\n\\printbibliography\n```\n')

        # 图表
        ff = spec.get('figure_format', {})
        if ff:
            lines.append(f'## 图\n')
            pos = ff.get('default_position', 'htbp')
            lines.append(f'```latex\n\\begin{{figure}}[{pos}]\n  \\includegraphics{{...}}\n  \\caption{{...}}\n\\end{{figure}}\n```\n')

        # 模板特有命令
        ts = spec.get('template_specific', {})
        req = ts.get('required_declarations', [])
        if req:
            lines.append(f'## 必填声明\n')
            for d in req:
                lines.append(f'- {d}')

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))


def derive_skeleton_info(spec, cls_path=None):
    """从spec字典推导skeleton_info（替代骨架tex解析）

    Args:
        spec: extract_spec()返回的spec字典
        cls_path: .cls文件路径（用于补充信息）

    Returns:
        dict: skeleton_info字典，与parse_template_skeleton()返回值同构
    """
    info = {
        'metadata_commands': {},
        'introduction_cmd': None,
        'conclusions_cmd': None,
        'statement_cmds': {},
        'ack_env': None,
        'bib_style': None,
        'abstract_env': 'abstract',
        'keywords_cmd': None,
        'bib_filename': 'references',
        'abstract_after_maketitle': True,
        'metadata_block': '',
    }

    # introduction_cmd / conclusions_cmd: special_envs中type=command
    for cmd_name in ('introduction', 'conclusions'):
        env = spec.get('special_envs', {}).get(cmd_name, {})
        if env and env.get('type') == 'command':
            info[f'{cmd_name}_cmd'] = f'\\{cmd_name}'

    # statement_cmds: template_specific.required_declarations
    decls = spec.get('template_specific', {}).get('required_declarations', [])
    for cmd_name in decls:
        info['statement_cmds'][cmd_name] = ''

    # ack_env: special_envs中acknowledgements
    for ack_name in ('acknowledgements', 'acknowledgment'):
        ack_info = spec.get('special_envs', {}).get(ack_name, {})
        if ack_info:
            info['ack_env'] = ack_name
            break

    # bib_style: bibliography_format.bst_file 或 class_name
    bib_fmt = spec.get('bibliography_format', {})
    bst = bib_fmt.get('bst_file', '')
    if bst:
        info['bib_style'] = bst
    elif cls_path:
        info['bib_style'] = Path(cls_path).stem

    # abstract_env
    abs_fmt = spec.get('abstract_format', {})
    if abs_fmt.get('type') == 'environment':
        info['abstract_env'] = 'abstract'

    # keywords_cmd
    kw_fmt = spec.get('keywords_format', {})
    if kw_fmt.get('type') == 'command':
        info['keywords_cmd'] = '\\keywords'

    # abstract_after_maketitle
    if abs_fmt.get('keywords_inside_abstract'):
        info['abstract_after_maketitle'] = True

    # metadata_commands: 从spec推导
    title_fmt = spec.get('title_format', {})
    author_fmt = spec.get('author_format', {})
    tmpl_specific = spec.get('template_specific', {})

    if title_fmt:
        if title_fmt.get('has_short_title') or title_fmt.get('title_args', 0) > 0:
            info['metadata_commands']['title'] = '\\title[短标题]{标题}'
        else:
            info['metadata_commands']['title'] = '\\title{标题}'

    if author_fmt.get('author_args') == 2:
        info['metadata_commands']['author'] = '\\Author[][EMAIL]{given_name}{surname}'
    elif author_fmt.get('author_args') == 1:
        info['metadata_commands']['author'] = '\\author{AUTHOR}'

    if tmpl_specific.get('runningtitle', {}).get('exists'):
        info['metadata_commands']['runningtitle'] = '\\runningtitle{SHORT TITLE}'
    if tmpl_specific.get('runningauthor', {}).get('exists'):
        info['metadata_commands']['runningauthor'] = '\\runningauthor{SHORT AUTHOR}'
    if tmpl_specific.get('correspondence', {}).get('exists'):
        info['metadata_commands']['correspondence'] = '\\correspondence{EMAIL}'

    return info


def derive_metadata_block(spec, skeleton_info):
    """从spec+skeleton_info生成元数据区LaTeX代码

    注意：只生成框架命令（空占位符），具体内容由assemble_tex从Word内容填充。
    不生成\title等由Word内容替换的命令。

    Returns:
        str: 元数据区LaTeX代码（\\begin{document}之后、正文之前的内容）
    """
    lines = []

    # 模板特有命令（空占位符，由assemble_tex填充）
    tmpl = spec.get('template_specific', {})

    # date declarations（空占位符）
    for date_cmd in ('received', 'revised', 'accepted', 'published'):
        if tmpl.get(date_cmd, {}).get('exists'):
            lines.append(f'\\{date_cmd}{{}}')

    lines.append('\\maketitle')
    lines.append('')

    return '\n'.join(lines)
