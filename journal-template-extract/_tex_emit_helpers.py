#!/usr/bin/env python3
r"""
LaTeX文件生成辅助函数

包含:
- _emit_spec_comments: 将排版规格以注释形式注入 .tex 文件
- _emit_cmd_with_spec: 生成命令并附加排版规格注释
- _emit_section_cmd: 生成章节命令并附加排版规格注释
- _emit_body_from_template: 从 template.tex 结构生成中间正文段落
"""


def _emit_spec_comments(L, spec):
    """将排版规格以注释形式注入 .tex 文件"""
    if not spec:
        return

    # 页面布局
    page_layout = spec.get('page_layout', {})
    if page_layout:
        pw = page_layout.get('paperwidth', '')
        ph = page_layout.get('paperheight', '')
        if pw and ph:
            L.append(f'% 纸张: {pw} × {ph}')
        for k in ('textheight', 'textwidth', 'oddsidemargin', 'topmargin',
                   'columnsep'):
            if k in page_layout:
                L.append(f'% {k}: {page_layout[k]}')

    # 栏数
    cols = spec.get('columns', '')
    if cols:
        L.append(f'% 栏数: {cols}')

    # 字体
    fonts = spec.get('fonts', {})
    if fonts:
        for k in ('serif_name', 'sans_name', 'mono_name', 'math_name'):
            if k in fonts:
                L.append(f'% 字体-{k.replace("_name","")}: {fonts[k]}')

    # 基准字号
    base = spec.get('base_size_pt', 10)
    L.append(f'% 基准字号: {base}pt')

    # 行距
    body_spec = spec.get('body_text', {})
    if body_spec.get('line_spacing'):
        L.append(f'% 行距: {body_spec["line_spacing"]}')
    if body_spec.get('first_line_indent'):
        L.append(f'% 首行缩进: {body_spec["first_line_indent"]}')

    # 章节标题格式
    headings = spec.get('headings', {})
    for level, h in headings.items():
        parts = [f'{h.get("size_name","?")}({h.get("size_pt","?")}pt)']
        parts.append(h.get('weight', 'normal'))
        if h.get('shape') and h['shape'] != 'normal':
            parts.append(h['shape'])
        if h.get('before_skip'):
            parts.append(f'前{h["before_skip"]}')
        if h.get('after_skip'):
            parts.append(f'后{h["after_skip"]}')
        L.append(f'% {level}: {", ".join(parts)}')


def _emit_cmd_with_spec(L, cmd, latex_code, spec_dict, desc):
    """生成命令并附加排版规格注释"""
    notes = []
    if spec_dict.get('size_name'):
        notes.append(f'{spec_dict["size_name"]}({spec_dict.get("size_pt","?")}pt)')
    if spec_dict.get('weight') and spec_dict['weight'] != 'normal':
        notes.append(spec_dict['weight'])
    if spec_dict.get('shape') and spec_dict['shape'] != 'normal':
        notes.append(spec_dict['shape'])
    if spec_dict.get('alignment'):
        notes.append(spec_dict['alignment'])
    if spec_dict.get('font_family'):
        notes.append(spec_dict['font_family'])
    note_str = f' ({", ".join(notes)})' if notes else ''
    L.append(f'% -- \\{cmd} - {desc}{note_str} --')
    L.append(latex_code)
    L.append('')


def _emit_section_cmd(L, cmd, title, headings, special_envs, placeholder):
    """生成章节命令并附加排版规格注释"""
    # 查找对应的排版规格
    level_name = cmd
    if cmd in ('introduction', 'conclusions', 'methods', 'results', 'discussion'):
        env_info = special_envs.get(cmd, {})
        if isinstance(env_info, dict) and env_info.get('maps_to'):
            level_name = env_info['maps_to']
        else:
            level_name = 'section'

    h = headings.get(level_name, {})
    notes = []
    if h.get('size_name'):
        notes.append(f'{h["size_name"]}({h.get("size_pt","?")}pt)')
    if h.get('weight') and h['weight'] != 'normal':
        notes.append(h['weight'])
    if h.get('shape') and h['shape'] != 'normal':
        notes.append(h['shape'])
    if h.get('before_skip'):
        notes.append(f'前间距: {h["before_skip"]}')
    if h.get('after_skip'):
        notes.append(f'后间距: {h["after_skip"]}')
    if h.get('alignment'):
        notes.append(h['alignment'])
    note_str = f' ({", ".join(notes)})' if notes else ''

    if cmd in ('introduction', 'conclusions') and cmd != level_name:
        # Copernicus 自定义命令
        L.append(f'% -- \\{cmd} -> \\{level_name}{note_str} --')
        L.append(f'\\{cmd}')
        L.append(placeholder)
    elif cmd in ('section', 'subsection', 'subsubsection'):
        L.append(f'% -- \\{cmd} {{}} {note_str} --')
        L.append(f'\\{cmd}{{{title}}}')
        L.append(placeholder)
    else:
        L.append(f'% -- \\{cmd}{note_str} --')
        L.append(f'\\{cmd}')
        L.append(placeholder)
    L.append('')


def _emit_body_from_template(L, sections, headings, special_envs):
    """从 template.tex 结构生成中间正文段落"""
    seen = set()
    skip_cmds = {
        'title', 'Author', 'author', 'affil', 'affiliation',
        'runningtitle', 'runningauthor', 'runninghead', 'correspondence',
        'received', 'pubdiscuss', 'revised', 'accepted', 'published',
        'firstpage', 'maketitle', 'introduction', 'conclusions',
        'codeavailability', 'dataavailability', 'codedataavailability',
        'sampleavailability', 'videosupplement',
        'authorcontribution', 'competinginterests', 'disclaimer',
        'copyrightstatement', 'appendix', 'appendixfigures', 'appendixtables',
        'keywords', 'bibliographystyle', 'bibliography', 'label', 'caption',
    }
    skip_envs = {'abstract', 'acknowledgements', 'acknowledgment', 'thebibliography'}

    for sec in sections:
        if sec.startswith('begin_'):
            env = sec[6:]
            if env in skip_envs or env in seen:
                continue
            seen.add(env)
            placeholder = {
                'figure': '\\centering\\includegraphics[width=\\columnwidth]{figure_file}\n\\caption{FIGURE CAPTION}',
                'table': '\\caption{TABLE CAPTION}\n\\centering\n\\begin{tabular}{cc}\nA & B \\\\\nC & D\n\\end{tabular}',
                'equation': 'EQUATION CONTENT',
                'align': 'a &= b \\\\ c &= d',
                'itemizewithoutindent': '\\item ITEM',
                'plainlist': '\\item ITEM',
                'listing': '\\item ITEM',
                'reaction': 'REACTION CONTENT',
                'algorithm': '\\caption{ALGORITHM CAPTION}\n\\begin{algorithmic}\n\\STATE $algorithm steps$\n\\end{algorithmic}',
            }.get(env, 'CONTENT')
            L.append(f'\\begin{{{env}}}')
            if placeholder:
                L.append(placeholder)
            L.append(f'\\end{{{env}}}')
            L.append('')
        elif sec.startswith('end_'):
            pass  # begin 已处理
        else:
            if sec in skip_cmds or sec in seen:
                continue
            seen.add(sec)
            if sec == 'section':
                _emit_section_cmd(L, 'section', 'SECTION TITLE', headings, {}, 'SECTION TEXT')
            elif sec == 'subsection':
                _emit_section_cmd(L, 'subsection', 'Subsection Title', headings, {}, 'Subsection text.')
            elif sec == 'subsubsection':
                _emit_section_cmd(L, 'subsubsection', 'Subsubsection Title', headings, {}, 'Subsubsection text.')
            elif sec == 'paragraph':
                _emit_section_cmd(L, 'paragraph', 'Paragraph Title', headings, {}, 'Paragraph text.')
            else:
                # 其他命令
                L.append(f'\\{sec}{{}}')
                L.append('')
