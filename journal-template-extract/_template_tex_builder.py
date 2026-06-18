#!/usr/bin/env python3
r"""
核心tex构建模块: 从排版规格和模板结构生成template.tex

主入口:
- build_template_tex: 根据排版规格和模板结构生成完整的template.tex

辅助函数已移至:
- _tex_emit_helpers.py: 规格注释生成、命令输出、正文段落生成
- _tex_transform.py: 模板变换、用户命令判断、完整规格写入
"""

import re
from collections import OrderedDict
from pathlib import Path
import sys
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.latex_parse_utils import BS
from _tex_emit_helpers import (
    _emit_spec_comments,
    _emit_cmd_with_spec,
    _emit_section_cmd,
    _emit_body_from_template,
)
from _tex_transform import (
    _transform_template_tex,
    _is_user_cmd,
    _write_full_spec,
)


def build_template_tex(spec, template_tex_path=None, output_dir=None,
                       journal_name='unknown'):
    """根据排版规格和模板结构生成完整的template.tex

    Args:
        spec: 完整排版规格字典（由 LayoutSpecExtract 提取）
        template_tex_path: 原始模板 .tex 文件路径（可选）
        output_dir: 输出目录
        journal_name: 期刊名称

    Returns:
        生成的 template.tex 文件路径
    """
    L = []

    source_template = Path(template_tex_path) if template_tex_path else None
    if source_template and source_template.exists():
        template_lines = source_template.read_text(encoding='utf-8').splitlines(True)
        _emit_from_existing_template(L, template_lines, spec)

        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f'{journal_name}_template.tex'
        else:
            output_path = Path(f'{journal_name}_template.tex')

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(L) + '\n')

        return str(output_path)

    # ─── 1. Documentclass ───
    doc_spec = spec.get('document', {})
    doc_class = doc_spec.get('documentclass', 'article')
    doc_options = doc_spec.get('options', [])
    if doc_options:
        opts_str = ', '.join(doc_options)
        L.append(f'\\documentclass[{opts_str}]{{{doc_class}}}')
    else:
        L.append(f'\\documentclass{{{doc_class}}}')
    L.append('')

    # ─── 2. Preamble: 必需包 ───
    _emit_required_packages(L, spec)

    # ─── 3. Preamble: 排版规格注入 ───
    _emit_spec_comments(L, spec)

    # ─── 4. Preamble: 字体设置 ───
    _emit_font_settings(L, spec)

    # ─── 5. Preamble: 页面布局 ───
    _emit_page_layout(L, spec)

    # ─── 6. Preamble: 章节标题设置 ───
    _emit_heading_settings(L, spec)

    # ─── 7. Preamble: 其他设置 ───
    _emit_misc_settings(L, spec)

    # ─── 8. 生成模板骨架结构 ───
    if template_tex_path and Path(template_tex_path).exists():
        template_lines = Path(template_tex_path).read_text(encoding='utf-8').splitlines(True)
        _emit_from_existing_template(L, template_lines, spec)
    else:
        _emit_skeleton_structure(L, spec)

    # 写入文件
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f'{journal_name}_template.tex'
    else:
        output_path = Path(f'{journal_name}_template.tex')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L) + '\n')

    return str(output_path)


def _emit_required_packages(L, spec):
    """生成必需包声明"""
    doc_spec = spec.get('document', {})

    # 基础包
    base_pkgs = [
        ('amsmath', None),
        ('graphicx', None),
        ('hyperref', None),
    ]

    # 从spec获取额外包
    pkgs_spec = spec.get('required_packages', {})
    extra_pkgs = []
    for pkg_name, pkg_opts in pkgs_spec.items():
        if pkg_name not in ('amsmath', 'graphicx', 'hyperref'):
            extra_pkgs.append((pkg_name, pkg_opts))

    # 参考文献包
    bib_spec = spec.get('bibliography', {})
    bib_style = bib_spec.get('bib_style', 'natbib')
    if bib_style == 'natbib':
        natbib_opts = bib_spec.get('natbib_options', '')
        if natbib_opts:
            base_pkgs.append(('natbib', natbib_opts))
        else:
            base_pkgs.append(('natbib', None))
    elif bib_style == 'biblatex':
        biblatex_opts = bib_spec.get('biblatex_options', '')
        if biblatex_opts:
            base_pkgs.append(('biblatex', biblatex_opts))
        else:
            base_pkgs.append(('biblatex', None))

    # TikZ（表格需要）
    base_pkgs.append(('tikz', None))

    # caption
    cap_spec = spec.get('figure', {})
    if cap_spec.get('caption_setup'):
        cap_opts = cap_spec.get('caption_package_options', '')
        if cap_opts:
            base_pkgs.append(('caption', cap_opts))
        else:
            base_pkgs.append(('caption', None))

    # 输出所有包
    all_pkgs = base_pkgs + extra_pkgs
    for pkg_name, opts in all_pkgs:
        if opts:
            L.append(f'\\usepackage[{opts}]{{{pkg_name}}}')
        else:
            L.append(f'\\usepackage{{{pkg_name}}}')
    L.append('')


def _emit_font_settings(L, spec):
    """生成字体设置"""
    fonts = spec.get('fonts', {})
    if not fonts:
        return

    needs_fontspec = False
    if fonts.get('serif_name') or fonts.get('sans_name') or fonts.get('mono_name'):
        needs_fontspec = True

    if needs_fontspec:
        L.append('\\usepackage{fontspec}')
        if fonts.get('serif_name'):
            L.append(f'\\setmainfont{{{fonts["serif_name"]}}}')
        if fonts.get('sans_name'):
            L.append(f'\\setsansfont{{{fonts["sans_name"]}}}')
        if fonts.get('mono_name'):
            L.append(f'\\setmonofont{{{fonts["mono_name"]}}}')
        L.append('')

    if fonts.get('math_name'):
        L.append('\\usepackage{unicode-math}')
        L.append(f'\\setmathfont{{{fonts["math_name"]}}}')
        L.append('')

    # CJK字体
    if fonts.get('cjk_main_font'):
        L.append('\\usepackage[UTF8]{xeCJK}')
        if fonts.get('cjk_main_font'):
            L.append(f'\\setCJKmainfont{{{fonts["cjk_main_font"]}}}')
        L.append('')


def _emit_page_layout(L, spec):
    """生成页面布局设置"""
    page = spec.get('page_layout', {})
    if not page:
        return

    L.append('\\usepackage{geometry}')
    geo_opts = []
    for k in ('paperwidth', 'paperheight', 'left', 'right', 'top', 'bottom',
               'headheight', 'headsep', 'footskip', 'columnsep'):
        if k in page:
            geo_opts.append(f'{k}={page[k]}')

    if spec.get('columns') == 2:
        geo_opts.append('twocolumn')

    if geo_opts:
        L.append(f'\\geometry{{{", ".join(geo_opts)}}}')
    L.append('')


def _emit_heading_settings(L, spec):
    """生成章节标题设置"""
    headings = spec.get('headings', {})
    if not headings:
        return

    has_titlesec = any(h.get('before_skip') or h.get('after_skip') for h in headings.values())
    if has_titlesec:
        L.append('\\usepackage{titlesec}')
        for level, h in headings.items():
            if level not in ('section', 'subsection', 'subsubsection', 'paragraph'):
                continue
            fmt_parts = []
            if h.get('weight') == 'bold':
                fmt_parts.append('\\bfseries')
            if h.get('size_name'):
                fmt_parts.append(f'\\{h["size_name"]}')
            if h.get('shape') == 'italic':
                fmt_parts.append('\\itshape')
            fmt = ''.join(fmt_parts) if fmt_parts else '\\normalfont'

            label = h.get('label_format', '')
            sep = h.get('sep', '0.5em')

            if label:
                L.append(f'\\titleformat{{{BS}{level}}}{{{fmt}}}{{{label}}}{{{sep}}}{{}}')
            else:
                L.append(f'\\titleformat{{{BS}{level}}}{{{fmt}}}{{\\thesection}}{{{sep}}}{{}}')

            before = h.get('before_skip', '')
            after = h.get('after_skip', '')
            if before and after:
                L.append(f'\\titlespacing*{{{BS}{level}}}{{0pt}}{{{before}}}{{{after}}}{{0pt}}')
        L.append('')


def _emit_misc_settings(L, spec):
    """生成其他设置"""
    # 行距
    body = spec.get('body_text', {})
    spacing = str(body.get('line_spacing', '')).strip()
    spacing_key = spacing.lower()
    if spacing and spacing_key not in ('single', 'single (1.0)', '1', '1.0'):
        L.append('\\usepackage{setspace}')
        if spacing_key in ('1.5', 'onehalf', 'onehalfspacing') or spacing_key.startswith('onehalf'):
            L.append('\\onehalfspacing')
        elif spacing_key in ('2', '2.0', 'double', 'doublespacing') or spacing_key.startswith('double'):
            L.append('\\doublespacing')
        elif re.fullmatch(r'\d+(?:\.\d+)?', spacing):
            L.append(f'\\setstretch{{{spacing}}}')
        else:
            L.append(f'% line spacing from template: {spacing}')
        L.append('')

    # 首行缩进
    indent = body.get('first_line_indent', '')
    if indent == 'none':
        L.append('\\setlength{\\parindent}{0pt}')
        L.append('')
    elif indent and indent != 'default':
        L.append(f'\\setlength{{\\parindent}}{{{indent}}}')
        L.append('')

    # figure/table float position
    fig = spec.get('figure', {})
    float_pos = fig.get('float_position', 'htbp')
    if float_pos != 'htbp':
        L.append(f'% 图表浮动位置: {float_pos}')
        L.append('')

    # placeins
    L.append('\\usepackage{placeins}')
    L.append('')

    # subcaption
    if fig.get('subcaption'):
        L.append('\\usepackage{subcaption}')
        L.append('')


def _emit_from_existing_template(L, template_lines, spec):
    """从已有模板生成（变换模式）"""
    headings = spec.get('headings', {})
    special_envs = spec.get('special_envs', {})
    doc_class_options = ', '.join(spec.get('document', {}).get('options', []))
    required_packages = set(spec.get('required_packages', {}).keys())
    document_format = spec.get('document', {})

    transformed = _transform_template_tex(
        template_lines, spec, headings, special_envs,
        doc_class_options, required_packages, document_format)

    L.extend(line.rstrip('\n') for line in transformed)


def _emit_skeleton_structure(L, spec):
    """从排版规格生成骨架结构（无原始模板时）"""
    headings = spec.get('headings', {})
    special_envs = spec.get('special_envs', {})

    # \begin{document}
    L.append('\\begin{document}')
    L.append('')

    # 标题
    _emit_cmd_with_spec(L, 'title', '\\title{TITLE}', spec.get('headings', {}).get('title', {}), '论文标题')
    _emit_cmd_with_spec(L, 'author', '\\author{AUTHOR NAME}', {}, '作者')

    # \maketitle
    L.append('\\maketitle')
    L.append('')

    # abstract
    abs_spec = spec.get('abstract', {})
    abs_notes = []
    if abs_spec.get('label_weight') == 'bold':
        abs_notes.append('bold')
    if abs_spec.get('label_size'):
        abs_notes.append(f'{abs_spec["label_size"]}pt')
    note = f' ({", ".join(abs_notes)})' if abs_notes else ''
    L.append(f'% -- abstract{note} --')
    L.append('\\begin{abstract}')
    L.append('ABSTRACT CONTENT')
    L.append('\\end{abstract}')
    L.append('')

    # keywords
    kw_spec = spec.get('keywords', {})
    kw_notes = []
    if kw_spec.get('prefix_weight') == 'bold':
        kw_notes.append('bold')
    note = f' ({", ".join(kw_notes)})' if kw_notes else ''
    L.append(f'% -- keywords{note} --')
    L.append('\\keywords{KEYWORDS}')
    L.append('')

    # 各章节
    sections = _infer_sections(spec)
    _emit_body_from_template(L, sections, headings, special_envs)

    # 参考文献
    bib_spec = spec.get('bibliography', {})
    bst = bib_spec.get('bst_file', '')
    bib_file = bib_spec.get('bib_filename', 'references')
    if bst:
        L.append(f'\\bibliographystyle{{{bst}}}')
    L.append(f'\\bibliography{{{bib_file}}}')
    L.append('')

    # \end{document}
    L.append('\\end{document}')


def _infer_sections(spec):
    """从排版规格推断章节结构"""
    sections = []
    doc_spec = spec.get('document', {})

    # 基于模板类型推断
    template_type = doc_spec.get('template_type', '')
    if template_type == 'copernicus':
        sections = ['introduction', 'section', 'section', 'section',
                     'conclusions', 'acknowledgements', 'thebibliography']
    elif template_type == 'elsevier':
        sections = ['section', 'subsection', 'section', 'subsection',
                     'section', 'acknowledgements', 'thebibliography']
    elif template_type == 'springer':
        sections = ['section', 'section', 'section',
                     'acknowledgements', 'thebibliography']
    else:
        sections = ['section', 'subsection', 'section', 'subsection',
                     'section', 'acknowledgements', 'thebibliography']

    return sections
