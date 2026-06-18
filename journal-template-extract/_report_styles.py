#!/usr/bin/env python3
r"""
报告生成 + Word样式映射模块
包含函数:
  - generate_report(spec)
  - spec_to_word_styles(spec)
"""
from collections import OrderedDict


def generate_report(spec):
    """将排版规格转为可读的Markdown报告"""
    lines = []
    lines.append(f'# 排版规格报告: {spec["source"]}')
    lines.append(f'\n> 基准字号: {spec["base_size_pt"]}pt')
    lines.append('')

    # 1. 页面布局
    if spec.get('page_layout'):
        lines.append('## 1. 页面布局')
        pl = spec['page_layout']
        if 'paperwidth' in pl or 'paperheight' in pl:
            lines.append(f'- 纸张: {pl.get("paperwidth", "?")} × {pl.get("paperheight", "?")}')
        for key in ('textheight', 'textwidth', 'oddsidemargin', 'evensidemargin',
                     'topmargin', 'headheight', 'headsep', 'footskip',
                     'marginparwidth', 'marginparsep', 'columnsep'):
            if key in pl:
                lines.append(f'- {key}: {pl[key]}')
        if 'geometry' in pl:
            lines.append(f'- geometry: `{pl["geometry"]}`')
        lines.append('')

    # 2. 栏数
    lines.append('## 2. 栏数')
    lines.append(f'- {spec.get("columns", "未指定")}')
    lines.append('')

    # 3. 字体族
    if spec.get('fonts'):
        lines.append('## 3. 字体族')
        for k, v in spec['fonts'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 4. 论文标题
    lines.append('## 4. 论文标题')
    t = spec.get('title', {})
    lines.append(f'- 字号: {t.get("size_name", "?")} ({t.get("size_pt", "?")}pt)')
    lines.append(f'- 字重: {t.get("weight", "?")}')
    lines.append(f'- 字形: {t.get("shape", "?")}')
    if t.get('font_family'):
        lines.append(f'- 字体族: {t["font_family"]}')
    if t.get('alignment'):
        lines.append(f'- 对齐: {t["alignment"]}')
    if t.get('color'):
        lines.append(f'- 颜色: {t["color"]}')
    lines.append('')

    # 5. 作者
    lines.append('## 5. 作者')
    a = spec.get('author', {})
    lines.append(f'- 字号: {a.get("size_name", "未指定")} ({a.get("size_pt", "?")}pt)')
    lines.append(f'- 字重: {a.get("weight", "?")}')
    if a.get('font_declaration'):
        lines.append(f'- 字体声明: `{a["font_declaration"]}`')
    if a.get('affilfont'):
        lines.append(f'- 单位字体: `{a["affilfont"]}`')
    lines.append('')

    # 6. 摘要
    lines.append('## 6. 摘要')
    ab = spec.get('abstract', {})
    lines.append(f'- 标签: {ab.get("label", "Abstract")}')
    lines.append(f'- 标签字重: {ab.get("label_weight", "?")}')
    lines.append(f'- 正文字号: {ab.get("size_name", "未指定")} ({ab.get("size_pt", "?")}pt)')
    if ab.get('indent'):
        lines.append('- 缩进: 是')
    if ab.get('width'):
        lines.append(f'- 宽度: {ab["width"]}')
    lines.append('')

    # 7. 章节标题
    lines.append('## 7. 章节标题')
    for name, h in spec.get('headings', {}).items():
        parts = [f'字号: {h.get("size_name", "?")}({h.get("size_pt", "?")}pt)']
        parts.append(f'字重: {h.get("weight", "?")}')
        parts.append(f'字形: {h.get("shape", "?")}')
        if h.get('font_family'):
            parts.append(f'字体族: {h["font_family"]}')
        if h.get('before_skip'):
            parts.append(f'前间距: {h["before_skip"]}')
        if h.get('after_skip'):
            parts.append(f'后间距: {h["after_skip"]}')
        if h.get('alignment'):
            parts.append(f'对齐: {h["alignment"]}')
        if h.get('color'):
            parts.append(f'颜色: {h["color"]}')
        lines.append(f'- **{name}**: {", ".join(parts)}')
    lines.append('')

    # 8. 正文
    lines.append('## 8. 正文')
    bt = spec.get('body_text', {})
    lines.append(f'- 基准字号: {bt.get("font_size_pt", "?")}pt')
    lines.append(f'- 字体: {bt.get("font_family", "?")}')
    lines.append(f'- 行距: {bt.get("line_spacing", "?")}')
    lines.append(f'- 首行缩进: {bt.get("first_line_indent", "?")}')
    lines.append(f'- 段间距: {bt.get("paragraph_skip", "?")}')
    lines.append('')

    # 9. Caption
    lines.append('## 9. Caption (图/表标题)')
    cap = spec.get('caption', {})
    lines.append(f'- 字号: {cap.get("font_size", "?")} ({cap.get("font_size_pt", "?")}pt)')
    lines.append(f'- 字重: {cap.get("weight", "?")}')
    lines.append(f'- 标签字重: {cap.get("label_weight", "?")}')
    lines.append(f'- 分隔符: {cap.get("separator", "?")}')
    lines.append(f'- 图标题位置: {cap.get("figure_position", "?")}')
    lines.append(f'- 表标题位置: {cap.get("table_position", "?")}')
    lines.append('')

    # 10. 表格
    lines.append('## 10. 表格')
    tbl = spec.get('table', {})
    lines.append(f'- 表头字重: {tbl.get("header_weight", "?")}')
    lines.append(f'- 表头字号: {tbl.get("header_size", "?")} ({tbl.get("header_size_pt", "?")}pt)')
    lines.append(f'- 表体字号: {tbl.get("body_size", "?")} ({tbl.get("body_size_pt", "?")}pt)')
    lines.append(f'- 标题位置: {tbl.get("caption_position", "?")}')
    lines.append(f'- 线型: {tbl.get("rule_style", "?")}')
    lines.append('')

    # 11. 脚注
    lines.append('## 11. 脚注')
    fn = spec.get('footnote', {})
    lines.append(f'- 字号: {fn.get("font_size", "?")} ({fn.get("font_size_pt", "?")}pt)')
    if fn.get('mark_style'):
        lines.append(f'- 标记样式: {fn["mark_style"]}')
    if fn.get('numbering_format'):
        lines.append(f'- 编号格式: `{fn["numbering_format"]}`')
    lines.append('')

    # 12. 参考文献
    lines.append('## 12. 参考文献')
    bib = spec.get('bibliography', {})
    lines.append(f'- 字号: {bib.get("font_size", "?")} ({bib.get("font_size_pt", "?")}pt)')
    lines.append(f'- 引用风格: {bib.get("style", "?")}')
    if bib.get('bst_file'):
        lines.append(f'- bst文件: {bib["bst_file"]}')
    if bib.get('label_format'):
        lines.append(f'- 标签格式: `{bib["label_format"]}`')
    lines.append('')

    # 13. 编号
    if spec.get('numbering'):
        lines.append('## 13. 编号格式')
        for k, v in spec['numbering'].items():
            lines.append(f'- {k}: `{v}`')
        lines.append('')

    # 14. 间距
    if spec.get('spacing'):
        lines.append('## 14. 间距')
        for k, v in spec['spacing'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 15. 页眉页脚
    if spec.get('header_footer'):
        lines.append('## 15. 页眉页脚')
        hf = spec['header_footer']
        if 'pagestyle' in hf:
            lines.append(f'- 页面样式: {hf["pagestyle"]}')
        if 'header' in hf:
            for pos, content in hf['header'].items():
                lines.append(f'- 页眉({pos}): `{content}`')
        if 'footer' in hf:
            for pos, content in hf['footer'].items():
                lines.append(f'- 页脚({pos}): `{content}`')
        if 'page_number_format' in hf:
            lines.append(f'- 页码格式: `{hf["page_number_format"]}`')
        lines.append('')

    # 16. 颜色
    if spec.get('colors'):
        lines.append('## 16. 颜色')
        for name, val in spec['colors'].items():
            if isinstance(val, dict):
                lines.append(f'- {name}: {val["model"]}({val["value"]})')
            else:
                lines.append(f'- {name}: {val}')
        lines.append('')

    # 17. 列表
    if spec.get('lists'):
        lines.append('## 17. 列表')
        for k, v in spec['lists'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 18. 特殊环境
    if spec.get('special_environments'):
        lines.append('## 18. 特殊环境')
        for name, info in spec['special_environments'].items():
            parts = []
            for k, v in info.items():
                parts.append(f'{k}={v}')
            lines.append(f'- **{name}**: {", ".join(str(p) for p in parts)}')
        lines.append('')

    # 19. 自定义命令
    if spec.get('custom_commands'):
        lines.append('## 19. 自定义命令 (前20个)')
        count = 0
        for name, info in spec.get('custom_commands', {}).items():
            if count >= 20: break
            if 'nargs' in info:
                lines.append(f'- \\{name} (参数: {info["nargs"]}): `{info.get("definition_preview", "")}`')
            else:
                lines.append(f'- \\{name}: `{info.get("definition_preview", info.get("defined_in_template", ""))}`')
            count += 1
        lines.append('')

    # 20. 依赖宏包
    if spec.get('packages'):
        lines.append('## 20. 依赖宏包')
        lines.append(', '.join(spec['packages']))
        lines.append('')

    # ═══ v3.1 新增: 14个遗漏类别 ═══

    # 21. 浮动体设置
    if spec.get('float_settings'):
        lines.append('## 21. 浮动体设置')
        for k, v in spec['float_settings'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 22. 数学间距
    if spec.get('math_display_skip'):
        lines.append('## 22. 数学间距')
        for k, v in spec['math_display_skip'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 23. 脚注详细格式
    if spec.get('footnote_detail'):
        lines.append('## 23. 脚注详细格式')
        for k, v in spec['footnote_detail'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 24. hyperref配置
    if spec.get('hyperref'):
        lines.append('## 24. hyperref配置')
        for k, v in spec['hyperref'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 25. 全局排版设置
    if spec.get('global_typography'):
        lines.append('## 25. 全局排版设置')
        for k, v in spec['global_typography'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 26. 标题页布局
    if spec.get('title_page_layout'):
        lines.append('## 26. 标题页布局')
        tpl = spec['title_page_layout']
        if 'elements' in tpl:
            for i, el in enumerate(tpl['elements'], 1):
                desc = el.get('description', el.get('name', '?'))
                size = el.get('size', '?')
                weight = el.get('weight', '')
                weight_str = f', {weight}' if weight else ''
                lines.append(f'- {i}. {desc}: {size}{weight_str}')
        if 'vspace_values' in tpl:
            lines.append(f'- 间距: {", ".join(tpl["vspace_values"])}')
        lines.append('')

    # 27. 日期声明格式
    if spec.get('date_declarations'):
        lines.append('## 27. 日期声明格式')
        for k, v in spec['date_declarations'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 28. 摘要详细格式
    if spec.get('abstract_detail'):
        lines.append('## 28. 摘要详细格式')
        for k, v in spec['abstract_detail'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 29. 关键词格式
    if spec.get('keywords'):
        lines.append('## 29. 关键词格式')
        for k, v in spec['keywords'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 30. 作者详细格式
    if spec.get('author_detail'):
        lines.append('## 30. 作者详细格式')
        for k, v in spec['author_detail'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 31. 自定义排版命令
    if spec.get('custom_typo_commands'):
        lines.append('## 31. 自定义排版命令')
        for k, v in spec['custom_typo_commands'].items():
            lines.append(f'- \\{k}: `{v}`')
        lines.append('')

    # 32. 子图subfig设置
    if spec.get('subfig_settings'):
        lines.append('## 32. 子图subfig设置')
        for k, v in spec['subfig_settings'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 33. 名称重定义
    if spec.get('name_redefinitions'):
        lines.append('## 33. 名称重定义')
        for k, v in spec['name_redefinitions'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 34. 页面样式详细
    if spec.get('page_style_detail'):
        lines.append('## 34. 页面样式详细')
        for k, v in spec['page_style_detail'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # ═══ v3.2 新增: 普适性增强 ═══

    # 35. geometry包配置
    if spec.get('geometry_config'):
        lines.append('## 35. geometry包配置')
        for k, v in spec['geometry_config'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 36. fontspec/XeLaTeX字体配置
    if spec.get('fontspec_config'):
        lines.append('## 36. fontspec/XeLaTeX字体配置')
        for k, v in spec['fontspec_config'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 37. titlesec包章节配置
    if spec.get('titlesec_config'):
        lines.append('## 37. titlesec包章节配置')
        for level, info in spec['titlesec_config'].items():
            if isinstance(info, dict):
                parts = [f'{k}={v}' for k, v in info.items()]
                lines.append(f'- **{level}**: {", ".join(parts)}')
            else:
                lines.append(f'- {level}: {info}')
        lines.append('')

    # 38. caption包配置
    if spec.get('caption_package_config'):
        lines.append('## 38. caption包配置')
        for k, v in spec['caption_package_config'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 39. biblatex配置
    if spec.get('biblatex_config'):
        lines.append('## 39. biblatex配置')
        for k, v in spec['biblatex_config'].items():
            if isinstance(v, list):
                lines.append(f'- {k}: {", ".join(str(i) for i in v)}')
            else:
                lines.append(f'- {k}: {v}')
        lines.append('')

    # 40. 定理环境
    if spec.get('theorem_envs'):
        lines.append('## 40. 定理环境')
        for name, info in spec['theorem_envs'].items():
            if isinstance(info, dict):
                parts = [f'{k}={v}' for k, v in info.items()]
                lines.append(f'- **{name}**: {", ".join(parts)}')
            else:
                lines.append(f'- {name}: {info}')
        lines.append('')

    # 41. 算法环境
    if spec.get('algorithm_envs'):
        lines.append('## 41. 算法环境')
        for k, v in spec['algorithm_envs'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 42. 行距设置
    if spec.get('setspace_config'):
        lines.append('## 42. 行距设置')
        for k, v in spec['setspace_config'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 43. enumerate样式
    if spec.get('enumerate_styles'):
        lines.append('## 43. enumerate样式')
        for k, v in spec['enumerate_styles'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 44. 表格详细格式
    if spec.get('table_detail'):
        lines.append('## 44. 表格详细格式')
        for k, v in spec['table_detail'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 45. 边注设置
    if spec.get('marginpar'):
        lines.append('## 45. 边注设置')
        for k, v in spec['marginpar'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 46. DOI/URL格式
    if spec.get('doi_url_format'):
        lines.append('## 46. DOI/URL格式')
        for k, v in spec['doi_url_format'].items():
            lines.append(f'- {k}: {v}')
        lines.append('')

    # 47. cleveref交叉引用
    if spec.get('cleveref_config'):
        lines.append('## 47. cleveref交叉引用')
        for k, v in spec['cleveref_config'].items():
            if isinstance(v, dict):
                parts = [f'{kk}={vv}' for kk, vv in v.items()]
                lines.append(f'- {k}: {", ".join(parts)}')
            else:
                lines.append(f'- {k}: {v}')
        lines.append('')

    # 48. 文档类模式/选项
    if spec.get('mode_options'):
        lines.append('## 48. 文档类模式/选项')
        mo = spec['mode_options']
        if 'declared_options' in mo:
            lines.append(f'- 声明选项: {", ".join(mo["declared_options"][:20])}')
        if 'default_options' in mo:
            lines.append(f'- 默认选项: {", ".join(mo["default_options"])}')
        if 'journal_options' in mo:
            lines.append(f'- 期刊选项: {", ".join(mo["journal_options"])}')
        if 'base_class' in mo:
            lines.append(f'- 基类: {mo["base_class"]}')
        for k, v in mo.items():
            if k not in ('declared_options', 'default_options', 'journal_options', 'base_class'):
                lines.append(f'- {k}: {v}')
        lines.append('')

    return '\n'.join(lines)


def spec_to_word_styles(spec):
    """将排版规格转为Word可用的样式参数"""
    styles = OrderedDict()
    fonts = spec.get('fonts', {})
    serif_name = fonts.get('serif_name', 'Times New Roman')

    def _base(font_en=None, size=10, bold=False, italic=False):
        return OrderedDict([
            ('font_cn', '宋体'), ('font_en', font_en or serif_name),
            ('size_pt', size), ('bold', bold), ('italic', italic),
        ])

    t = spec.get('title', {})
    styles['title'] = _base(size=t.get('size_pt') or 12, bold=t.get('weight') == 'bold', italic=t.get('shape') == 'italic')

    a = spec.get('author', {})
    styles['author'] = _base(size=a.get('size_pt') or 10, bold=a.get('weight') == 'bold')

    h = spec.get('headings', {})
    sec = h.get('section', {})
    styles['heading1'] = _base(size=sec.get('size_pt') or 10, bold=sec.get('weight') == 'bold', italic=sec.get('shape') == 'italic')
    subsec = h.get('subsection', {})
    styles['heading2'] = _base(size=subsec.get('size_pt') or 10, bold=subsec.get('weight') == 'bold', italic=subsec.get('shape') == 'italic')
    subsubsec = h.get('subsubsection', {})
    styles['heading3'] = _base(size=subsubsec.get('size_pt') or 10, bold=subsubsec.get('weight') == 'bold', italic=subsubsec.get('shape') == 'italic')

    styles['body'] = _base(size=spec.get('base_size_pt', 10))

    cap = spec.get('caption', {})
    styles['caption'] = OrderedDict([
        ('font_cn', '宋体'), ('font_en', serif_name),
        ('size_pt', cap.get('font_size_pt') or 9), ('bold', False),
        ('separator', cap.get('separator', '.')),
        ('figure_position', cap.get('figure_position', 'below')),
        ('table_position', cap.get('table_position', 'above')),
    ])

    tbl = spec.get('table', {})
    styles['table_header'] = _base(size=tbl.get('body_size_pt') or 9, bold=tbl.get('header_weight') == 'bold')
    styles['table_body'] = _base(size=tbl.get('body_size_pt') or 9)

    bib = spec.get('bibliography', {})
    styles['bibliography'] = _base(size=bib.get('font_size_pt') or 9)

    fn = spec.get('footnote', {})
    styles['footnote'] = _base(size=fn.get('font_size_pt') or 8)

    # v3.2: fontspec字体名 → Word样式覆盖
    fs = spec.get('fontspec_config', {})
    if fs.get('main_font'):
        for key in ('title', 'heading1', 'heading2', 'heading3', 'body', 'caption', 'table_header', 'table_body', 'bibliography', 'footnote', 'author'):
            if key in styles:
                styles[key]['font_en'] = fs['main_font']
    if fs.get('sans_font'):
        styles['heading_sans'] = _base(font_en=fs['sans_font'], size=10, bold=True)

    # v3.2: 行距
    ss = spec.get('setspace_config', {})
    if ss.get('stretch'):
        styles['line_spacing'] = ss['stretch']
    elif ss.get('linespread'):
        styles['line_spacing'] = ss['linespread']
    elif ss.get('baselinestretch'):
        styles['line_spacing'] = ss['baselinestretch']

    # v3.2: 算法环境字号
    styles['algorithm'] = _base(size=9)

    # v3.2: 定理环境字号/字重
    styles['theorem'] = _base(size=spec.get('base_size_pt', 10), italic=True)

    # v3.2: 边注字号
    styles['marginpar'] = _base(size=8)

    return styles
