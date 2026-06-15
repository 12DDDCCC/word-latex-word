#!/usr/bin/env python3
r"""
LaTeX模板完整排版规格提取器
从.cls/.sty/.cfg文件中提取所有排版要求:
  - 页面布局(纸张大小、边距、栏数)
  - 字体族(衬线/无衬线/等宽/数学)
  - 各元素字号/字重/字形/对齐
  - 标题/作者/摘要/章节/正文/脚注/表格/图片/参考文献
  - 编号格式(章节/图/表/公式/脚注)
  - 间距(行距/段距/首行缩进/标题前后间距)
  - 页眉页脚内容与格式
  - 颜色定义
  - 列表样式
  - 特殊环境(acknowledgments/appendix/dataavailability等)
  - caption格式(字号/位置/分隔符)
  - 表格样式(线型/字号/表头字重)

输出: 结构化JSON + 可读Markdown报告 + Word样式映射

用法:
  python layout_spec_extract.py <cls文件> [sty文件...] [-o 输出目录]
  python layout_spec_extract.py copernicus.cls -o output/
"""
import os, sys, re, json
from collections import OrderedDict
try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, ValueError):
    pass

# 导入共享工具
from shared.latex_parse_utils import _cmd, _read

# 导入子模块方法
from _extract_page_layout import (
    _extract_page_layout, _eval_dimexpr, _extract_columns, _extract_fonts
)
from _extract_element_spec import (
    _extract_title, _extract_author, _extract_abstract,
    _find_balanced_braces, _extract_headings,
    _extract_body_text, _extract_caption,
    _extract_table_spec, _extract_figure_spec,
    _extract_footnote, _extract_bibliography
)
from _extract_detail_spec import (
    _extract_numbering, _extract_spacing, _extract_header_footer,
    _extract_colors, _extract_lists, _extract_special_envs,
    _extract_custom_commands, _extract_packages, _extract_float_settings
)
from _extract_advanced_spec import (
    _extract_math_display_skip, _extract_footnote_detail, _extract_hyperref,
    _extract_global_typography, _extract_title_page_layout, _extract_date_declarations,
    _extract_abstract_detail, _extract_keywords, _extract_author_detail,
    _extract_custom_typo_commands, _extract_subfig_settings, _extract_name_redefinitions,
    _extract_page_style_detail, _extract_geometry_config, _extract_fontspec_config,
    _extract_titlesec_config, _extract_caption_package_config, _extract_biblatex_config,
    _extract_theorem_envs, _extract_algorithm_envs, _extract_setspace_config,
    _extract_enumerate_styles, _extract_table_detail, _extract_marginpar,
    _extract_doi_url_format, _extract_cleveref_config, _extract_mode_options
)
from _report_styles import generate_report, spec_to_word_styles


# ═══════════════════════════════════════════════════════════════
# 提取器类
# ═══════════════════════════════════════════════════════════════

class LayoutSpecExtractor:
    def __init__(self, cls_path, extra_paths=None):
        self.cls_path = cls_path
        self.content = _read(cls_path)
        if extra_paths:
            for p in extra_paths:
                if os.path.exists(p):
                    self.content += '\n' + _read(p)
        self.base_size = self._detect_base_size()
        self.spec = OrderedDict()

    def _detect_base_size(self):
        """检测基准字号"""
        m = re.search(_cmd('ExecuteOptions') + r'\{(\d+)pt\}', self.content)
        if m: return int(m.group(1))
        m = re.search(_cmd('documentclass') + r'.*?(\d+)pt', self.content)
        if m: return int(m.group(1))
        m = re.search(_cmd('LoadClass') + r'.*?(\d+)pt', self.content)
        if m: return int(m.group(1))
        return 10

    def extract_all(self):
        """提取所有排版规格"""
        self.spec['source'] = os.path.basename(self.cls_path)
        self.spec['base_size_pt'] = self.base_size
        self.spec['page_layout'] = self._extract_page_layout()
        self.spec['columns'] = self._extract_columns()
        self.spec['fonts'] = self._extract_fonts()
        self.spec['title'] = self._extract_title()
        self.spec['author'] = self._extract_author()
        self.spec['abstract'] = self._extract_abstract()
        self.spec['headings'] = self._extract_headings()
        self.spec['body_text'] = self._extract_body_text()
        self.spec['caption'] = self._extract_caption()
        self.spec['table'] = self._extract_table_spec()
        self.spec['figure'] = self._extract_figure_spec()
        self.spec['footnote'] = self._extract_footnote()
        self.spec['bibliography'] = self._extract_bibliography()
        self.spec['numbering'] = self._extract_numbering()
        self.spec['spacing'] = self._extract_spacing()
        self.spec['header_footer'] = self._extract_header_footer()
        self.spec['colors'] = self._extract_colors()
        self.spec['lists'] = self._extract_lists()
        self.spec['special_environments'] = self._extract_special_envs()
        self.spec['custom_commands'] = self._extract_custom_commands()
        self.spec['packages'] = self._extract_packages()
        # ─── v3.1 新增: 14个遗漏的提取类别 ───
        self.spec['float_settings'] = self._extract_float_settings()
        self.spec['math_display_skip'] = self._extract_math_display_skip()
        self.spec['footnote_detail'] = self._extract_footnote_detail()
        self.spec['hyperref'] = self._extract_hyperref()
        self.spec['global_typography'] = self._extract_global_typography()
        self.spec['title_page_layout'] = self._extract_title_page_layout()
        self.spec['date_declarations'] = self._extract_date_declarations()
        self.spec['abstract_detail'] = self._extract_abstract_detail()
        self.spec['keywords'] = self._extract_keywords()
        self.spec['author_detail'] = self._extract_author_detail()
        self.spec['custom_typo_commands'] = self._extract_custom_typo_commands()
        self.spec['subfig_settings'] = self._extract_subfig_settings()
        self.spec['name_redefinitions'] = self._extract_name_redefinitions()
        self.spec['page_style_detail'] = self._extract_page_style_detail()
        # ─── v3.2 新增: 普适性增强 ───
        self.spec['geometry_config'] = self._extract_geometry_config()
        self.spec['fontspec_config'] = self._extract_fontspec_config()
        self.spec['titlesec_config'] = self._extract_titlesec_config()
        self.spec['caption_package_config'] = self._extract_caption_package_config()
        self.spec['biblatex_config'] = self._extract_biblatex_config()
        self.spec['theorem_envs'] = self._extract_theorem_envs()
        self.spec['algorithm_envs'] = self._extract_algorithm_envs()
        self.spec['setspace_config'] = self._extract_setspace_config()
        self.spec['enumerate_styles'] = self._extract_enumerate_styles()
        self.spec['table_detail'] = self._extract_table_detail()
        self.spec['marginpar'] = self._extract_marginpar()
        self.spec['doi_url_format'] = self._extract_doi_url_format()
        self.spec['cleveref_config'] = self._extract_cleveref_config()
        self.spec['mode_options'] = self._extract_mode_options()
        return self.spec


# ═══════════════════════════════════════════════════════════════
# 注入子模块方法到类
# ═══════════════════════════════════════════════════════════════

# _extract_page_layout.py
LayoutSpecExtractor._extract_page_layout = _extract_page_layout
LayoutSpecExtractor._eval_dimexpr = _eval_dimexpr
LayoutSpecExtractor._extract_columns = _extract_columns
LayoutSpecExtractor._extract_fonts = _extract_fonts

# _extract_element_spec.py
LayoutSpecExtractor._extract_title = _extract_title
LayoutSpecExtractor._extract_author = _extract_author
LayoutSpecExtractor._extract_abstract = _extract_abstract
LayoutSpecExtractor._find_balanced_braces = staticmethod(_find_balanced_braces)
LayoutSpecExtractor._extract_headings = _extract_headings
LayoutSpecExtractor._extract_body_text = _extract_body_text
LayoutSpecExtractor._extract_caption = _extract_caption
LayoutSpecExtractor._extract_table_spec = _extract_table_spec
LayoutSpecExtractor._extract_figure_spec = _extract_figure_spec
LayoutSpecExtractor._extract_footnote = _extract_footnote
LayoutSpecExtractor._extract_bibliography = _extract_bibliography

# _extract_detail_spec.py
LayoutSpecExtractor._extract_numbering = _extract_numbering
LayoutSpecExtractor._extract_spacing = _extract_spacing
LayoutSpecExtractor._extract_header_footer = _extract_header_footer
LayoutSpecExtractor._extract_colors = _extract_colors
LayoutSpecExtractor._extract_lists = _extract_lists
LayoutSpecExtractor._extract_special_envs = _extract_special_envs
LayoutSpecExtractor._extract_custom_commands = _extract_custom_commands
LayoutSpecExtractor._extract_packages = _extract_packages
LayoutSpecExtractor._extract_float_settings = _extract_float_settings

# _extract_advanced_spec.py
LayoutSpecExtractor._extract_math_display_skip = _extract_math_display_skip
LayoutSpecExtractor._extract_footnote_detail = _extract_footnote_detail
LayoutSpecExtractor._extract_hyperref = _extract_hyperref
LayoutSpecExtractor._extract_global_typography = _extract_global_typography
LayoutSpecExtractor._extract_title_page_layout = _extract_title_page_layout
LayoutSpecExtractor._extract_date_declarations = _extract_date_declarations
LayoutSpecExtractor._extract_abstract_detail = _extract_abstract_detail
LayoutSpecExtractor._extract_keywords = _extract_keywords
LayoutSpecExtractor._extract_author_detail = _extract_author_detail
LayoutSpecExtractor._extract_custom_typo_commands = _extract_custom_typo_commands
LayoutSpecExtractor._extract_subfig_settings = _extract_subfig_settings
LayoutSpecExtractor._extract_name_redefinitions = _extract_name_redefinitions
LayoutSpecExtractor._extract_page_style_detail = _extract_page_style_detail
LayoutSpecExtractor._extract_geometry_config = _extract_geometry_config
LayoutSpecExtractor._extract_fontspec_config = _extract_fontspec_config
LayoutSpecExtractor._extract_titlesec_config = _extract_titlesec_config
LayoutSpecExtractor._extract_caption_package_config = _extract_caption_package_config
LayoutSpecExtractor._extract_biblatex_config = _extract_biblatex_config
LayoutSpecExtractor._extract_theorem_envs = _extract_theorem_envs
LayoutSpecExtractor._extract_algorithm_envs = _extract_algorithm_envs
LayoutSpecExtractor._extract_setspace_config = _extract_setspace_config
LayoutSpecExtractor._extract_enumerate_styles = _extract_enumerate_styles
LayoutSpecExtractor._extract_table_detail = _extract_table_detail
LayoutSpecExtractor._extract_marginpar = _extract_marginpar
LayoutSpecExtractor._extract_doi_url_format = _extract_doi_url_format
LayoutSpecExtractor._extract_cleveref_config = _extract_cleveref_config
LayoutSpecExtractor._extract_mode_options = _extract_mode_options


# ═══════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='LaTeX模板完整排版规格提取')
    parser.add_argument('cls_file', help='.cls文件路径')
    parser.add_argument('--extra', nargs='*', help='额外的.sty/.cfg文件')
    parser.add_argument('-o', '--output', help='输出目录')
    args = parser.parse_args()

    extractor = LayoutSpecExtractor(args.cls_file, args.extra)
    spec = extractor.extract_all()

    out_dir = args.output or os.path.splitext(args.cls_file)[0] + '_spec'
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(args.cls_file))[0]

    # JSON
    json_path = os.path.join(out_dir, f'{base}_layout_spec.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(spec, f, indent=2, ensure_ascii=False)
    print(f'[1/3] JSON保存: {json_path}')

    # Markdown报告
    report = generate_report(spec)
    md_path = os.path.join(out_dir, f'{base}_layout_report.md')
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(report)
    print(f'[2/3] Markdown报告: {md_path}')

    # Word样式
    word_styles = spec_to_word_styles(spec)
    result = OrderedDict([('layout_spec', spec), ('word_styles', word_styles)])
    ws_path = os.path.join(out_dir, f'{base}_word_styles.json')
    with open(ws_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    print(f'[3/3] Word样式映射: {ws_path}')

    # 打印摘要
    print('\n' + '=' * 60)
    print(f'排版规格提取完成: {base}')
    print(f'  基准字号: {spec["base_size_pt"]}pt')
    print(f'  栏数: {spec.get("columns", "?")}')
    print(f'  字体: {spec.get("fonts", {}).get("serif_name", "?")}')
    print(f'  章节标题: {len(spec.get("headings", {}))}级')
    print(f'  特殊环境: {len(spec.get("special_environments", {}))}个')
    print(f'  自定义命令: {len(spec.get("custom_commands", {}))}个')
    print(f'  依赖宏包: {len(spec.get("packages", {}))}个')
