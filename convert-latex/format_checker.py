#!/usr/bin/env python3
"""
LaTeX排版合规性检查器 v2.0
基于 template-extract-lite 的 spec.json 动态检查 .tex 文件是否符合期刊模板规范。
支持任意期刊模板，不硬编码。
"""
import re
import json
import sys
from pathlib import Path
from collections import OrderedDict


class ComplianceChecker:
    """基于模板spec的LaTeX合规性检查器"""

    def __init__(self, tex_path, spec_path=None, spec_dict=None):
        self.tex_path = Path(tex_path)
        self.tex_content = self.tex_path.read_text(encoding='utf-8')
        self.lines = self.tex_content.split('\n')
        self.results = OrderedDict()  # category -> {pass, issues, warnings}

        # 加载spec
        if spec_dict:
            self.spec = spec_dict
        elif spec_path:
            with open(spec_path, encoding='utf-8') as f:
                self.spec = json.load(f, object_pairs_hook=OrderedDict)
        else:
            self.spec = {}

    def check_all(self):
        """执行所有合规性检查"""
        print(f"合规性检查: {self.tex_path.name}")
        if self.spec:
            print(f"模板规格: 已加载 ({len(self.spec)} 类)")
        else:
            print("模板规格: 未加载（仅基础检查）")
        print("=" * 60)

        self.check_document_class()
        self.check_required_packages()
        self.check_title_format()
        self.check_author_format()
        self.check_abstract_format()
        self.check_keywords_format()
        self.check_section_commands()
        self.check_special_envs()
        self.check_figure_format()
        self.check_table_format()
        self.check_caption_format()
        self.check_equation_format()
        self.check_bibliography_format()
        self.check_appendix_format()
        self.check_template_specific()
        self.check_page_layout()
        self.check_paragraph_format()

        return self.results

    def _add(self, category, status, msg):
        if category not in self.results:
            self.results[category] = {'pass': [], 'issues': [], 'warnings': []}
        self.results[category][status].append(msg)

    # ─── 1. 文档类 ─────────────────────────────────────
    def check_document_class(self):
        cat = 'document_class'
        spec = self.spec.get('document_class', {})
        # 提取.tex中的\documentclass
        m = re.search(r'\\documentclass(?:\[([^\]]*)\])?\{(\w+)\}', self.tex_content)
        if not m:
            self._add(cat, 'issues', '未找到\\documentclass声明')
            return

        tex_opts = m.group(1) or ''
        tex_cls = m.group(2)

        # 检查类名
        expected_cls = spec.get('class_name', '')
        if expected_cls and tex_cls != expected_cls:
            self._add(cat, 'issues', f'文档类不匹配: 期望{expected_cls}, 实际{tex_cls}')
        else:
            self._add(cat, 'pass', f'文档类: {tex_cls}')

        # 检查选项
        default_opts = spec.get('default_options', [])
        if default_opts:
            tex_opt_list = [o.strip() for o in tex_opts.split(',') if o.strip()]
            for opt in default_opts:
                if opt not in tex_opt_list:
                    self._add(cat, 'warnings', f'缺少推荐选项: {opt}')

    # ─── 2. 必需包 ─────────────────────────────────────
    def check_required_packages(self):
        cat = 'required_packages'
        spec_pkgs = self.spec.get('required_packages', {})
        if not spec_pkgs:
            return

        # 提取.tex中所有\usepackage
        tex_pkgs = set()
        for m in re.finditer(r'\\usepackage(?:\s*\[[^\]]*\])?\s*\{([^}]+)\}', self.tex_content):
            for p in m.group(1).split(','):
                tex_pkgs.add(p.strip())

        # 从spec.required_packages动态获取cls内置包列表，而非硬编码
        # spec中列出的包就是cls通过RequirePackage加载的，不需要在.tex中重复声明
        rp = self.spec.get('required_packages', {}) if self.spec else {}
        cls_builtin = set(rp.keys())

        # .tex中必须手动声明的包（不在cls内置列表中的）
        for pkg in spec_pkgs:
            if pkg in cls_builtin:
                continue
            if pkg not in tex_pkgs:
                self._add(cat, 'warnings', f'模板要求包{pkg}但.tex中未声明')
            else:
                self._add(cat, 'pass', f'包{pkg}已声明')

        # 检查.tex中是否使用了与cls冲突的包
        # 从spec获取禁止的包列表，默认禁止geometry（大多数cls已定义页面布局）
        disallowed = self.spec.get('page_layout', {}).get('disallow_packages', ['geometry'])
        for pkg in disallowed:
            if pkg in tex_pkgs:
                self._add(cat, 'issues', f'不应使用{pkg}包，模板已定义相关布局')

    # ─── 3. 标题格式 ────────────────────────────────────
    def check_title_format(self):
        cat = 'title_format'
        spec = self.spec.get('title_format', {})
        # 检查\title命令
        title_m = re.search(r'\\title\s*(?:\[[^\]]*\])?\s*\{', self.tex_content)
        if not title_m:
            self._add(cat, 'issues', '缺少\\title命令')
        else:
            # 检查参数结构
            if spec.get('has_short_title') or spec.get('title_args', 0) > 0:
                short_m = re.search(r'\\title\s*\[[^\]]*\]\s*\{', self.tex_content)
                if not short_m:
                    self._add(cat, 'warnings', '模板支持短标题但未使用\\title[短标题]{标题}')
                else:
                    self._add(cat, 'pass', '\\title含短标题参数')
            else:
                self._add(cat, 'pass', '\\title命令存在')

    # ─── 4. 作者格式 ────────────────────────────────────
    def check_author_format(self):
        cat = 'author_format'
        spec = self.spec.get('author_format', {})
        # 检查\Author或\author
        if re.search(r'\\Author\b', self.tex_content):
            self._add(cat, 'pass', '\\Author命令存在')
        elif re.search(r'\\author\b', self.tex_content):
            self._add(cat, 'pass', '\\author命令存在')
        else:
            self._add(cat, 'issues', '缺少作者命令')

    # ─── 5. 摘要格式 ────────────────────────────────────
    def check_abstract_format(self):
        cat = 'abstract_format'
        spec = self.spec.get('abstract_format', {})
        expected_type = spec.get('type', 'environment')

        if expected_type == 'environment':
            if re.search(r'\\begin\{abstract\}', self.tex_content):
                self._add(cat, 'pass', '\\begin{abstract}环境存在')
            else:
                self._add(cat, 'issues', '缺少\\begin{abstract}环境')
        elif expected_type == 'command':
            if re.search(r'\\abstract\s*\{', self.tex_content):
                self._add(cat, 'pass', '\\abstract命令存在')
            else:
                self._add(cat, 'issues', '缺少\\abstract命令')

    # ─── 6. 关键词格式 ──────────────────────────────────
    def check_keywords_format(self):
        cat = 'keywords_format'
        spec = self.spec.get('keywords_format', {})
        if not spec:
            return
        if re.search(r'\\keywords\s*\{', self.tex_content):
            self._add(cat, 'pass', '\\keywords命令存在')
        elif re.search(r'\\begin\{keywords\}', self.tex_content):
            self._add(cat, 'pass', '\\begin{keywords}环境存在')
        else:
            self._add(cat, 'warnings', '模板要求\\keywords但.tex中未找到')

    # ─── 7. 章节命令 ────────────────────────────────────
    def check_section_commands(self):
        cat = 'section_commands'
        spec = self.spec.get('section_commands', {})
        if not spec:
            return
        for cmd, info in spec.items():
            if info.get('alias_of'):
                alias = info['alias_of']
                # 检查是否使用了别名命令（如\introduction代替\section）
                if re.search(r'\\' + cmd + r'\b', self.tex_content):
                    self._add(cat, 'pass', f'\\{cmd}已使用（别名→\\{alias}）')
                # 也检查原始命令
                elif re.search(r'\\' + alias + r'\b', self.tex_content):
                    self._add(cat, 'pass', f'\\{alias}已使用（\\{cmd}的别名目标）')

    # ─── 8. 特殊声明环境 ────────────────────────────────
    def check_special_envs(self):
        cat = 'special_envs'
        spec = self.spec.get('special_envs', {})
        if not spec:
            return
        for env_name, info in spec.items():
            tp = info.get('type', 'environment')
            if tp == 'environment':
                if re.search(r'\\begin\{' + env_name + r'\}', self.tex_content):
                    self._add(cat, 'pass', f'\\begin{{{env_name}}}存在')
                else:
                    # 声明段落缺失是warning而非error（有些可选）
                    self._add(cat, 'warnings', f'声明环境{env_name}未使用')
            elif tp == 'command':
                if re.search(r'\\' + env_name + r'\b', self.tex_content):
                    self._add(cat, 'pass', f'\\{env_name}命令存在')
                else:
                    self._add(cat, 'warnings', f'声明命令\\{env_name}未使用')

    # ─── 9. 图格式 ─────────────────────────────────────
    def check_figure_format(self):
        cat = 'figure_format'
        spec = self.spec.get('figure_format', {})

        fig_count = self.tex_content.count('\\begin{figure}')
        print(f"  图片数量: {fig_count}")

        if fig_count == 0:
            return

        # 检查图片位置参数
        default_pos = spec.get('default_position', 'htbp')
        fig_positions = re.findall(r'\\begin\{figure\}(\[[^\]]*\])', self.tex_content)
        for pos in fig_positions:
            if default_pos and default_pos not in pos and 'H' not in pos:
                self._add(cat, 'warnings', f'图片位置参数{pos}与模板默认[{default_pos}]不一致')

        # 检查\includegraphics路径
        img_paths = re.findall(r'\\includegraphics[^\{]*\{([^}]+)\}', self.tex_content)
        expected_path = spec.get('graphics_path', '')
        for img_path in img_paths:
            if expected_path and not img_path.startswith(expected_path.strip('{}')):
                self._add(cat, 'warnings', f'图片路径{img_path}不在模板指定目录')

        # 检查子图包（cls内置的包不需要在.tex中重复声明）
        sub_pkg = spec.get('subfigure_package', '')
        if sub_pkg and sub_pkg not in self.spec.get('required_packages', {}):
            # 只有spec的required_packages中没有的包才需要.tex手动声明
            if sub_pkg not in self.tex_content:
                self._add(cat, 'warnings', f'模板使用{sub_pkg}子图包但.tex中未加载')

        self._add(cat, 'pass', f'{fig_count}个figure环境')

    # ─── 10. 表格式（由增强版check_table_format处理）──────────────

    # ─── 11. Caption格式（由增强版check_caption_format处理）────────

    # ─── 12. 公式格式 ────────────────────────────────────
    def check_equation_format(self):
        cat = 'equation_format'
        spec = self.spec.get('equation_format', {})

        eq_count = self.tex_content.count('\\begin{equation}')
        gather_count = self.tex_content.count('\\begin{gather}')
        align_count = self.tex_content.count('\\begin{align}')
        print(f"  公式数量: equation={eq_count}, gather={gather_count}, align={align_count}")

        # 检查amsmath
        if not re.search(r'\\usepackage.*\{amsmath\}', self.tex_content):
            if eq_count + gather_count + align_count > 0:
                self._add(cat, 'warnings', '使用了公式环境但未显式加载amsmath（可能由cls内置）')

        # 检查公式编号格式
        numbering = spec.get('numbering_format', '')
        number_within = spec.get('number_within', '')
        if number_within and not re.search(r'\\numberwithin\{equation\}', self.tex_content):
            self._add(cat, 'pass', f'公式编号由cls控制(numberwithin={number_within})')

        # 检查空公式行（gather中的幽灵编号）
        gather_envs = re.findall(r'\\begin\{gather\}(.*?)\\end\{gather\}', self.tex_content, re.DOTALL)
        for i, env in enumerate(gather_envs):
            for line in env.split('\n'):
                stripped = line.strip()
                if stripped == '\\\\' or (stripped and not stripped.startswith('%')
                        and stripped != '\\\\'
                        and re.match(r'^\\label\{', stripped) is None
                        and stripped.endswith('\\\\')
                        and len(stripped) <= 4):
                    self._add(cat, 'warnings', f'gather环境#{i+1}可能含空公式行')

    # ─── 13. 参考文献格式 ────────────────────────────────
    def check_bibliography_format(self):
        cat = 'bibliography_format'
        spec = self.spec.get('bibliography_format', {})

        # 检查\bibliographystyle
        bibstyle_m = re.search(r'\\bibliographystyle\{([^}]+)\}', self.tex_content)
        if not bibstyle_m:
            self._add(cat, 'issues', '缺少\\bibliographystyle命令')
        else:
            expected_bst = spec.get('bst_file', '')
            if expected_bst and bibstyle_m.group(1) != expected_bst:
                self._add(cat, 'warnings', f'bibliographystyle为{bibstyle_m.group(1)}，模板默认{expected_bst}')
            else:
                self._add(cat, 'pass', f'\\bibliographystyle{{{bibstyle_m.group(1)}}}')

        # 检查\bibliography
        if not re.search(r'\\bibliography\{', self.tex_content):
            if spec.get('style') != 'biblatex':
                self._add(cat, 'issues', '缺少\\bibliography命令')

        # 检查natbib/biblatex
        if spec.get('style') == 'natbib':
            if not re.search(r'\\usepackage.*\{natbib\}', self.tex_content):
                self._add(cat, 'pass', 'natbib由cls内置')
        elif spec.get('style') == 'biblatex':
            if not re.search(r'\\usepackage.*\{biblatex\}', self.tex_content):
                self._add(cat, 'warnings', '模板使用biblatex但.tex中未加载')

    # ─── 14. 附录格式 ────────────────────────────────────
    def check_appendix_format(self):
        cat = 'appendix_format'
        spec = self.spec.get('appendix_format', {})
        if not spec:
            return
        # 附录检查为可选，不强制
        if re.search(r'\\appendix\b', self.tex_content) or re.search(r'\\begin\{appendix\}', self.tex_content):
            self._add(cat, 'pass', '附录声明存在')

    # ─── 15. 模板特有命令 ────────────────────────────────
    def check_template_specific(self):
        cat = 'template_specific'
        spec = self.spec.get('template_specific', {})
        if not spec:
            return

        # 检查必填声明
        required = spec.get('required_declarations', [])
        for decl in required:
            if re.search(r'\\' + decl + r'\b', self.tex_content) or re.search(r'\\begin\{' + decl + r'\}', self.tex_content):
                self._add(cat, 'pass', f'必填声明{decl}已使用')
            else:
                self._add(cat, 'warnings', f'必填声明{decl}未使用')

        # 检查\maketitle
        if '\\maketitle' not in self.tex_content:
            self._add(cat, 'issues', '缺少\\maketitle命令')
        else:
            self._add(cat, 'pass', '\\maketitle存在')

    # ─── 页面布局（通用检查）─────────────────────────────
    def check_page_layout(self):
        cat = 'page_layout'
        # 不应手动设置页面布局（由cls控制）
        if re.search(r'\\usepackage.*\{geometry\}', self.tex_content):
            self._add(cat, 'issues', '不应使用geometry包，模板已定义页面布局')

        if re.search(r'\\setlength.*\\textwidth', self.tex_content):
            self._add(cat, 'warnings', '手动设置\\textwidth，可能与模板冲突')

        if re.search(r'\\setlength.*\\textheight', self.tex_content):
            self._add(cat, 'warnings', '手动设置\\textheight，可能与模板冲突')

    # ─── 段落格式（通用检查）─────────────────────────────
    def check_paragraph_format(self):
        cat = 'paragraph_format'
        # 行间距: 如果spec指定了非single行距，检查是否设置了
        body_spec = self.spec.get('body_text', {})
        line_spacing = body_spec.get('line_spacing', '')
        if line_spacing and line_spacing != 'single':
            if '\\usepackage{setspace}' in self.tex_content:
                # 基础行间距命令映射，可从spec扩展
                spacing_cmd_map = {'1.5': 'onehalfspacing', 'onehalf': 'onehalfspacing',
                                   'double': 'doublespacing', '2': 'doublespacing',
                                   '1': 'singlespacing', 'single': 'singlespacing'}
                spec_spacing = body_spec.get('spacing_commands', {})
                for k, v in spec_spacing.items():
                    spacing_cmd_map[k] = v.lstrip('\\')
                expected_cmd = spacing_cmd_map.get(str(line_spacing).lower(), 'setstretch')
                if expected_cmd in self.tex_content:
                    self._add(cat, 'pass', f'行间距设置正确: {line_spacing}')
                else:
                    self._add(cat, 'warnings', f'模板要求行间距{line_spacing}但未设置')
            else:
                self._add(cat, 'warnings', f'模板要求行间距{line_spacing}但未加载setspace包')
        else:
            # cls已控制行距（默认single），不应手动设置
            if '\\onehalfspacing' in self.tex_content or '\\doublespacing' in self.tex_content:
                self._add(cat, 'issues', '不应手动设置行间距，模板已定义')

        # \noindent过多
        noindent_count = self.tex_content.count('\\noindent')
        if noindent_count > 5:
            self._add(cat, 'warnings', f'使用了{noindent_count}次\\noindent，可能影响排版')

        # 中文支持
        if not re.search(r'\\usepackage.*\{ctex\}', self.tex_content):
            has_chinese = bool(re.search(r'[一-鿿]', self.tex_content))
            if has_chinese:
                self._add(cat, 'issues', '含中文内容但未加载ctex包')

    # ─── Caption格式增强检查 ──────────────────────────────
    def check_caption_format(self):
        cat = 'caption_format'
        # 从多个spec源获取caption格式（与SpecAdapter一致）
        cap_fmt = self.spec.get('caption_format', {})
        fig_fmt = self.spec.get('figure_format', {})
        tbl_fmt = self.spec.get('table_format', {})
        if not cap_fmt and not fig_fmt and not tbl_fmt:
            return

        # 检查分隔符（从spec获取，默认"."）
        sep = cap_fmt.get('separator', '')
        if sep:
            if f'separator={sep}' in self.tex_content:
                self._add(cat, 'pass', f'Caption分隔符设置正确: {sep}')
            elif sep == '.' and re.search(r'\\captionsetup.*separator', self.tex_content):
                self._add(cat, 'pass', 'Caption分隔符已设置')
            else:
                if f'\\captionsetup{{separator={sep}}}' not in self.tex_content:
                    self._add(cat, 'warnings', f'模板要求Caption分隔符{sep}但未设置')

        # 检查label粗体（从spec获取）
        label_weight = cap_fmt.get('label_weight', '')
        if label_weight == 'bold':
            has_labelbf = ('labelfont=bf' in self.tex_content or 'label=bf' in self.tex_content)
            if has_labelbf:
                self._add(cat, 'pass', 'Caption标签粗体已设置')
            elif not re.search(r'\\textbf\{.*?\}.*\\caption', self.tex_content):
                self._add(cat, 'warnings', '模板要求Caption标签粗体但未设置')

        # 检查caption字体字号（从spec获取）
        cap_font = cap_fmt.get('font_size', '')
        if cap_font:
            fig_envs = re.findall(r'\\begin\{figure\}.*?\\end\{figure\}', self.tex_content, re.DOTALL)
            for env in fig_envs[:3]:
                if f'\\{cap_font}' in env:
                    self._add(cat, 'pass', f'Caption字号{cap_font}已应用')
                    break
            else:
                self._add(cat, 'warnings', f'模板要求Caption字号{cap_font}但未应用')

        # 检查caption位置（从spec获取，有默认值）
        fig_pos = cap_fmt.get('figure_position', '') or fig_fmt.get('caption_position', '')
        tbl_pos = cap_fmt.get('table_position', '') or tbl_fmt.get('caption_position', '')
        if fig_pos:
            if f'position={fig_pos}' in self.tex_content or \
               (fig_pos == 'below' and '\\captionsetup[figure]' not in self.tex_content):
                self._add(cat, 'pass', f'Figure caption位置: {fig_pos}')
            else:
                self._add(cat, 'warnings', f'模板要求figure caption在{fig_pos}位置')
        if tbl_pos:
            if f'position={tbl_pos}' in self.tex_content or \
               (tbl_pos == 'above' and '\\captionsetup[table]' not in self.tex_content):
                self._add(cat, 'pass', f'Table caption位置: {tbl_pos}')
            else:
                self._add(cat, 'warnings', f'模板要求table caption在{tbl_pos}位置')

    # ─── 表格字体检查 ────────────────────────────────────
    def check_table_format(self):
        cat = 'table_format'
        spec = self.spec.get('table_format', {})
        cap_fmt = self.spec.get('caption_format', {})

        table_count = self.tex_content.count('\\begin{table}')
        tikz_count = self.tex_content.count('\\begin{tikzpicture}')
        print(f"  表格数量: {table_count} (TikZ: {tikz_count})")

        # 检查表格字体（从spec获取）
        tbl_body_font = spec.get('body_size', '')
        tbl_header_font = spec.get('header_size', '')
        if tbl_body_font:
            tikz_nodes = re.findall(r'font=\\\\?(\w+)', self.tex_content)
            if tbl_body_font in tikz_nodes:
                self._add(cat, 'pass', f'表格体字号{tbl_body_font}已应用')
            else:
                self._add(cat, 'warnings', f'模板要求表格体字号{tbl_body_font}但未应用')
        if tbl_header_font:
            tikz_nodes = re.findall(r'font=\\\\?(\w+)', self.tex_content)
            if tbl_header_font in tikz_nodes:
                self._add(cat, 'pass', f'表头字号{tbl_header_font}已应用')
            else:
                self._add(cat, 'warnings', f'模板要求表头字号{tbl_header_font}但未应用')

        # 检查表格浮动位置（从spec获取）
        tbl_float = spec.get('float_position', '')
        if tbl_float:
            tbl_positions = re.findall(r'\\begin\{table\}(\[[^\]]*\])', self.tex_content)
            for pos in tbl_positions:
                if tbl_float not in pos and 'H' not in pos:
                    self._add(cat, 'warnings', f'表格位置参数{pos}与模板默认[{tbl_float}]不一致')

        # 检查表格对齐方式（从spec获取）
        tbl_align = spec.get('alignment', '')
        if tbl_align == 'center' and '\\centering' not in self.tex_content:
            if table_count > 0:
                self._add(cat, 'warnings', '模板要求表格居中对齐但未使用\\centering')

        # 检查表格相关包（从spec获取）
        table_pkgs = spec.get('table_packages', [])
        for pkg in table_pkgs:
            if pkg in self.spec.get('required_packages', {}):
                continue
            if pkg not in self.tex_content:
                self._add(cat, 'warnings', f'模板使用{pkg}包但.tex中未加载')

        if table_count > 0 or tikz_count > 0:
            self._add(cat, 'pass', f'{table_count}个table + {tikz_count}个tikzpicture')

    # ─── 报告 ────────────────────────────────────────────
    def print_report(self):
        """打印合规性报告"""
        print("\n" + "=" * 60)
        print("合规性检查报告")
        print("=" * 60)

        total_pass = 0
        total_issues = 0
        total_warnings = 0

        for cat, result in self.results.items():
            passes = result.get('pass', [])
            issues = result.get('issues', [])
            warnings = result.get('warnings', [])
            total_pass += len(passes)
            total_issues += len(issues)
            total_warnings += len(warnings)

            if issues or warnings:
                print(f"\n[{cat}]")
                for iss in issues:
                    print(f"  [ERROR] {iss}")
                for warn in warnings:
                    print(f"  [WARN] {warn}")
                for p in passes:
                    print(f"  [OK] {p}")

        print(f"\n{'=' * 60}")
        print(f"通过: {total_pass}  警告: {total_warnings}  错误: {total_issues}")

        if total_issues == 0 and total_warnings == 0:
            print("[PASS] All compliance checks passed!")
        elif total_issues == 0:
            print("[OK] No errors, but warnings need attention")

        print("=" * 60)
        return total_issues, total_warnings

    def to_json(self, output_path):
        """输出结构化JSON报告"""
        report = OrderedDict()
        report['tex_file'] = str(self.tex_path)
        report['spec_source'] = 'template-extract-lite'
        report['categories'] = self.results

        total_issues = sum(len(r.get('issues', [])) for r in self.results.values())
        total_warnings = sum(len(r.get('warnings', [])) for r in self.results.values())
        total_pass = sum(len(r.get('pass', [])) for r in self.results.values())
        report['summary'] = {
            'pass': total_pass,
            'warnings': total_warnings,
            'issues': total_issues,
            'status': 'PASS' if total_issues == 0 else 'FAIL',
        }

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        return report


def main():
    import argparse
    parser = argparse.ArgumentParser(description='LaTeX排版合规性检查器 v2.0')
    parser.add_argument('tex_file', help='待检查的.tex文件')
    parser.add_argument('--spec', help='template-extract-lite输出的spec.json')
    parser.add_argument('--output', help='输出JSON报告路径')
    args = parser.parse_args()

    checker = ComplianceChecker(args.tex_file, spec_path=args.spec)
    checker.check_all()
    issues, warnings = checker.print_report()

    if args.output:
        checker.to_json(args.output)
        print(f"\n报告已保存: {args.output}")

    sys.exit(1 if issues > 0 else 0)


if __name__ == '__main__':
    main()
