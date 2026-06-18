#!/usr/bin/env python3
"""SpecAdapter: template-extract-lite spec → layout_spec 适配层

将template-extract-lite的spec.json适配为convert_direct.py所需的layout_spec。

设计原则：
- cls控制编译时格式（行距、字号、边距），preamble只补xelatex缺失的字体设置
- spec提取的是"必须在.tex中正确书写"的结构信息，不是编译时自动应用的排版细节
"""

import re
import shutil
import subprocess
import tempfile
import os
from pathlib import Path


def _select_probe_engine(cls_name):
    try:
        from tex_postprocess import _select_engine
        return _select_engine(cls_name)
    except Exception:
        lualatex_cls = {
            'copernicus', 'acp', 'amt', 'angeo', 'bg', 'cp', 'esd', 'essd',
            'gmd', 'hess', 'nhess', 'os', 'se', 'tc', 'wcd',
        }
        cls_lower = Path(str(cls_name or '')).stem.lower()
        return 'lualatex' if 'copernicus' in cls_lower or cls_lower in lualatex_cls else 'xelatex'


class SpecAdapter:
    """将template-extract-lite的spec.json适配为convert_direct.py所需的layout_spec"""

    def __init__(self, spec, skeleton_info, cls_path=None, template_dir=None,
                 config_mode=None):
        self.spec = spec
        self.skeleton_info = skeleton_info
        self.cls_path = cls_path
        self.template_dir = template_dir
        self.config_mode = config_mode
        self._probe_log_cache = None

    def _effective_config_mode(self):
        try:
            from skeleton_builder import normalize_config_mode
            return normalize_config_mode(self.config_mode)
        except Exception:
            return 'manuscript' if self.config_mode in (None, '', 'classic') else self.config_mode

    def to_layout_spec(self):
        """从spec推导layout_spec（兼容旧版5个使用点）"""
        layout_spec = {}
        layout_spec['numbering'] = self._derive_numbering()
        layout_spec['body_text'] = self._derive_body_text_spec()
        layout_spec['spacing'] = {}     # cls已控制行距等，不手动设置
        layout_spec['figure'] = self._derive_figure_spec()
        layout_spec['caption'] = self._derive_caption_spec()
        layout_spec['table'] = self._derive_table_spec()
        layout_spec['bibliography'] = self._derive_bibliography_spec()
        layout_spec['document'] = self._derive_document_spec()
        layout_spec['abstract'] = self._derive_abstract_spec()
        layout_spec['equation'] = self._derive_equation_spec()
        layout_spec['page_geometry'] = self._derive_page_geometry()
        layout_spec['float_policy'] = self._derive_float_policy()
        return layout_spec

    def _derive_page_geometry(self):
        if not self.cls_path:
            return {}
        try:
            from pathlib import Path
            import sys
            skill_dir = Path(__file__).resolve().parent.parent
            if str(skill_dir) not in sys.path:
                sys.path.insert(0, str(skill_dir))
            from shared.template_config import get_page_geometry_for_mode

            content = Path(self.cls_path).read_text(encoding='utf-8', errors='ignore')
            # 使用实际config_mode，而非硬编码'manuscript'
            static_geo = get_page_geometry_for_mode(
                content, config_mode=self._effective_config_mode()) or {}
            probe_geo = {}
            if self._page_geometry_needs_probe(static_geo, content):
                probe_geo = self._probe_page_geometry()
            if static_geo and probe_geo:
                merged = dict(static_geo)
                for key, value in probe_geo.items():
                    if value is not None:
                        merged[key] = value
                return self._with_effective_columns(merged, content)
            return self._with_effective_columns(static_geo or probe_geo, content)
        except Exception:
            return self._probe_page_geometry()

    def _page_geometry_needs_probe(self, geo, cls_content=''):
        if not geo:
            return True
        paper = float(geo.get('paperwidth_mm') or 0)
        text = float(geo.get('textwidth_mm') or 0)
        if text and paper and paper < text:
            return True
        if float(geo.get('right_margin_mm') or 0) < 0:
            return True
        column_count = int(geo.get('column_count', 1) or 1)
        if column_count >= 2:
            return True
        if column_count < 2:
            return self._single_column_geometry_needs_probe(cls_content)
        return False

    def _single_column_geometry_needs_probe(self, cls_content=''):
        try:
            from skeleton_builder import class_options_from_spec
            dc = self.spec.get('document_class', {}) if self.spec else {}
            cls_name = dc.get('class_name', 'article')
            options = {
                str(item).lower()
                for item in class_options_from_spec(self.spec, cls_name, self.config_mode)
            }
            if 'onecolumn' in options:
                return False
            if 'twocolumn' in options:
                return True
            content = cls_content or ''
            mode = self._effective_config_mode()
            if mode == 'final' and re.search(r'\\(?:@twocolumntrue|twocolumn)\b', content):
                return True
            defaults = {str(item).lower() for item in dc.get('default_options', [])}
            return 'twocolumn' in defaults
        except Exception:
            return False

    def _with_effective_columns(self, geo, cls_content):
        if not geo:
            return geo
        try:
            import sys
            skill_dir = Path(__file__).resolve().parent.parent
            if str(skill_dir) not in sys.path:
                sys.path.insert(0, str(skill_dir))
            from shared.template_config import detect_effective_column_count

            from skeleton_builder import class_options_from_spec

            dc = self.spec.get('document_class', {}) if self.spec else {}
            cls_name = dc.get('class_name', 'article')
            options = class_options_from_spec(self.spec, cls_name, self.config_mode)
            columns = detect_effective_column_count(cls_content, options)
            if columns > int(geo.get('column_count', 1) or 1):
                geo = dict(geo)
                geo['column_count'] = columns
        except Exception:
            pass
        return geo

    def _probe_page_geometry(self):
        """Ask the target class for effective dimensions when static CLS parsing is incomplete."""
        return self._geometry_from_probe_log(self._class_probe_log())

    def _class_probe_log(self):
        """Compile one template probe and reuse its effective class values."""
        if self._probe_log_cache is not None:
            return self._probe_log_cache
        if not self.template_dir:
            self._probe_log_cache = ''
            return self._probe_log_cache
        template_dir = Path(self.template_dir)
        cls_name = self.spec.get('document_class', {}).get('class_name') or Path(self.cls_path).stem
        try:
            from skeleton_builder import class_options_from_spec
            opts = class_options_from_spec(
                self.spec, cls_name, self.config_mode)
            option_text = '[' + ','.join(opts) + ']' if opts else ''
            with tempfile.TemporaryDirectory(prefix='skill_geometry_probe_') as tmp:
                work_dir = Path(tmp)
                self._copy_probe_support_files(template_dir, work_dir)
                if self.cls_path:
                    cls_src = Path(self.cls_path)
                    if cls_src.exists():
                        shutil.copy2(str(cls_src), str(work_dir / cls_src.name))
                probe = work_dir / '_skill_geometry_probe.tex'
                probe.write_text(
                    '\\documentclass' + option_text + '{' + cls_name + '}\n'
                    '\\begin{document}\n'
                    '\\typeout{SKILL-PROBE-TEXTWIDTH=\\the\\textwidth}\n'
                    '\\typeout{SKILL-PROBE-TEXTHEIGHT=\\the\\textheight}\n'
                    '\\typeout{SKILL-PROBE-PAPERWIDTH=\\the\\paperwidth}\n'
                    '\\typeout{SKILL-PROBE-PAPERHEIGHT=\\the\\paperheight}\n'
                    '\\typeout{SKILL-PROBE-ODDSIDEMARGIN=\\the\\oddsidemargin}\n'
                    '\\typeout{SKILL-PROBE-TOPMARGIN=\\the\\topmargin}\n'
                    '\\typeout{SKILL-PROBE-LEFTMARGIN=\\the\\dimexpr1in+\\hoffset+\\oddsidemargin\\relax}\n'
                    '\\typeout{SKILL-PROBE-TEXTTOPMARGIN=\\the\\dimexpr1in+\\voffset+\\topmargin+\\headheight+\\headsep\\relax}\n'
                    '\\typeout{SKILL-PROBE-COLUMNWIDTH=\\the\\columnwidth}\n'
                    '\\typeout{SKILL-PROBE-COLUMNSEP=\\the\\columnsep}\n'
                    '\\makeatletter\n'
                    '\\typeout{SKILL-PROBE-TOPFRACTION=\\topfraction}\n'
                    '\\typeout{SKILL-PROBE-TEXTFRACTION=\\textfraction}\n'
                    '\\typeout{SKILL-PROBE-FLOATPAGEFRACTION=\\floatpagefraction}\n'
                    '\\typeout{SKILL-PROBE-DBLTOPFRACTION=\\dbltopfraction}\n'
                    '\\typeout{SKILL-PROBE-DBLFLOATPAGEFRACTION=\\dblfloatpagefraction}\n'
                    '\\makeatother\n'
                    '\\end{document}\n',
                    encoding='utf-8',
                )
                engine = _select_probe_engine(cls_name)
                env = None
                if engine == 'lualatex':
                    env = os.environ.copy()
                    cache_dir = work_dir / '.texmf-var'
                    cache_dir.mkdir(parents=True, exist_ok=True)
                    env['TEXMFVAR'] = str(cache_dir.resolve())
                subprocess.run(
                    [engine, '-interaction=nonstopmode', probe.name],
                    cwd=str(work_dir), capture_output=True, env=env, timeout=120,
                )
                log = probe.with_suffix('.log').read_text(
                    encoding='utf-8', errors='ignore')
                self._probe_log_cache = log
                return log
        except Exception:
            self._probe_log_cache = ''
            return self._probe_log_cache

    def _derive_float_policy(self):
        """Extract effective float thresholds without changing the journal class."""
        names = (
            'topfraction', 'textfraction', 'floatpagefraction',
            'dbltopfraction', 'dblfloatpagefraction',
        )
        evidence = self._template_float_evidence_text()
        policy = {}
        for name in names:
            pattern = (
                rf'\\(?:def|renewcommand)\s*\{{?\\{name}\}}?'
                rf'\s*\{{\s*([0-9]*\.?[0-9]+)\s*\}}'
            )
            match = re.search(pattern, evidence)
            if match:
                policy[name] = float(match.group(1))

        missing = [name for name in names if name not in policy]
        if missing:
            log = self._class_probe_log()
            for name in missing:
                probe_name = name.upper()
                match = re.search(
                    rf'SKILL-PROBE-{probe_name}=([0-9]*\.?[0-9]+)', log)
                if match:
                    policy[name] = float(match.group(1))
        return policy

    def _copy_probe_support_files(self, template_dir, work_dir):
        suffixes = {
            '.cls', '.sty', '.cfg', '.clo', '.fd', '.def',
            '.ldf', '.bbd', '.bbx', '.cbx', '.lbx', '.bst',
        }
        for base in (template_dir, Path(self.cls_path).parent if self.cls_path else None):
            if not base or not base.exists():
                continue
            for src in base.iterdir():
                if src.is_file() and src.suffix.lower() in suffixes:
                    dst = work_dir / src.name
                    if not dst.exists():
                        shutil.copy2(str(src), str(dst))

    def _geometry_from_probe_log(self, log_text):
        values = {}
        for name in (
            'TEXTWIDTH', 'TEXTHEIGHT', 'PAPERWIDTH', 'PAPERHEIGHT',
            'ODDSIDEMARGIN', 'TOPMARGIN', 'LEFTMARGIN', 'TEXTTOPMARGIN',
            'COLUMNWIDTH', 'COLUMNSEP'
        ):
            m = re.search(rf'SKILL-PROBE-{name}=(-?[\d.]+)pt', log_text)
            if m:
                values[name] = float(m.group(1)) * 25.4 / 72.27
        if not values.get('TEXTWIDTH'):
            return {}
        textwidth = values['TEXTWIDTH']
        columnwidth = values.get('COLUMNWIDTH') or textwidth
        columnsep = values.get('COLUMNSEP') or 0.0
        column_count = 1
        if columnwidth and columnwidth < textwidth * 0.8:
            column_count = max(2, round((textwidth + columnsep) / (columnwidth + columnsep)))
        paperwidth = values.get('PAPERWIDTH')
        paperheight = values.get('PAPERHEIGHT')
        left_margin = values.get('ODDSIDEMARGIN')
        effective_left = values.get('LEFTMARGIN')
        if self._margin_fits(effective_left, paperwidth, textwidth):
            left_margin = effective_left
        top_margin = values.get('TOPMARGIN')
        effective_top = values.get('TEXTTOPMARGIN')
        if self._margin_fits(effective_top, paperheight, values.get('TEXTHEIGHT')):
            top_margin = effective_top
        right_margin = None
        if paperwidth and left_margin is not None:
            right_margin = paperwidth - textwidth - left_margin
        return {
            'paperwidth_mm': round(paperwidth, 2) if paperwidth else None,
            'paperheight_mm': round(paperheight, 2) if paperheight else None,
            'textwidth_mm': round(textwidth, 2),
            'textheight_mm': round(values.get('TEXTHEIGHT'), 2) if values.get('TEXTHEIGHT') else None,
            'oddsidemargin_mm': round(left_margin, 2) if left_margin is not None else None,
            'right_margin_mm': round(right_margin, 2) if right_margin is not None else None,
            'topmargin_mm': round(top_margin, 2) if top_margin is not None else None,
            'column_count': int(column_count),
            'column_sep_mm': round(columnsep, 2),
        }

    @staticmethod
    def _margin_fits(margin, paper, text):
        if margin is None or paper is None or text is None:
            return False
        return 0 <= margin <= max(paper - text, 0)

    def _cleanup_probe_files(self, directory, stem):
        for suffix in ('.aux', '.log', '.out', '.pdf', '.tex'):
            path = directory / f'{stem}{suffix}'
            try:
                if path.exists():
                    path.unlink()
            except OSError:
                pass

    def _derive_double_column_float_support(self, kind):
        """Derive figure*/table* support evidence from the template files."""
        evidence = self._template_float_evidence_text()
        if re.search(rf'\\newenvironment\{{{kind}\*\}}|\\@dblfloat\{{{kind}\}}', evidence):
            return {'supports': True, 'source': f'explicit-class-{kind}-star'}
        if re.search(rf'\\begin\{{{kind}\*\}}', evidence):
            return {'supports': True, 'source': f'explicit-example-{kind}-star'}
        if re.search(r'\\begin\{strip\}', evidence):
            return {'supports': True, 'source': 'explicit-example-strip'}
        if (re.search(r'\\LoadClass(?:\[[^\]]*\])?\{article\}', evidence)
                and (self._template_accepts_twocolumn() or self._template_uses_twocolumn())):
            return {'supports': True, 'source': 'inherited-article-twocolumn'}
        return {'supports': False, 'source': 'not-detected'}

    @staticmethod
    def _full_width_container(float_support, kind=None):
        if not float_support.get('supports'):
            return ''
        if kind in ('figure', 'table'):
            return 'native-star'
        return 'strip'

    def _derive_float_position(self, kind, starred=False):
        """Derive a float placement option from template examples."""
        evidence = self._template_float_evidence_text()
        star = r'\*' if starred else ''
        pattern = rf'\\begin\{{{kind}{star}\}}\s*(?:%\s*)?\[([^\]]+)\]'
        match = re.search(pattern, evidence)
        return match.group(1).strip() if match else ''

    def _template_float_evidence_text(self):
        chunks = []
        if self.cls_path:
            try:
                chunks.append(Path(self.cls_path).read_text(encoding='utf-8', errors='ignore'))
            except OSError:
                pass
        if self.template_dir:
            template_dir = Path(self.template_dir)
            for pattern in ('*.tex', '*.dtx'):
                for path in template_dir.glob(pattern):
                    if path.name.endswith('_full.tex') or path.name.startswith('_skill_'):
                        continue
                    try:
                        chunks.append(path.read_text(encoding='utf-8', errors='ignore'))
                    except OSError:
                        pass
        return '\n'.join(chunks)

    def _template_accepts_twocolumn(self):
        dc = self.spec.get('document_class', {})
        declared = set(dc.get('declared_options', []) or [])
        defaults = set(dc.get('default_options', []) or [])
        return '*' in declared or 'twocolumn' in declared or 'twocolumn' in defaults

    def _template_uses_twocolumn(self):
        if not self.cls_path:
            return False
        try:
            content = Path(self.cls_path).read_text(encoding='utf-8', errors='ignore')
            return self._detect_twocolumn_from_cls('', []) or bool(
                re.search(r'\\(?:@twocolumntrue|twocolumn)\b', content)
            )
        except OSError:
            return False

    def _derive_numbering(self):
        eq_fmt = self.spec.get('equation_format', {})
        fig_fmt = self.spec.get('figure_format', {})
        tbl_fmt = self.spec.get('table_format', {})
        numbering = {}
        # 各自独立读取 numbering_format
        eq_nf = eq_fmt.get('numbering_format', '')
        fig_nf = fig_fmt.get('numbering_format', '')
        tbl_nf = tbl_fmt.get('numbering_format', '')
        if eq_nf:
            numbering['equation_format'] = eq_nf
        if fig_nf:
            numbering['figure_format'] = fig_nf
        if tbl_nf:
            numbering['table_format'] = tbl_nf
        # number_within
        num_within = eq_fmt.get('number_within', '')
        if num_within:
            numbering['equation_within'] = num_within
        # 从 numbering spec 的条件重置推断模式标记
        num_spec = self.spec.get('numbering', {})
        if isinstance(num_spec, dict):
            # 存储无条件/条件重置信息供模式检测使用
            numbering['_unconditional_resets'] = num_spec.get('unconditional_resets', [])
            numbering['_conditional_resets'] = num_spec.get('conditional_resets', [])
        return numbering

    def _derive_figure_spec(self):
        fig_fmt = self.spec.get('figure_format', {})
        cap_fmt = self.spec.get('caption_format', {})
        spec = {}
        # 浮动位置: 从spec获取，默认htbp
        spec['float_position'] = fig_fmt.get('default_position', 'htbp')
        # caption位置: 从caption_format或figure_format获取
        fig_cap_pos = cap_fmt.get('figure_position', '') or fig_fmt.get('caption_position', '')
        if fig_cap_pos:
            spec['caption_position'] = fig_cap_pos
        else:
            spec['caption_position'] = 'below'  # LaTeX默认figure caption在下方
        # 图片宽度: 从spec获取; 否则按模板规定动态判断
        # CLS规定: figure用\columnwidth(半栏), figure*用\textwidth(跨栏)
        # 我们使用figure(非starred), 所以双栏时用\columnwidth
        if fig_fmt.get('width'):
            spec['width'] = fig_fmt['width']
        else:
            is_twocolumn = self._detect_twocolumn_from_cls(
                self.spec.get('document_class', {}).get('base_class_options', '') or '',
                self.spec.get('document_class', {}).get('default_options', [])
            )
            spec['width'] = '\\columnwidth' if is_twocolumn else '\\textwidth'
        if fig_fmt.get('subfigure_package'):
            spec['subfigure_package'] = fig_fmt['subfigure_package']
        figure_float_support = self._derive_double_column_float_support('figure')
        spec['allow_full_width'] = bool(figure_float_support['supports'])
        spec['full_width_source'] = figure_float_support['source']
        if figure_float_support['supports']:
            spec['full_width_container'] = self._full_width_container(
                figure_float_support, kind='figure')
        full_pos = (
            self._derive_float_position('figure', starred=True) or
            self._derive_float_position('figure', starred=False)
        )
        if full_pos:
            spec['full_width_float_position'] = full_pos
        return spec

    def _derive_caption_spec(self):
        cap_fmt = self.spec.get('caption_format', {})
        fig_fmt = self.spec.get('figure_format', {})
        tbl_fmt = self.spec.get('table_format', {})
        spec = {}
        if cap_fmt.get('separator'):
            spec['separator'] = cap_fmt['separator']
        if cap_fmt.get('label_weight'):
            spec['label_weight'] = cap_fmt['label_weight']
        # figure/table caption位置: 从caption_format或figure_format/table_format取
        fig_pos = cap_fmt.get('figure_position', '') or fig_fmt.get('caption_position', '')
        tbl_pos = cap_fmt.get('table_position', '') or tbl_fmt.get('caption_position', '')
        if fig_pos:
            spec['figure_position'] = fig_pos
        if tbl_pos:
            spec['table_position'] = tbl_pos
        if cap_fmt.get('font_size'):
            spec['font_size'] = cap_fmt['font_size']
        return spec

    def _derive_table_spec(self):
        tbl_fmt = self.spec.get('table_format', {})
        cap_fmt = self.spec.get('caption_format', {})
        spec = {}
        # caption位置: 从spec获取，无spec时默认above
        tbl_pos = cap_fmt.get('table_position', '') or tbl_fmt.get('caption_position', '')
        spec['caption_position'] = tbl_pos if tbl_pos else 'above'
        # 表格字体: 从spec获取，无spec时不设置(让cls控制)
        if tbl_fmt.get('header_size'):
            spec['header_size'] = tbl_fmt['header_size']
        else:
            # 无spec时header_size跟随body_size，或留空让cls控制
            spec['header_size'] = ''
        if tbl_fmt.get('body_size'):
            spec['body_size'] = tbl_fmt['body_size']
        else:
            spec['body_size'] = ''
        if tbl_fmt.get('rule_style'):
            spec['rule_style'] = tbl_fmt['rule_style']
        elif tbl_fmt.get('table_packages'):
            if 'booktabs' in tbl_fmt['table_packages']:
                spec['rule_style'] = 'booktabs'
        if tbl_fmt.get('hline_commands'):
            spec['hline_commands'] = tbl_fmt['hline_commands']
        if tbl_fmt.get('vertical_rules'):
            spec['vertical_rules'] = tbl_fmt['vertical_rules']
            spec['no_vertical_rules'] = tbl_fmt['vertical_rules'] == 'none'
        if tbl_fmt.get('float_position'):
            spec['float_position'] = tbl_fmt['float_position']
        else:
            spec['float_position'] = 'htbp'
        table_float_support = self._derive_double_column_float_support('table')
        if table_float_support['supports']:
            spec['full_width_container'] = self._full_width_container(
                table_float_support, kind='table')
            spec['full_width_source'] = table_float_support['source']
        full_pos = (
            self._derive_float_position('table', starred=True) or
            self._derive_float_position('table', starred=False)
        )
        if full_pos:
            spec['full_width_float_position'] = full_pos
        if tbl_fmt.get('alignment'):
            spec['alignment'] = tbl_fmt['alignment']
        else:
            spec['alignment'] = 'center'
        if self.spec.get('required_packages', {}).get('supertabular'):
            spec['multipage_environment'] = 'supertabular'
        return spec

    def _derive_bibliography_spec(self):
        bib_fmt = self.spec.get('bibliography_format', {})
        spec = {}
        if bib_fmt.get('style') == 'natbib':
            spec['style'] = 'author-year (natbib)'
            spec['natbib'] = True
        elif bib_fmt.get('style') == 'biblatex':
            spec['style'] = 'biblatex'
        if bib_fmt.get('bst_file'):
            spec['bst_file'] = bib_fmt['bst_file']
        if bib_fmt.get('natbib_options'):
            spec['natbib_options'] = bib_fmt['natbib_options']
        if bib_fmt.get('bibpunct'):
            spec['bibpunct'] = bib_fmt['bibpunct']
        if bib_fmt.get('font_size'):
            spec['font_size'] = bib_fmt['font_size']
        # bib文件名: 从skeleton_info获取，或从bib_path推导
        if bib_fmt.get('bib_filename'):
            spec['bib_filename'] = bib_fmt['bib_filename']
        return spec

    # 字体包→xelatex字体名映射（times/txfonts等在xelatex下不生效，需fontspec替代）
    _FONT_PKG_MAP = {
        'times':    {'main': 'Times New Roman',  'sans': 'Arial',       'mono': 'Courier New', 'math': 'XITS Math'},
        'txfonts':  {'main': 'Times New Roman',  'sans': 'Arial',       'mono': 'Courier New', 'math': 'XITS Math'},
        'ptm':      {'main': 'Times New Roman',  'sans': 'Arial',       'mono': 'Courier New', 'math': 'XITS Math'},
        'newtx':    {'main': 'Times New Roman',  'sans': 'Helvetica',   'mono': 'Courier New', 'math': 'XITS Math'},
        'mathptmx': {'main': 'Times New Roman',  'sans': 'Helvetica',   'mono': 'Courier New', 'math': 'XITS Math'},
        'mathtime': {'main': None, 'sans': None, 'mono': None, 'math': 'XITS Math'},
        'mtpro2':   {'main': None, 'sans': None, 'mono': None, 'math': 'XITS Math'},
        'newtxmath':{'main': None, 'sans': None, 'mono': None, 'math': 'XITS Math'},
    }

    def _derive_document_spec(self):
        """推导文档级spec: 栏数、字体需求、cls选项

        栏数检测: 从CLS内部代码提取，而非检查documentclass选项。
        很多模板（如Copernicus）的twocolumn是在CLS内部根据config_mode
        通过\\@twocolumntrue激活的，不在documentclass选项中。
        """
        dc = self.spec.get('document_class', {})
        default_opts = dc.get('default_options', [])
        base_opts = dc.get('base_class_options', '') or ''
        required_pkgs = self.spec.get('required_packages', {})

        # 从CLS内部检测当前config_mode对应的栏数
        is_twocolumn = self._detect_twocolumn_from_cls(
            base_opts, default_opts
        )

        needs_fontspec = False
        main_font = sans_font = mono_font = math_font = None
        # 优先从spec.fonts获取字体名（template-extract-lite从cls中提取）
        fonts_spec = self.spec.get('fonts', {})
        if fonts_spec.get('main_font'):
            main_font = fonts_spec['main_font']
            needs_fontspec = True
        if fonts_spec.get('sans_font'):
            sans_font = fonts_spec['sans_font']
            needs_fontspec = True
        if fonts_spec.get('mono_font'):
            mono_font = fonts_spec['mono_font']
            needs_fontspec = True
        if fonts_spec.get('math_font'):
            math_font = fonts_spec['math_font']
            needs_fontspec = True
        # 回退: 从FONT_PKG_MAP根据字体包名推导
        if not needs_fontspec:
            for pkg_name, font_map in self._FONT_PKG_MAP.items():
                if pkg_name in required_pkgs:
                    needs_fontspec = True
                    if font_map['main'] and not main_font:
                        main_font = font_map['main']
                    if font_map['sans'] and not sans_font:
                        sans_font = font_map['sans']
                    if font_map['mono'] and not mono_font:
                        mono_font = font_map['mono']
                    if font_map['math'] and not math_font:
                        math_font = font_map['math']

        table_float_support = self._derive_double_column_float_support('table')
        figure_float_support = self._derive_double_column_float_support('figure')
        any_float_support = table_float_support['supports'] or figure_float_support['supports']
        return {
            'is_twocolumn': is_twocolumn,
            'supports_double_column_floats': bool(any_float_support),
            'double_column_float_source': (
                table_float_support['source'] if table_float_support['supports']
                else figure_float_support['source']
            ),
            'supports_double_column_tables': bool(table_float_support['supports']),
            'double_column_table_source': table_float_support['source'],
            'supports_double_column_figures': bool(figure_float_support['supports']),
            'double_column_figure_source': figure_float_support['source'],
            'needs_fontspec': needs_fontspec,
            'has_manuscript_option': 'manuscript' in dc.get('declared_options', []),
            'main_font': main_font, 'sans_font': sans_font,
            'mono_font': mono_font, 'math_font': math_font,
        }

    def _detect_twocolumn_from_cls(self, base_opts, default_opts):
        """从CLS文件内部代码检测当前config_mode对应是否为双栏

        检测策略（优先级从高到低）:
        1. 使用template_config.py从CLS中提取column_count（最准确）
        2. 检查documentclass选项中是否显式声明twocolumn
        3. 默认单栏
        """
        # 方法1: 从CLS内部提取（处理Copernicus等通过\\@twocolumntrue切换的模板）
        if self.cls_path:
            try:
                from pathlib import Path
                import sys
                skill_dir = Path(__file__).resolve().parent.parent
                if str(skill_dir) not in sys.path:
                    sys.path.insert(0, str(skill_dir))
                from shared.template_config import (
                    get_page_geometry_for_mode, detect_effective_column_count)

                content = Path(self.cls_path).read_text(encoding='utf-8', errors='ignore')
                geo = get_page_geometry_for_mode(
                    content, config_mode=self._effective_config_mode())
                if geo and isinstance(geo, dict) and 'column_count' in geo:
                    return geo.get('column_count', 1) >= 2
                return detect_effective_column_count(
                    content, default_opts) >= 2
            except Exception:
                pass

        # 方法2: documentclass选项直接声明
        if 'twocolumn' in base_opts or 'twocolumn' in default_opts:
            return True

        return False

    def _derive_abstract_spec(self):
        """推导摘要spec: 关键词位置、字号"""
        abs_fmt = self.spec.get('abstract_format', {})
        return {
            'font_size': abs_fmt.get('font_size', ''),
            'keywords_inside': abs_fmt.get('keywords_inside_abstract', False),
        }

    def _derive_body_text_spec(self):
        """推导正文spec: 行间距、首行缩进、段间距"""
        body_fmt = self.spec.get('body_text', {})
        return {
            'line_spacing': body_fmt.get('line_spacing', ''),
            'first_line_indent': body_fmt.get('first_line_indent', ''),
            'paragraph_skip': body_fmt.get('paragraph_skip', ''),
        }

    def _derive_equation_spec(self):
        """推导公式spec: 字体字号"""
        eq_fmt = self.spec.get('equation_format', {})
        spec = {
            'font_size': eq_fmt.get('font_size', ''),
        }
        if eq_fmt.get('inline_identifier_style'):
            spec['inline_identifier_style'] = eq_fmt['inline_identifier_style']
        return spec
