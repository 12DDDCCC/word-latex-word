"""
精简版LaTeX模板提取器 — 核心提取逻辑
包含 TemplateExtractLite 类及所有 _extract_* 方法
"""
import re
import json
import sys
from collections import OrderedDict
from pathlib import Path

# shared模块导入路径
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.latex_parse_utils import BS, cmd as _cmd
from _skeleton_derive import _GuideMixin


def _looks_like_template_placeholder(value):
    text = str(value or '').strip().lower()
    return (
        not text
        or '...' in text
        or '<' in text
        or '>' in text
        or any(token in text for token in ('your ', 'filename', 'path/to', 'bibdatabase'))
    )


def _safe_graphics_extensions(value):
    text = str(value or '').strip()
    if _looks_like_template_placeholder(text):
        return ''
    parts = [part.strip() for part in text.strip('{}').split(',') if part.strip()]
    if not parts:
        return ''
    if any(re.fullmatch(r'\.[A-Za-z0-9]+', part) is None for part in parts):
        return ''
    return ','.join(parts)


def _safe_graphics_path(value):
    text = str(value or '').strip()
    if _looks_like_template_placeholder(text):
        return ''
    return text


def _safe_bib_resource(value):
    text = str(value or '').strip()
    if _looks_like_template_placeholder(text):
        return ''
    if re.search(r'[<>:"|?*]', text):
        return ''
    return text


class TemplateExtractLite(_GuideMixin):
    def __init__(self, cls_path):
        self.cls_path = Path(cls_path)
        self.content = self.cls_path.read_text(encoding='utf-8', errors='ignore')
        self.flat = re.sub(r'\n\s+', ' ', self.content)  # 压平换行
        self.base_size = 10
        self._detect_base_size()
        self.journal = self.cls_path.stem

    def _detect_base_size(self):
        m = re.search(BS + r'LoadClass\s*(?:\[([^\]]*)\])?\s*\{(?:article|report|book)\}', self.content)
        if m and m.group(1):
            for opt in m.group(1).split(','):
                opt = opt.strip()
                if opt in ('10pt', '11pt', '12pt'):
                    self.base_size = int(opt.replace('pt', ''))
                    break

    def extract_all(self):
        spec = OrderedDict()
        spec['document_class'] = self._extract_document_class()
        spec['required_packages'] = self._extract_packages()
        spec['title_format'] = self._extract_title_format()
        spec['author_format'] = self._extract_author_format()
        spec['abstract_format'] = self._extract_abstract_format()
        spec['keywords_format'] = self._extract_keywords_format()
        spec['section_command'] = self._extract_section_commands()
        spec['special_envs'] = self._extract_special_envs()
        spec['figure_format'] = self._extract_figure_format()
        spec['table_format'] = self._extract_table_format()
        spec['caption_format'] = self._extract_caption_format()
        spec['equation_format'] = self._extract_equation_format()
        spec['bibliography_format'] = self._extract_bibliography_format()
        spec['appendix_format'] = self._extract_appendix_format()
        spec['template_specific'] = self._extract_template_specific()
        spec['numbering'] = self._extract_numbering()
        return spec

    # ─── 1. 文档类 ─────────────────────────────────────
    def _extract_document_class(self):
        info = OrderedDict()
        # 类名
        m = re.search(BS + r'ProvidesClass\{(\w+)\}', self.content)
        if m:
            info['class_name'] = m.group(1)
        # 选项
        m = re.search(BS + r'LoadClass\[([^\]]+)\]', self.content)
        if m:
            info['base_class_options'] = m.group(1).strip()
        m = re.search(BS + r'LoadClass(?:\[[^\]]*\])?\{(\w+)\}', self.content)
        if m:
            info['base_class'] = m.group(1)
        # 类选项定义
        options = []
        for m in re.finditer(BS + r'DeclareOption\{(\w+)\}', self.content):
            options.append(m.group(1))
        for m in re.finditer(BS + r'DeclareOption\*\{(.{0,80})', self.content):
            options.append('*')
        if options:
            info['declared_options'] = options
        # 默认选项
        m = re.search(BS + r'ExecuteOptions\{([^}]+)\}', self.content)
        if m:
            info['default_options'] = [o.strip() for o in m.group(1).split(',')]
        return info

    # ─── 2. 必需包 ─────────────────────────────────────
    def _extract_packages(self):
        packages = OrderedDict()
        for m in re.finditer(BS + r'(?:RequirePackage|usepackage)(?:\s*\[([^\]]*)\])?\s*\{([^}]+)\}', self.content):
            opts = m.group(1) or ''
            for pkg in m.group(2).split(','):
                pkg = pkg.strip()
                if pkg and pkg not in packages:
                    packages[pkg] = opts if opts else True
        return packages

    # ─── 3. 标题格式 ────────────────────────────────────
    def _extract_title_format(self):
        info = OrderedDict()
        # \title 命令的参数结构
        m = re.search(BS + r'(?:newcommand|renewcommand|DeclareRobustCommand)\{?' + BS + r'title\}?\s*(\[\d\])?\s*\{(.{0,500})', self.content, re.DOTALL)
        if m:
            nargs = m.group(1)
            info['title_args'] = int(nargs.strip('[]')) if nargs else 0
            body = m.group(2)
            if BS + r'shorttitle' in body or '[#1]' in body or '[##1]' in body:
                info['has_short_title'] = True
        # 搜索\maketitle中标题的排版
        mt = re.search(BS + r'(?:long' + BS + r')?def' + BS + r'@?maketitle\b(.{0,4000})', self.content, re.DOTALL)
        if mt:
            body = mt.group(1)
            if BS + r'LARGE' in body or BS + r'LARGE' in body:
                info['title_size'] = 'LARGE'
            elif BS + r'Large' in body:
                info['title_size'] = 'Large'
            elif BS + r'large' in body:
                info['title_size'] = 'large'
            if 'bfseries' in body or BS + r'bf' in body:
                info['title_weight'] = 'bold'
            if BS + r'centering' in body or BS + r'center' in body:
                info['title_alignment'] = 'center'
            if BS + r'sffamily' in body:
                info['title_font_family'] = 'sans-serif'
        return info

    # ─── 4. 作者格式 ────────────────────────────────────
    def _extract_author_format(self):
        info = OrderedDict()
        # \author 命令参数
        for cmd in ('author', 'affil', 'affiliation', 'correspondence', 'corresponding_author',
                     'thanks', 'email', 'orcid', 'presentaddress'):
            m = re.search(BS + r'(?:newcommand|renewcommand|DeclareRobustCommand)\{?' + BS + cmd + r'\}?\s*(\[\d\])', self.content)
            if m:
                info[f'{cmd}_args'] = int(m.group(1).strip('[]'))
            # \def形式
            m = re.search(BS + r'def\s*' + BS + cmd + r'\s*#(\d)', self.content)
            if m:
                info[f'{cmd}_args'] = int(m.group(1))
        # 从\maketitle中看作者排版
        mt = re.search(BS + r'(?:long' + BS + r')?def' + BS + r'@?maketitle\b(.{0,4000})', self.content, re.DOTALL)
        if mt:
            body = mt.group(1)
            if BS + r'and' in body or BS + r'authormark' in body:
                info['author_separator'] = r'\and'
        return info

    # ─── 5. 摘要格式 ────────────────────────────────────
    def _extract_abstract_format(self):
        info = OrderedDict()
        # \begin{abstract} 还是 \abstract{} 命令
        if re.search(BS + r'newenvironment\{abstract\}', self.content):
            info['type'] = 'environment'  # \begin{abstract}
        elif re.search(BS + r'(?:newcommand|def)\{?' + BS + r'abstract\}', self.content):
            info['type'] = 'command'  # \abstract{...}
        else:
            info['type'] = 'environment'  # 默认环境
        # 摘要宽度限制
        m = re.search(BS + r'begin\{abstract\}.*?' + BS + r'small', self.content, re.DOTALL)
        if m:
            info['font_size'] = 'small'
        # 摘要关键词是否在abstract环境内
        if re.search(BS + r'begin\{abstract\}.*?' + BS + r'keywords', self.content, re.DOTALL):
            info['keywords_inside_abstract'] = True
        return info

    # ─── 6. 关键词格式 ──────────────────────────────────
    def _extract_keywords_format(self):
        info = OrderedDict()
        # 多种定义形式
        patterns = [
            (BS + r'(?:newcommand|renewcommand)\{?' + BS + r'keywords\}?\s*(\[\d\])?\s*\{(.{0,200})', 'cmd'),
            (BS + r'def\s*' + BS + r'keywords\s*#\d\s*\{(.{0,200})', 'def'),
            (BS + r'DeclareRobustCommand\s*' + BS + r'keywords\s*(\[\d\])?\s*\{(.{0,200})', 'robust'),
        ]
        for pat, ptype in patterns:
            m = re.search(pat, self.content, re.DOTALL)
            if m:
                if ptype == 'def':
                    info['type'] = 'command'
                    body = m.group(1)
                else:
                    nargs = m.group(1) if len(m.groups()) >= 2 else None
                    info['args'] = int(nargs.strip('[]')) if nargs else 0
                    body = m.group(2) if len(m.groups()) >= 2 else m.group(1)
                    info['type'] = 'command'
                if 'bfseries' in body or BS + r'textbf' in body:
                    info['label_weight'] = 'bold'
                lbl = re.search(r'\{([^}]*[Kk]eyword[^}]*)\}', body)
                if lbl:
                    info['label_text'] = lbl.group(1)
                break
        # 检查关键词是否存在但未匹配到格式
        if not info and re.search(BS + r'keywords\b', self.content):
            info['exists'] = True
        return info

    # ─── 7. 章节命令 ────────────────────────────────────
    def _extract_section_commands(self):
        info = OrderedDict()
        standard = ['section', 'subsection', 'subsubsection', 'paragraph', 'subparagraph']
        for cmd in standard:
            # 检查是否有\let别名
            let_m = re.search(BS + r'let\s*' + BS + cmd + r'\s*=?' + BS + r'(\w+)', self.flat)
            if let_m:
                info[cmd] = {'alias_of': let_m.group(1)}
                continue
            # 检查是否有重定义
            if re.search(BS + r'(?:renewcommand|def)\s*' + BS + cmd + r'\b', self.flat):
                info[cmd] = {'custom': True}
            # 检查\@startsection定义
            for sm in re.finditer(_cmd('@startsection') + r'\{' + cmd + r'\}', self.flat):
                info[cmd] = {'startsection': True}
                break
        return info

    # ─── 8. 特殊声明环境 ────────────────────────────────
    def _extract_special_envs(self):
        envs = OrderedDict()
        env_names = [
            'acknowledgements', 'acknowledgment', 'dataavailability',
            'codeavailability', 'authorcontribution', 'competinginterests',
            'supplement', 'supplementary', 'correspondence',
            'introduction', 'conclusions', 'methods', 'results', 'discussion',
        ]
        for env_name in env_names:
            # \newenvironment{envname}
            m = re.search(BS + r'newenvironment\{' + env_name + r'\}', self.content)
            if m:
                envs[env_name] = {'type': 'environment'}
                # 检查映射到什么section级别
                body_m = re.search(BS + r'newenvironment\{' + env_name + r'\}\s*(?:\[\d?\])?\s*\{(.{0,300})', self.content, re.DOTALL)
                if body_m:
                    body = body_m.group(1)
                    if r'\section' in body and r'\subsection' not in body:
                        envs[env_name]['maps_to'] = 'section'
                    elif r'\subsection' in body:
                        envs[env_name]['maps_to'] = 'subsection'
                continue
            # \def\envname{\section{...}}
            m = re.search(BS + r'def\s*' + BS + env_name + r'\b\s*\{(.{0,300})', self.content, re.DOTALL)
            if m:
                body = m.group(1)
                e = OrderedDict()
                e['type'] = 'command'
                if r'\section' in body:
                    e['maps_to'] = 'section'
                    sec_title = re.search(r'\\section\*?\{([^}]*)\}', body)
                    if sec_title:
                        e['section_title'] = sec_title.group(1)
                elif r'\subsection' in body:
                    e['maps_to'] = 'subsection'
                envs[env_name] = e
                continue
            # \newcommand{\envname} 或 \newcommand\envname（含可选参数如[1][default]）
            m = re.search(BS + r'newcommand\{?' + BS + env_name + r'\}?\s*(?:\[\d\])?\s*(?:\[[^\]]*\])?\s*\{(.{0,300})', self.content, re.DOTALL)
            if m:
                body = m.group(1)
                e = OrderedDict()
                e['type'] = 'command'
                if r'\section' in body:
                    e['maps_to'] = 'section'
                elif r'\subsection' in body:
                    e['maps_to'] = 'subsection'
                envs[env_name] = e
                continue
            # \generateCommand{envname} (Copernicus风格)
            m = re.search(BS + r'generateCommand\{' + env_name + r'\}', self.content)
            if m:
                envs[env_name] = {'type': 'command', 'generated': True}
                continue
        return envs

    # ─── 9. 图格式 ─────────────────────────────────────
    def _extract_figure_format(self):
        info = OrderedDict()
        # 默认位置参数
        m = re.search(BS + r'newenvironment\{figure\*?\}[^\{]*\{(.{0,200})', self.content, re.DOTALL)
        if m:
            body = m.group(1)
            loc = re.search(r'\[(h|t|b|p|htbp|H)\]', body)
            if loc:
                info['default_position'] = loc.group(1)
        # 子图包
        for pkg in ('subfigure', 'subfig', 'subcaption'):
            if pkg in self.content:
                info['subfigure_package'] = pkg
                break
        # 图片扩展名
        m = re.search(BS + r'DeclareGraphicsExtensions\{([^}]+)\}', self.content)
        if m:
            graphics_extensions = _safe_graphics_extensions(m.group(1))
            if graphics_extensions:
                info['graphics_extensions'] = graphics_extensions
        # 图片路径
        m = re.search(BS + r'graphicspath\{([^}]+)\}', self.content)
        if m:
            graphics_path = _safe_graphics_path(m.group(1))
            if graphics_path:
                info['graphics_path'] = graphics_path
        # 图片编号格式
        m = re.search(BS + r'def\s*' + BS + r'thefigure\s*\{([^}]+)\}', self.content)
        if m:
            info['numbering_format'] = m.group(1)
        return info

    # ─── 10. 表格式 ─────────────────────────────────────
    def _extract_table_format(self):
        info = OrderedDict()
        # 表格相关包
        table_pkgs = ['booktabs', 'tabularx', 'longtable', 'multirow', 'array', 'siunitx']
        found = [p for p in table_pkgs if p in self.content]
        if found:
            info['table_packages'] = found
        hline_cmds = [
            cmd for cmd in (
                'tophline', 'middlehline', 'bottomhline',
                'toprule', 'midrule', 'bottomrule',
                'colrule', 'botrule',
                'hline', 'cline', 'cmidrule'
            )
            if re.search(BS + cmd + r'\b', self.content)
        ]
        if hline_cmds:
            info['hline_commands'] = hline_cmds
        if {'tophline', 'middlehline', 'bottomhline'}.issubset(set(hline_cmds)) \
                or {'toprule', 'colrule', 'botrule'}.issubset(set(hline_cmds)):
            info['rule_style'] = 'template_hlines'
        elif 'booktabs' in found or {'toprule', 'midrule', 'bottomrule'}.issubset(set(hline_cmds)):
            info['rule_style'] = 'booktabs'
        tabular_specs = re.findall(
            BS + r'begin\{(?:tabular|tabularx|longtable)\}\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}',
            self.content
        )
        if tabular_specs:
            info['vertical_rules'] = 'source' if any('|' in spec for spec in tabular_specs) else 'none'
        if 'vertical_rules' not in info and info.get('rule_style') in ('template_hlines', 'booktabs'):
            info['vertical_rules'] = 'none'
        # 表格编号格式
        m = re.search(BS + r'def\s*' + BS + r'thetable\s*\{([^}]+)\}', self.content)
        if m:
            info['numbering_format'] = m.group(1)
        return info

    # ─── 11. Caption格式 ─────────────────────────────────
    def _extract_caption_format(self):
        info = OrderedDict()
        cap_patterns = [
            BS + r'long' + BS + r'def' + BS + r'@makecaption\s*#\d\s*#\d\s*\{(.{0,1200})\}',
            BS + r'long' + BS + r'def' + BS + r'@makecaption\{[^}]*\}\s*\{(.{0,1200})\}',
            BS + r'def' + BS + r'@makecaption\{[^}]*\}\s*\{(.{0,1200})\}',
            BS + r'renewcommand\s*' + BS + r'@makecaption\s*\[\d\]\s*\{(.{0,1200})\}',
        ]
        body = None
        for pat in cap_patterns:
            m = re.search(pat, self.content, re.DOTALL)
            if m:
                body = m.group(1)
                break
        if body:
            if ': ' in body or BS + r'colon' in body:
                info['separator'] = ':'
            elif '. ' in body:
                info['separator'] = '.'
            if 'bfseries' in body:
                info['label_weight'] = 'bold'
            if BS + r'@captype' in body:
                if re.search(BS + r'def' + BS + r'@tempa\{figure\}', body) and re.search(BS + r'def' + BS + r'@tempa\{table\}', body):
                    info['figure_position'] = 'below'
                    info['table_position'] = 'above'
        if not body:
            m = re.search(BS + r'usepackage(?:\[([^\]]*)\])?\{caption\}', self.content)
            if m:
                info['caption_package'] = True
                if m.group(1):
                    info['caption_options'] = m.group(1)
        return info

    # ─── 12. 公式格式 ────────────────────────────────────
    def _extract_equation_format(self):
        info = OrderedDict()
        # amsmath选项
        m = re.search(BS + r'(?:RequirePackage|usepackage)\[([^\]]+)\]\{amsmath\}', self.content)
        if m:
            info['amsmath_options'] = m.group(1)
        # 公式编号格式
        m = re.search(BS + r'def\s*' + BS + r'theequation\s*\{([^}]+)\}', self.content)
        if m:
            info['numbering_format'] = m.group(1)
        # \numberwithin
        m = re.search(BS + r'numberwithin\{equation\}\{(\w+)\}', self.content)
        if m:
            info['number_within'] = m.group(1)
        inline_style = self._extract_inline_identifier_style()
        if inline_style:
            info['inline_identifier_style'] = inline_style
        return info

    def _extract_inline_identifier_style(self):
        """Detect explicit template rules for inline identifier-like formulas."""
        normalized = self.flat.lower()
        identifier_macros = (
            'inlineidentifier', 'inline_identifier', 'identifier',
            'varname', 'variable', 'code', 'datasetname'
        )
        for macro in identifier_macros:
            pat = BS + r'(?:newcommand|renewcommand|DeclareRobustCommand)\{?' + BS + macro + r'\}?\s*(?:\[\d\])?\s*\{(.{0,300})'
            m = re.search(pat, self.content, re.DOTALL)
            if not m:
                continue
            body = m.group(1)
            if re.search(BS + r'(?:mathit|emph|itshape)\b', body):
                return 'math_italic'
            if re.search(BS + r'(?:mathrm|textrm|textnormal|textup|upshape)\b', body):
                return 'body_upright'
        if re.search(r'inline\s+(?:identifier|variable|math).{0,80}(?:italic|mathit|itshape)', normalized):
            return 'math_italic'
        if re.search(r'inline\s+(?:identifier|variable|math).{0,80}(?:upright|roman|textnormal|mathrm)', normalized):
            return 'body_upright'
        return ''

    # ─── 13. 参考文献格式 ────────────────────────────────
    def _extract_bibliography_format(self):
        info = OrderedDict()
        # natbib
        m = re.search(BS + r'(?:RequirePackage|usepackage)\[([^\]]*)\]\{natbib\}', self.content)
        if m:
            info['style'] = 'natbib'
            info['natbib_options'] = m.group(1) if m.group(1) else ''
        # biblatex
        m = re.search(BS + r'usepackage\[([^\]]*)\]\{biblatex\}', self.content)
        if m:
            info['style'] = 'biblatex'
            info['biblatex_options'] = m.group(1)
            bm = re.search(BS + r'addbibresource\{([^}]+)\}', self.content)
            if bm:
                bib_resource = _safe_bib_resource(bm.group(1))
                if bib_resource:
                    info['bib_resource'] = bib_resource
        # bst
        m = re.search(BS + r'def\s*' + BS + r'bibliographystyle\{([^}]+)\}', self.content)
        if m:
            info['bst_file'] = m.group(1)
        # \bibliography命令
        m = re.search(BS + r'def\s*' + BS + r'bibliography(?:name)?\{([^}]+)\}', self.content)
        if m:
            info['default_bib_name'] = m.group(1)
        # 引用格式
        m = re.search(BS + r'bibpunct\{([^}]+)\}\{([^}]+)\}\{([^}]+)\}\{([^}]+)\}\{([^}]+)\}\{([^}]+)\}', self.content)
        if m:
            info['bibpunct'] = {
                'open': m.group(1), 'close': m.group(2),
                'style': m.group(3), 'sep': m.group(4),
                'between': m.group(5), 'after': m.group(6),
            }
        # 参考文献字号：兼容 \let\bibfont=\footnotesize / \let\bibfont\small / \renewcommand{\bibfont}{\small}
        m = re.search(
            BS + r'(?:global\s*)?(?:let|def|renewcommand)\{?' + BS + r'?bibfont\}?\s*=?\s*\{?' + BS + r'(tiny|scriptsize|footnotesize|small|normalsize|large)\b',
            self.content,
        )
        if m:
            info['font_size'] = m.group(1)
        return info

    # ─── 14. 附录格式 ────────────────────────────────────
    def _extract_appendix_format(self):
        info = OrderedDict()
        # appendix环境 vs \appendix命令
        if re.search(BS + r'newenvironment\{appendix\}', self.content):
            info['type'] = 'environment'
        elif re.search(BS + r'appendix', self.content):
            info['type'] = 'command'
        # 附录编号
        m = re.search(BS + r'def\s*' + BS + r'the' + BS + r'(?:appendix|section)\s*\{[^}]*' + BS + r'Alph', self.content)
        if m:
            info['numbering'] = 'Alphabetic'
        return info

    # ─── 15a. 编号系统 ─────────────────────────────────────
    def _extract_numbering(self):
        r"""统一提取所有编号定义和条件重置

        copernicus.cls 等模板中 \@addtoreset 可能位于条件块内
        （如 \\if@stage@final），需要区分无条件/条件重置。
        策略：先识别条件块范围，条件块内的 \@addtoreset 标记为 conditional，
        其余为 unconditional。
        """
        info = OrderedDict()

        # 1. 找出所有条件块的范围（\if@xxx ... \fi）
        cond_ranges = []  # list of (start, end, cond_name)
        for m in re.finditer(r'\\(if@[\w@]+)', self.content):
            cond_name = m.group(1)
            start = m.start()
            # 简化：找对应的 \fi（可能嵌套，但 copernicus.cls 中条件块一般不嵌套）
            fi_pos = self.content.find('\\fi', m.end())
            if fi_pos >= 0:
                cond_ranges.append((start, fi_pos + 2, cond_name))

        # 2. 遍历所有 \@addtoreset，判断是否在条件块内
        unconditional_resets = []
        conditional_resets = []
        for m in re.finditer(BS + r'@addtoreset\{(\w+)\}\{(\w+)\}', self.content):
            counter, parent = m.group(1), m.group(2)
            pos = m.start()
            in_cond = None
            for cs, ce, cn in cond_ranges:
                if cs <= pos <= ce:
                    in_cond = cn
                    break
            if in_cond:
                conditional_resets.append((counter, parent, in_cond))
            else:
                unconditional_resets.append((counter, parent))

        if unconditional_resets:
            info['unconditional_resets'] = unconditional_resets
        if conditional_resets:
            info['conditional_resets'] = conditional_resets

        # \numberwithin{counter}{parent}
        numberwithin = []
        for m in re.finditer(BS + r'numberwithin\{(\w+)\}\{(\w+)\}', self.content):
            numberwithin.append((m.group(1), m.group(2)))
        if numberwithin:
            info['numberwithin'] = numberwithin

        return info

    # ─── 15. 模板特有命令 ────────────────────────────────
    def _extract_template_specific(self):
        info = OrderedDict()
        # 常见模板特有命令
        special_cmds = [
            'correspondence', 'corresponding_author', 'runningauthor',
            'runningtitle', 'shorttitle', 'authormark',
            'received', 'accepted', 'published', 'revised',
            'disclosure', 'fundinginfo', 'funding',
            'supplementary', 'supplement', 'coverletter',
            'editor', 'reviewer', 'dedication',
        ]
        for cmd in special_cmds:
            # 检查命令是否存在
            if re.search(BS + r'(?:newcommand|renewcommand|def|DeclareRobustCommand)\{?' + BS + cmd + r'\}?\b', self.content):
                # 获取参数数量
                m = re.search(BS + r'(?:newcommand|renewcommand)\{?' + BS + cmd + r'\}?\s*(\[\d\])', self.content)
                nargs = int(m.group(1).strip('[]')) if m else 0
                info[cmd] = {'exists': True, 'args': nargs}
        # 声明段落列表（模板要求的必填段落）
        declarations = []
        for env_name in ('acknowledgements', 'acknowledgment', 'dataavailability',
                          'codeavailability', 'authorcontribution', 'competinginterests',
                          'fundinginfo', 'supplementary'):
            if re.search(BS + r'(?:newenvironment|newcommand|def)\{?' + BS + env_name, self.content):
                declarations.append(env_name)
        if declarations:
            info['required_declarations'] = declarations
        return info

    def to_json(self, output_path):
        spec = self.extract_all()
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(spec, f, indent=2, ensure_ascii=False)
        return spec
