"""LaTeX模板配置提取与选择模块

从 .cls 文件中动态提取所有配置选项（classic/manuscript/final/discussions等），
提取每个选项对应的页面参数。默认选择经典版（manuscript），非交互模式。

策略：逐行扫描 .cls 文件，追踪 \if/@else/\fi 嵌套状态，
按 active 条件分支收集各模式的页面参数。不使用正则嵌套追踪。
"""
import re
import json
import sys
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional


# ─── 数据结构 ──────────────────────────────────────────────────

@dataclass
class PageGeometry:
    """页面几何参数（单位: mm）"""
    paperwidth_mm: float = 210.0
    paperheight_mm: float = 277.0
    textwidth_mm: float = 177.0
    textheight_mm: Optional[float] = None
    oddsidemargin_mm: float = 16.4
    evensidemargin_mm: Optional[float] = None
    topmargin_mm: float = 0.0
    column_count: int = 1
    column_sep_mm: Optional[float] = None

    def to_dict(self):
        return asdict(self)


@dataclass
class ConfigOption:
    """模板配置选项"""
    name: str
    label: str
    is_default: bool = False
    page_geometry: Optional[PageGeometry] = None
    doc_options: list = field(default_factory=list)

    def to_dict(self):
        d = {'name': self.name, 'label': self.label, 'is_default': self.is_default,
             'doc_options': self.doc_options}
        if self.page_geometry:
            d['page_geometry'] = self.page_geometry.to_dict()
        return d


@dataclass
class TemplateConfig:
    """模板配置（解析结果+用户选择）"""
    cls_path: str
    cls_name: str
    options: list = field(default_factory=list)
    selected: Optional[ConfigOption] = None
    bleed_mm: float = 0.0

    def to_dict(self):
        return {
            'cls_path': self.cls_path,
            'cls_name': self.cls_name,
            'options': [o.to_dict() for o in self.options],
            'selected': self.selected.to_dict() if self.selected else None,
            'bleed_mm': self.bleed_mm,
        }


# ─── 单位转换 ──────────────────────────────────────────────────

_UNIT_TO_MM = {'mm': 1.0, 'cm': 10.0, 'pt': 0.3528, 'in': 25.4, 'bp': 0.3528}


def _to_mm(value: float, unit: str) -> float:
    return round(value * _UNIT_TO_MM.get(unit, 1.0), 2)


def detect_effective_column_count(cls_content: str, doc_options=None) -> int:
    """Detect the columns used for article body content by the class."""
    if re.search(r'\\twocolumn\s*\[\s*\\@maketitle\s*\]', cls_content or ''):
        return 2
    options = [str(item).lower() for item in (doc_options or [])]
    if 'twocolumn' in options:
        return 2
    if 'onecolumn' in options:
        return 1
    scanned = _scan_cls_for_page_params(cls_content or '')
    if 'manuscript' in options or 'classic' in options:
        mode = 'manuscript'
    elif 'final' in options or re.search(r'\\newif\\if@stage@final\s+\\@stage@finaltrue', cls_content or ''):
        mode = 'final_pub'
    else:
        mode = 'manuscript'
    branch_columns = scanned.get(mode, {}).get('column_count')
    if branch_columns:
        return int(branch_columns)
    execute_options = re.findall(r'\\ExecuteOptions\{([^}]+)\}', cls_content or '')
    flattened = [
        option.strip().lower()
        for group in execute_options
        for option in group.split(',')
    ]
    return 2 if 'twocolumn' in flattened else 1


# ─── 逐行条件追踪器 ────────────────────────────────────────────

class _ConditionTracker:
    """逐行追踪 \\if/@else/\\fi 嵌套，判断当前行属于哪个条件分支"""

    def __init__(self):
        self._stack = []  # list of (flag_name, in_true_branch)

    def process_line(self, line: str):
        """处理一行，更新条件状态。返回当前行是否应被处理"""
        stripped = line.strip()

        # \if@xxx → 推入条件（\w+ 不含@，需要用 [\w@]+ 匹配 if@stage@final 等）
        m = re.match(r'\\(if@[\w@]+)', stripped)
        if m:
            flag = m.group(1)
            self._stack.append((flag, True))  # 进入 true 分支
            return False  # \if 行本身不包含赋值

        # \else → 翻转当前分支
        if stripped.startswith('\\else'):
            if self._stack:
                flag, _ = self._stack[-1]
                self._stack[-1] = (flag, False)  # 切到 false 分支
            return False

        # \fi → 弹出条件
        if stripped == '\\fi' or stripped.startswith('\\fi ') or stripped.startswith('\\fi%'):
            if self._stack:
                self._stack.pop()
            return False

        return True  # 非条件控制行

    @property
    def active_conditions(self):
        """当前激活的条件列表 [(flag, in_true_branch), ...]"""
        return list(self._stack)

    def is_in_stage_final(self):
        """当前是否在 \\if@stage@final true 分支"""
        for flag, in_true in self._stack:
            if flag == 'if@stage@final' and in_true:
                return True
        return False

    def is_in_stage_not_final(self):
        """当前是否在 \\if@stage@final false 分支"""
        for flag, in_true in self._stack:
            if flag == 'if@stage@final' and not in_true:
                return True
        return False

    def is_in_manuscript(self):
        """当前是否在 \\if@manuscript true 分支"""
        for flag, in_true in self._stack:
            if flag == 'if@manuscript' and in_true:
                return True
        return False

    def is_in_not_manuscript(self):
        """当前是否在 \\if@manuscript false 分支"""
        for flag, in_true in self._stack:
            if flag == 'if@manuscript' and not in_true:
                return True
        return False


# ─── 逐行扫描 .cls ────────────────────────────────────────────

def _scan_cls_for_page_params(cls_content: str) -> dict:
    """逐行扫描 .cls，按条件分支收集页面参数

    Returns:
        dict: {
            'bleed_mm': float,
            'manuscript': {param: value},
            'final_pub': {param: value},
            'discussions': {param: value},
        }
    """
    bleed_mm = 3.0
    ms_params = {}      # manuscript (final+manuscript)
    pub_params = {}     # 出版版 (final+非manuscript)
    disc_params = {}    # discussions (非final)
    common_final = {}   # final 内但在 \if@manuscript 外的参数

    # 先提取 bleed
    m = re.search(r'\\bleed\s*([\d.]+)\s*(mm|cm|pt|in)\s*\\relax', cls_content)
    if m:
        bleed_mm = _to_mm(float(m.group(1)), m.group(2))

    tracker = _ConditionTracker()

    for line in cls_content.split('\n'):
        should_process = tracker.process_line(line)
        if not should_process:
            continue

        stripped = line.strip()

        # ─── dimexpr 赋值: \textheight\dimexpr660\p@-37mm+11.4mm\relax ───
        # 支持多步运算: base +/- val1 +/- val2 ...
        m = re.match(
            r'\\(paperwidth|paperheight|textwidth|textheight|oddsidemargin|evensidemargin|topmargin|headheight|headsep|footskip)\s*\\dimexpr\s*(.*?)\\relax',
            stripped
        )
        if m:
            cmd = m.group(1)
            expr = m.group(2).strip()
            # 解析 dimexpr: 依次处理 value+unit 和 sign
            val = 0.0
            sign = 1.0
            for token in re.split(r'([+-])', expr):
                token = token.strip()
                if token == '+':
                    sign = 1.0
                    continue
                elif token == '-':
                    sign = -1.0
                    continue
                elif not token:
                    continue
                # 数值+单位 或 \xxx\p@
                num_m = re.match(r'([\d.]+)\s*(mm|cm|pt|in|\\p@)', token)
                if num_m:
                    unit = num_m.group(2)
                    if unit == '\\p@':
                        unit = 'pt'
                    val += sign * _to_mm(float(num_m.group(1)), unit)
                    sign = 1.0  # reset
                    continue
                bleed_m = re.match(r'([\d.]+)?\s*\\bleed', token)
                if bleed_m:
                    factor = float(bleed_m.group(1) or 1.0)
                    val += sign * factor * bleed_mm
                    sign = 1.0
                    continue
                # \660\p@ 格式 (数字和\p@可能无空格)
                num_m2 = re.match(r'(\d+)\s*\\p@', token)
                if num_m2:
                    val += sign * _to_mm(float(num_m2.group(1)), 'pt')
                    sign = 1.0
                    continue
                # \baselineskip 等变量 → 忽略(无法静态计算)
                sign = 1.0
            val = round(val, 2)

            if tracker.is_in_stage_final():
                if tracker.is_in_manuscript():
                    ms_params[cmd] = val
                elif tracker.is_in_not_manuscript():
                    pub_params[cmd] = val
                else:
                    # 在 final 内但不在 \if@manuscript 内 → 共享
                    common_final[cmd] = val
            elif tracker.is_in_stage_not_final():
                disc_params[cmd] = val
            continue

        # ─── 简单赋值: \textwidth177mm, \oddsidemargin-15.4mm, \paperheight159mm ───
        #    注意: paperwidth/paperheight 也可能有简单赋值（不含\dimexpr）
        m = re.match(
            r'\\(paperwidth|paperheight|textwidth|oddsidemargin|evensidemargin|topmargin|textheight|headheight|headsep|footskip|columnsep)\s*(-?[\d.]+)\s*(mm|cm|pt|in)',
            stripped
        )
        if m:
            cmd = m.group(1)
            val = _to_mm(float(m.group(2)), m.group(3))

            if tracker.is_in_stage_final():
                if tracker.is_in_manuscript():
                    ms_params[cmd] = val
                elif tracker.is_in_not_manuscript():
                    pub_params[cmd] = val
                else:
                    common_final[cmd] = val
            elif tracker.is_in_stage_not_final():
                disc_params[cmd] = val
            continue

        # ─── \z@ 赋值 (=0pt): \topmargin\z@ ───
        m = re.match(
            r'\\(textwidth|oddsidemargin|evensidemargin|topmargin|textheight|headheight|headsep|footskip)\s*\\z@',
            stripped
        )
        if m:
            cmd = m.group(1)
            if tracker.is_in_stage_final():
                if tracker.is_in_manuscript():
                    ms_params[cmd] = 0.0
                elif tracker.is_in_not_manuscript():
                    pub_params[cmd] = 0.0
                else:
                    common_final[cmd] = 0.0
            elif tracker.is_in_stage_not_final():
                disc_params[cmd] = 0.0
            continue

        # ─── 双栏检测 ───
        if '\\@twocolumntrue' in stripped:
            if tracker.is_in_stage_final():
                if tracker.is_in_manuscript():
                    ms_params['column_count'] = 1
                elif tracker.is_in_not_manuscript():
                    pub_params['column_count'] = 2
                else:
                    common_final['column_count'] = 2
            elif tracker.is_in_stage_not_final():
                disc_params['column_count'] = 1

    # 将 common_final 参数合并到 ms_params 和 pub_params
    for k, v in common_final.items():
        if k not in ms_params:
            ms_params[k] = v
        if k not in pub_params:
            pub_params[k] = v

    return {
        'bleed_mm': bleed_mm,
        'manuscript': ms_params,
        'final_pub': pub_params,
        'discussions': disc_params,
    }


# ─── 构建 ConfigOption ────────────────────────────────────────

def _build_page_geometry(params: dict, bleed_mm: float,
                         default_pw: float, default_ph: float,
                         default_tw: float, default_om: float,
                         default_tm: float, columns: int) -> Optional[PageGeometry]:
    """从扫描参数构建 PageGeometry"""
    if not params:
        return None

    pw = params.get('paperwidth')
    ph = params.get('paperheight')

    # 从含bleed的值减去2*bleed得到实际页面尺寸
    if pw and pw > default_pw:
        pw = round(pw - 2 * bleed_mm, 2)
    if ph and ph > default_ph + 10:
        ph = round(ph - 2 * bleed_mm, 2)

    return PageGeometry(
        paperwidth_mm=pw or default_pw,
        paperheight_mm=ph or default_ph,
        textwidth_mm=params.get('textwidth', default_tw),
        textheight_mm=params.get('textheight'),
        oddsidemargin_mm=params.get('oddsidemargin', default_om),
        evensidemargin_mm=params.get('evensidemargin'),
        topmargin_mm=params.get('topmargin', default_tm),
        column_count=params.get('column_count', columns),
        column_sep_mm=params.get('columnsep'),
    )


# ─── 核心解析 ──────────────────────────────────────────────────

def parse_cls_options(cls_path: str) -> TemplateConfig:
    """解析 .cls 文件，提取所有可用配置选项及其页面参数"""
    cls_path = str(cls_path)
    content = Path(cls_path).read_text(encoding='utf-8', errors='ignore')

    cls_name_match = re.search(r'\\ProvidesClass\{(\w+)\}', content)
    cls_name = cls_name_match.group(1) if cls_name_match else Path(cls_path).stem

    result = _scan_cls_for_page_params(content)
    bleed_mm = result['bleed_mm']

    # 1. manuscript
    ms_geo = _build_page_geometry(
        result['manuscript'], bleed_mm,
        default_pw=210.0, default_ph=240.0,
        default_tw=177.0, default_om=16.4, default_tm=10.0, columns=1
    )
    ms_opt = ConfigOption(
        name='manuscript',
        label='Manuscript (经典版, 单栏+行号)',
        is_default=True,
        page_geometry=ms_geo,
        doc_options=['manuscript'],
    )

    # 2. final 出版版
    pub_geo = _build_page_geometry(
        result['final_pub'], bleed_mm,
        default_pw=210.0, default_ph=277.0,
        default_tw=177.0, default_om=16.4, default_tm=0.0, columns=2
    )
    pub_opt = ConfigOption(
        name='final',
        label='Final (最终出版版, 双栏)',
        is_default=False,
        page_geometry=pub_geo,
        doc_options=[],
    )

    options = [ms_opt, pub_opt]

    # 3. discussions
    if result['discussions']:
        disc_geo = _build_page_geometry(
            result['discussions'], bleed_mm,
            default_pw=166.0, default_ph=159.0,
            default_tw=146.0, default_om=-15.4, default_tm=-18.4, columns=1
        )
        disc_opt = ConfigOption(
            name='discussions',
            label='Discussions (讨论版, 小页面)',
            is_default=False,
            page_geometry=disc_geo,
            doc_options=[],
        )
        options.append(disc_opt)

    return TemplateConfig(
        cls_path=cls_path,
        cls_name=cls_name,
        options=options,
        bleed_mm=bleed_mm,
    )


# ─── 选择与缓存 ────────────────────────────────────────────────

def select_config(config: TemplateConfig, default_name: str = 'manuscript',
                  cache_path: str = None) -> ConfigOption:
    """选择配置，默认 manuscript，支持缓存"""
    # 检查缓存
    if cache_path and Path(cache_path).exists():
        try:
            cached = json.loads(Path(cache_path).read_text(encoding='utf-8'))
            cached_name = cached.get('selected_mode')
            if cached_name:
                for opt in config.options:
                    if opt.name == cached_name:
                        config.selected = opt
                        print(f'  [模板配置] 使用缓存: {opt.label}')
                        return opt
        except Exception:
            pass

    # 默认选择
    default_opt = None
    for opt in config.options:
        if opt.name == default_name:
            default_opt = opt
            break
    if default_opt is None and config.options:
        default_opt = config.options[0]

    config.selected = default_opt
    if default_opt:
        print(f'  [模板配置] 默认选择: {default_opt.label}')

    # 保存缓存
    if cache_path and default_opt:
        try:
            cache_data = {
                'cls_path': config.cls_path,
                'selected_mode': default_opt.name,
                'page_geometry': default_opt.page_geometry.to_dict() if default_opt.page_geometry else None,
            }
            Path(cache_path).write_text(json.dumps(cache_data, ensure_ascii=False, indent=2), encoding='utf-8')
        except Exception:
            pass

    return default_opt


# ─── 统一入口 ──────────────────────────────────────────────────

def get_template_config(cls_path: str, output_dir: str = None,
                        mode: str = None) -> TemplateConfig:
    """统一入口：解析 .cls + 选择配置 + 返回

    默认选择 manuscript（经典版），不需要交互UI。
    """
    config = parse_cls_options(cls_path)

    cache_path = None
    if output_dir:
        cache_path = str(Path(output_dir) / 'template_config_selection.json')

    if mode:
        for opt in config.options:
            if opt.name == mode:
                config.selected = opt
                print(f'  [模板配置] 指定模式: {opt.label}')
                return config

    select_config(config, default_name='manuscript', cache_path=cache_path)

    return config


def get_page_geometry_for_mode(cls_content: str, config_mode: str = None) -> Optional[dict]:
    """根据配置模式获取页面几何（供 _docx_insert.py 调用）

    Args:
        cls_content: .cls 文件内容
        config_mode: 配置模式名（None 则默认 manuscript）

    Returns:
        dict: 页面几何参数，或 None
    """
    if not cls_content:
        return None

    config_mode = 'manuscript' if config_mode in (None, '', 'classic') else config_mode
    result = _scan_cls_for_page_params(cls_content)
    bleed_mm = result['bleed_mm']

    if config_mode == 'manuscript':
        params = result['manuscript']
    elif config_mode == 'final':
        params = result['final_pub']
    elif config_mode == 'discussions':
        params = result['discussions']
    else:
        params = result['manuscript']

    if not params:
        return None

    pw = params.get('paperwidth', 210.0 + 2 * bleed_mm)
    ph = params.get('paperheight', 240.0 + 2 * bleed_mm)
    tw = params.get('textwidth', 177.0)
    om = params.get('oddsidemargin', 16.4)
    tm = params.get('topmargin', 10.0)

    right_margin = round(pw - tw - om, 2)

    return {
        'paperwidth_mm': float(pw),
        'paperheight_mm': float(ph),
        'textwidth_mm': float(tw),
        'textheight_mm': float(params.get('textheight')) if params.get('textheight') is not None else None,
        'oddsidemargin_mm': float(om),
        'right_margin_mm': float(right_margin),
        'topmargin_mm': float(tm),
        'column_count': params.get('column_count', 1),
        'column_sep_mm': params.get('columnsep'),
    }


def extract_section_spacing(cls_content: str, body_size_pt: float = 11.0,
                            config_mode: str = None) -> dict:
    """从CLS提取section/subsection/subsubsection的beforeskip/afterskip

    解析 \\def\\section{\\@startsection{...}{...}{...}{beforeskip}{afterskip}{...}} 格式。

    Args:
        cls_content: .cls 文件内容
        body_size_pt: 正文字号(pt), 用于em单位转换
        config_mode: 配置模式名 (None/classic/manuscript/final/discussions)

    Returns:
        dict: {
            'section': {'before_pt': float, 'after_pt': float, 'bold': bool, 'sans': bool},
            'subsection': {'before_pt': float, 'after_pt': float, 'bold': bool, 'sans': bool},
            'subsubsection': {'before_pt': float, 'after_pt': float, 'bold': bool, 'sans': bool},
        }
    """
    if not cls_content:
        return {}

    # 根据config_mode选择对应的section定义分支
    # manuscript模式: 在 if@stage@final 的 classical 分支 (1013-1025行)
    # final/sansserifface: 在 if@sansserifface 分支 (1001-1012行)
    # discussions: 在 else%discussions 分支 (1027-1039行)
    config_mode = 'manuscript' if config_mode in (None, '', 'classic') else config_mode

    # 提取所有 \@startsection 定义
    # 格式: \def\section{\@dolinesectrue\@startsection{section}{1}{\z@}{beforeskip}{afterskip}{...}}
    section_pattern = re.compile(
        r'\\def\\(section|subsection|subsubsection)\s*\{'
        r'[^}]*\\@startsection\{[^}]+\}\{[^}]+\}\{[^}]*\}'
        r'\s*\{([^}]+)\}'   # beforeskip (第4参数)
        r'\s*\{([^}]+)\}'   # afterskip (第5参数)
        r'\s*\{([^}]+)\}',  # 格式命令 (第6参数)
        re.DOTALL
    )

    # 找到所有section定义的上下文区域
    # 策略: 按条件分支分区提取
    result = {}

    # 找到 \if@stage@final 区域
    final_start = cls_content.find('\\if@stage@final')
    disc_start = cls_content.find('\\else%discussions')

    if final_start < 0:
        # 没有条件分支，直接解析所有定义
        zone = cls_content
    elif config_mode in ('manuscript',):
        # manuscript: final区域内, classical分支(\else之后, \fi之前)
        # 先找final区域
        if disc_start > final_start:
            zone = cls_content[final_start:disc_start]
        else:
            zone = cls_content[final_start:]
        # 在zone内找classical分支 (在 \else 之后，不含 \if@sansserifface)
        sans_end = zone.find('\\else%classical')
        if sans_end > 0:
            zone = zone[sans_end:]
    elif config_mode == 'final':
        zone = cls_content[final_start:disc_start] if disc_start > final_start else cls_content[final_start:]
        # final版默认使用sansserifface分支(如果有的话)
        sans_start = zone.find('\\if@sansserifface')
        if sans_start > 0:
            # 取sansserifface的true分支
            sans_else = zone.find('\\else', sans_start)
            if sans_else > 0:
                zone = zone[sans_start:sans_else]
    elif config_mode == 'discussions':
        if disc_start > 0:
            zone = cls_content[disc_start:]
        else:
            zone = cls_content
    else:
        zone = cls_content

    for m in section_pattern.finditer(zone):
        sec_name = m.group(1)
        before_raw = m.group(2).strip()
        after_raw = m.group(3).strip()
        format_cmd = m.group(4).strip()

        # 解析 skip 值 (支持 em, ex, pt, \p@ 等)
        def _parse_skip(raw, body_pt):
            """解析LaTeX skip值, 负号表示不缩进首行(Word取绝对值)"""
            raw = raw.strip()
            # 去掉 \@plus/\@minus 弹性部分
            raw = re.split(r'\\@plus|\\@minus', raw)[0].strip()

            # 负号
            sign = 1.0
            if raw.startswith('-'):
                sign = -1.0
                raw = raw[1:].strip()

            # em 单位
            em_m = re.match(r'([\d.]+)\s*em', raw)
            if em_m:
                return sign * float(em_m.group(1)) * body_pt

            # ex 单位 (1ex ≈ 0.43em ≈ 0.43 * body_pt)
            ex_m = re.match(r'([\d.]+)\s*ex', raw)
            if ex_m:
                return sign * float(ex_m.group(1)) * body_pt * 0.43

            # pt 单位
            pt_m = re.match(r'([\d.]+)\s*\\?p@?', raw)
            if pt_m:
                return sign * float(pt_m.group(1))

            # 纯数字 (默认pt)
            num_m = re.match(r'([\d.]+)', raw)
            if num_m:
                return sign * float(num_m.group(1))

            return 0.0

        before_pt = abs(_parse_skip(before_raw, body_size_pt))
        after_pt = abs(_parse_skip(after_raw, body_size_pt))
        bold = '\\bfseries' in format_cmd
        sans = '\\sffamily' in format_cmd

        result[sec_name] = {
            'before_pt': round(before_pt, 1),
            'after_pt': round(after_pt, 1),
            'bold': bold,
            'sans': sans,
        }

    # 如果没有提取到，返回合理的默认值
    if 'section' not in result:
        result['section'] = {'before_pt': 22.0, 'after_pt': 11.0, 'bold': True, 'sans': False}
    if 'subsection' not in result:
        result['subsection'] = {'before_pt': 11.0, 'after_pt': 11.0, 'bold': True, 'sans': False}
    if 'subsubsection' not in result:
        result['subsubsection'] = {'before_pt': 11.0, 'after_pt': 11.0, 'bold': True, 'sans': False}

    return result


def extract_page_footer_dims(cls_content: str, config_mode: str = None) -> dict:
    """从CLS提取页面底部/头部尺寸参数

    Args:
        cls_content: .cls 文件内容
        config_mode: 配置模式名

    Returns:
        dict: {
            'footskip_mm': float,
            'headheight_mm': float,
            'headsep_mm': float,
            'textheight_mm': float or None,
            'bottom_margin_mm': float or None,
        }
    """
    if not cls_content:
        return {}

    config_mode = 'manuscript' if config_mode in (None, '', 'classic') else config_mode
    result = _scan_cls_for_page_params(cls_content)

    if config_mode == 'manuscript':
        params = result['manuscript']
    elif config_mode == 'final':
        params = result['final_pub']
    elif config_mode == 'discussions':
        params = result['discussions']
    else:
        params = result['manuscript']

    footskip_mm = params.get('footskip', 10.6)  # 30pt ≈ 10.6mm
    headheight_mm = params.get('headheight', 0.0)
    headsep_mm = params.get('headsep', 0.0)
    textheight_mm = params.get('textheight')

    # 计算底边距: bottom = paperheight - topmargin - textheight - headheight - headsep
    bottom_margin_mm = None
    geo = get_page_geometry_for_mode(cls_content, config_mode=config_mode)
    if geo and textheight_mm is not None:
        bottom_margin_mm = round(
            geo['paperheight_mm'] - geo['topmargin_mm'] - textheight_mm - headheight_mm - headsep_mm,
            1
        )
        # 如果计算值为负，说明footskip等参数已包含在textheight中
        if bottom_margin_mm < 0:
            bottom_margin_mm = footskip_mm
        # 如果底边距异常大（>40mm），说明textheight被错误提取
        # 常见原因：discussions模式的textheight被误分配给final模式
        if bottom_margin_mm > 40:
            # 回退：使用manuscript模式的textheight重新计算
            ms_params = result.get('manuscript', {})
            ms_textheight = ms_params.get('textheight')
            if ms_textheight:
                bottom_margin_mm = round(
                    geo['paperheight_mm'] - geo['topmargin_mm'] - ms_textheight - headheight_mm - headsep_mm,
                    1
                )
            # 仍然异常则使用footskip
            if bottom_margin_mm is not None and (bottom_margin_mm < 0 or bottom_margin_mm > 40):
                bottom_margin_mm = footskip_mm + 10  # footskip + 合理余量

    return {
        'footskip_mm': float(footskip_mm),
        'headheight_mm': float(headheight_mm),
        'headsep_mm': float(headsep_mm),
        'textheight_mm': float(textheight_mm) if textheight_mm is not None else None,
        'bottom_margin_mm': float(bottom_margin_mm) if bottom_margin_mm is not None else footskip_mm,
    }


# ─── CLI 测试 ──────────────────────────────────────────────────

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')

    if len(sys.argv) < 2:
        print('用法: python template_config.py <cls_path> [--mode <name>]')
        sys.exit(1)

    cls_path = sys.argv[1]
    mode = None
    if '--mode' in sys.argv:
        idx = sys.argv.index('--mode')
        if idx + 1 < len(sys.argv):
            mode = sys.argv[idx + 1]

    config = get_template_config(cls_path, mode=mode)

    print(f'\n模板: {config.cls_name}')
    print(f'Bleed: {config.bleed_mm}mm')
    print(f'选项数: {len(config.options)}')
    for opt in config.options:
        sel = ' ← 选中' if config.selected and opt.name == config.selected.name else ''
        print(f'\n  [{opt.name}] {opt.label}{sel}')
        if opt.page_geometry:
            g = opt.page_geometry
            print(f'    paper: {g.paperwidth_mm}x{g.paperheight_mm}mm')
            print(f'    textwidth: {g.textwidth_mm}mm')
            print(f'    margins: L={g.oddsidemargin_mm}mm, T={g.topmargin_mm}mm')
            print(f'    columns: {g.column_count}')
