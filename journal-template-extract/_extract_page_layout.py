#!/usr/bin/env python3
r"""
页面布局/列/字体/几何提取模块
包含方法:
  - _extract_page_layout
  - _eval_dimexpr
  - _extract_columns
  - _extract_fonts
"""
import re
from collections import OrderedDict
from pathlib import Path
import sys
SKILL_DIR = Path(__file__).resolve().parent.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from shared.latex_parse_utils import (
    BS, _cmd, _len_to_mm, FONT_CODE_TO_NAME
)


def _extract_page_layout(self):
    """提取页面布局规格"""
    layout = OrderedDict()

    # 纸张大小
    for name in ('paperwidth', 'paperheight'):
        # 1) \paperwidth\dimexpr210mm+2\bleed\relax
        m = re.search(BS + name + BS + r'dimexpr\s*([^%]+?)' + BS + r'relax', self.content)
        if m:
            layout[name] = self._eval_dimexpr(m.group(1).strip())
            continue
        # 2) \paperwidth166mm
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:mm|cm|pt|in|bp))', self.content)
        if m:
            layout[name] = _len_to_mm(m.group(1).strip())
            continue
        # 3) \setlength
        m = re.search(BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            val = m.group(1).strip()
            if BS + 'dimexpr' in val:
                dm = re.search(BS + r'dimexpr\s*([^%]+?)' + BS + r'relax', val)
                layout[name] = self._eval_dimexpr(dm.group(1).strip()) if dm else val
            else:
                layout[name] = _len_to_mm(val)

    # \geometry
    gm = re.search(BS + r'geometry\{([^}]+)\}', self.content)
    if gm:
        layout['geometry'] = gm.group(1)

    # 边距
    for name in ('oddsidemargin', 'evensidemargin', 'topmargin',
                  'textheight', 'textwidth', 'headheight', 'headsep',
                  'footskip', 'marginparwidth', 'columnsep', 'columnwidth'):
        m = re.search(BS + r'setlength\{' + BS + name + r'\}\{([^}]+)\}', self.content)
        if m:
            layout[name] = _len_to_mm(m.group(1).strip())
            continue
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:mm|cm|pt|in|bp|em|ex))', self.content)
        if m:
            layout[name] = _len_to_mm(m.group(1).strip())

    # marginparsep
    m = re.search(BS + r'setlength\{' + BS + r'marginparsep\}\{([^}]+)\}', self.content)
    if m:
        layout['marginparsep'] = _len_to_mm(m.group(1).strip())
    else:
        m = re.search(BS + r'marginparsep\s*=?\s*([\d.]+\s*(?:mm|cm|pt|in|bp|em|ex))', self.content)
        if m:
            layout['marginparsep'] = _len_to_mm(m.group(1).strip())

    # \setlength 扫描补充
    for m in re.finditer(BS + r'setlength\{' + BS + r'(\w+)\}\{([^}]+)\}', self.content):
        name = m.group(1)
        layout_names = ('paperwidth', 'paperheight', 'oddsidemargin', 'evensidemargin',
                    'topmargin', 'textheight', 'textwidth', 'headheight', 'headsep',
                    'footskip', 'marginparwidth', 'marginparsep', 'columnsep', 'columnwidth')
        if name in layout_names and name not in layout:
            layout[name] = _len_to_mm(m.group(2).strip())

    # bleed (裁切出血)
    m = re.search(BS + r'bleed\s*([\d.]+\s*(?:mm|cm|pt|in))', self.content)
    if m:
        layout['bleed'] = _len_to_mm(m.group(1).strip())

    # topskip / maxdepth
    for name in ('topskip', 'maxdepth'):
        m = re.search(BS + name + r'\s*=?\s*([\d.]+\s*(?:pt|mm|cm|in|bp|em|ex))', self.content)
        if m:
            layout[name] = _len_to_mm(m.group(1).strip())

    return layout


def _eval_dimexpr(self, expr):
    """计算 \\dimexpr 表达式，支持变量替换"""
    # 先提取 \bleed 等变量的值
    bleed_m = re.search(BS + r'bleed\s*([\d.]+\s*(?:mm|cm|pt|in))', self.content)
    bleed_val = _len_to_mm(bleed_m.group(1).strip()) if bleed_m else '3mm'
    bleed_mm = float(bleed_val.replace('mm', '')) if 'mm' in bleed_val else 3.0

    # 替换 N\bleed 为 N*bleed_mm (如 2\bleed → 6mm)
    expr = re.sub(r'(\d+)\s*' + BS + r'bleed',
                  lambda m: f'{float(m.group(1)) * bleed_mm}mm', expr)
    # 替换 \bleed（无系数）为 bleed_mm
    expr = re.sub(BS + r'bleed', f'{bleed_mm}mm', expr)

    # 处理 -Nmm 格式的 \p@ 等单位
    # 计算表达式: 逐步匹配 数值+单位 加减 数值+单位
    result = 0.0
    tokens = re.findall(r'([+-]?)\s*([\d.]+)\s*(mm|cm|pt|in|em|ex|p@)', expr)
    if not tokens:
        return expr
    for sign, val, unit in tokens:
        v = float(val)
        if unit == 'p@':
            v = v * 0.3514598
        elif unit == 'em':
            v = v * self.base_size * 0.3514598
        elif unit == 'ex':
            v = v * self.base_size * 0.5 * 0.3514598
        elif unit in ('mm', 'cm', 'pt', 'in'):
            converted = _len_to_mm(f'{v}{unit}')
            v = float(converted.replace('mm', ''))
        if sign == '-':
            result -= v
        else:
            result += v
    return f'{result:.1f}mm'


def _extract_columns(self):
    """提取栏数设置"""
    if re.search(BS + r'twocolumn', self.content):
        if re.search(BS + r'onecolumn', self.content):
            return 'twocolumn (with onecolumn option available)'
        return 'twocolumn'
    if re.search(BS + r'onecolumn', self.content):
        return 'onecolumn'
    if 'multicol' in self.content:
        return 'multicol'
    return 'onecolumn (default)'


def _extract_fonts(self):
    """提取字体族设置"""
    fonts = OrderedDict()
    for cmd, key in [('rmdefault', 'serif'), ('sfdefault', 'sans'),
                     ('ttdefault', 'mono'), ('bfdefault', 'bold_series'),
                     ('itdefault', 'italic_shape')]:
        m = re.search(BS + cmd + r'\{(\w+)\}', self.content)
        if m:
            fonts[key] = m.group(1)
            fonts[key + '_name'] = FONT_CODE_TO_NAME.get(m.group(1), m.group(1))

    for m in re.finditer(BS + r'RequirePackage(?:\[.*?\])?\{([^}]+)\}', self.content):
        pkg = m.group(1).strip()
        if pkg in ('times', 'mathptmx', 'newtxtext', 'ptm'):
            fonts['serif'] = 'ptm'; fonts['serif_name'] = 'Times New Roman'
        elif pkg in ('helvet', 'phv', 'newtxsf'):
            fonts['sans'] = 'phv'; fonts['sans_name'] = 'Helvetica/Arial'
        elif pkg in ('courier', 'pcr', 'newtxtt'):
            fonts['mono'] = 'pcr'; fonts['mono_name'] = 'Courier New'
        elif pkg == 'lmodern':
            fonts['serif'] = 'lmr'; fonts['serif_name'] = 'Latin Modern Roman'
            fonts['sans'] = 'lms'; fonts['sans_name'] = 'Latin Modern Sans'
            fonts['mono'] = 'lmt'; fonts['mono_name'] = 'Latin Modern Mono'
        elif pkg in ('mathtime', 'mtpro', 'mtpro2'):
            fonts['math'] = 'mathtime'; fonts['math_name'] = 'MathTime Pro'

    # \let别名: \let\rmdefault\sfdefault → serif继承sans的值
    for src_cmd, src_key in [('sfdefault', 'sans'), ('ttdefault', 'mono')]:
        if src_key not in fonts:
            m = re.search(BS + r'let\s*' + BS + src_cmd + r'\s*' + BS + r'(\w+)', self.content)
            if m:
                fonts[src_key] = m.group(1)
                fonts[src_key + '_name'] = FONT_CODE_TO_NAME.get(m.group(1), m.group(1))
    # \let\rmdefault\sfdefault → serif被设为sans字体族
    let_rm = re.search(BS + r'let\s*' + BS + r'rmdefault\s*' + BS + r'(sfdefault|ttdefault)', self.content)
    if let_rm:
        alias = let_rm.group(1)
        alias_key = 'sans' if alias == 'sfdefault' else 'mono'
        if alias_key in fonts:
            fonts['serif'] = fonts[alias_key]
            fonts['serif_name'] = fonts.get(alias_key + '_name', fonts[alias_key])
            fonts['serif_is_alias'] = alias

    mf = re.search(BS + r'mathdefault\{(\w+)\}', self.content)
    if mf:
        fonts['math'] = mf.group(1)
    return fonts
