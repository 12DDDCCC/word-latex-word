"""LaTeX解析工具 — 跨skill使用的LaTeX常量和解析辅助函数

2个skill中重复的LaTeX解析逻辑合并到此模块：
- journal-template-extract/layout_spec_extract: 常量块 + _cmd + _la_size_to_pt + _size_style + _len_to_mm
- template-extract-lite/template_extract_lite: BS, _cmd
"""

import re

# ─── LaTeX字号常量 ───────────────────────────────────────────
LATEX_SIZE_PT = {
    'tiny': 5, 'scriptsize': 7, 'footnotesize': 8,
    'small': 9, 'normalsize': 10, 'large': 10.95,
    'Large': 11.49, 'LARGE': 12, 'huge': 14.4, 'Huge': 17.28,
}
LATEX_SIZE_PT_11 = {
    'tiny': 6, 'scriptsize': 8, 'footnotesize': 9,
    'small': 9.5, 'normalsize': 11, 'large': 11.5,
    'Large': 12, 'LARGE': 14.4, 'huge': 17.28, 'Huge': 20.74,
}
LATEX_SIZE_PT_12 = {
    'tiny': 6, 'scriptsize': 8, 'footnotesize': 10,
    'small': 10.95, 'normalsize': 12, 'large': 12.5,
    'Large': 14.4, 'LARGE': 17.28, 'huge': 20.74, 'Huge': 24.88,
}

WEIGHT_MAP = {
    'bfseries': 'bold', 'textbf': 'bold', 'bf': 'bold',
    'mdseries': 'normal', 'textmd': 'normal',
}
SHAPE_MAP = {
    'itshape': 'italic', 'textit': 'italic', 'it': 'italic',
    'slshape': 'slanted', 'textsl': 'slanted',
    'upshape': 'normal', 'textup': 'normal',
    'scshape': 'smallcaps', 'textsc': 'smallcaps',
}
FONT_CODE_TO_NAME = {
    'ptm': 'Times New Roman', 'cmr': 'Computer Modern Roman',
    'phv': 'Helvetica/Arial', 'cmss': 'Computer Modern Sans',
    'pcr': 'Courier New', 'cmtt': 'Computer Modern Mono',
    'ppl': 'Palatino', 'pbk': 'Bookman', 'pnc': 'New Century Schoolbook',
    'pag': 'Avant Garde', 'bch': 'Charter', 'lmr': 'Latin Modern Roman',
    'lms': 'Latin Modern Sans', 'lmt': 'Latin Modern Mono',
}

# ─── Python 3.12+ 正则兼容：用 re.escape(chr(92)) 替代 r'\\' ─
BS = re.escape(chr(92))  # 正则中匹配单个反斜杠的常量


def cmd(name):
    """构建匹配 LaTeX 命令 \\name 的正则片段"""
    return BS + re.escape(name)


def la_size_to_pt(name, base=10):
    """LaTeX字号名转pt值"""
    m = {'10': LATEX_SIZE_PT, '11': LATEX_SIZE_PT_11, '12': LATEX_SIZE_PT_12}
    return m.get(str(base), LATEX_SIZE_PT).get(name)


def size_style(text, base_size=10):
    """从LaTeX代码片段中提取字号/字重/字形"""
    size_map = {'10': LATEX_SIZE_PT, '11': LATEX_SIZE_PT_11, '12': LATEX_SIZE_PT_12}.get(
        str(base_size), LATEX_SIZE_PT)
    r = {'size_pt': None, 'size_name': None, 'weight': 'normal', 'shape': 'normal', 'alignment': None}

    for name in sorted(size_map, key=lambda x: -size_map[x]):
        if re.search(BS + re.escape(name) + r'\b', text):
            r['size_name'] = name
            r['size_pt'] = size_map[name]
            break

    if r['size_pt'] is None and re.search(BS + r'normalfont\b', text):
        r['size_name'] = 'normalsize'
        r['size_pt'] = base_size

    fsm = re.search(BS + r'fontsize\{([\d.]+)\}', text)
    if fsm:
        r['size_pt'] = float(fsm.group(1))
        r['size_name'] = f'{fsm.group(1)}pt'

    for kw, val in WEIGHT_MAP.items():
        if re.search(BS + re.escape(kw) + r'\b', text, re.IGNORECASE):
            r['weight'] = val
            break

    for kw, val in SHAPE_MAP.items():
        if re.search(BS + re.escape(kw) + r'\b', text, re.IGNORECASE):
            r['shape'] = val
            break

    if re.search(BS + r'centering\b', text):
        r['alignment'] = 'center'
    elif re.search(BS + r'raggedright\b', text):
        r['alignment'] = 'left'
    elif re.search(BS + r'raggedleft\b', text):
        r['alignment'] = 'right'

    return r


def len_to_mm(val):
    """将LaTeX长度转为mm"""
    m = re.match(r'([\d.]+)\s*(pt|mm|cm|in|em|ex|bp|dd|cc|sp)', val.strip())
    if not m:
        return val
    num, unit = float(m.group(1)), m.group(2)
    conv = {'pt': 0.3515, 'mm': 1, 'cm': 10, 'in': 25.4, 'bp': 0.3528,
            'dd': 0.376, 'cc': 1.128, 'sp': 2.54e-5}
    if unit in ('em', 'ex'):
        return f'{val} (≈{num * 2.5:.1f}mm at 10pt)'
    return f'{num * conv.get(unit, 1):.1f}mm'


def find_balanced_braces(text, start):
    """从start位置匹配平衡的大括号内容, 返回内容字符串

    与 latex_text_utils.match_balanced_braces 签名一致，此版本作为
    兼容别名保留。推荐使用 shared.latex_text_utils 版本。
    """
    depth = 0
    i = start
    while i < len(text):
        if text[i] == '{':
            depth += 1
        elif text[i] == '}':
            depth -= 1
            if depth == 0:
                return text[start + 1:i]
        i += 1
    return text[start + 1:]


# ─── 兼容别名（子模块使用旧名 _cmd/_la_size_to_pt/_size_style/_len_to_mm/_read）──
_cmd = cmd
_la_size_to_pt = la_size_to_pt
_size_style = size_style
_len_to_mm = len_to_mm
def _read(path):
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        return f.read()
