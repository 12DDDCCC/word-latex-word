"""LaTeX文本工具 — 转义、括号匹配、上下标转换

3个skill中重复的LaTeX文本处理逻辑合并到此模块：
- tikz_table_gen: esc() — LaTeX转义
- latex_to_omml: to_subscript(), to_superscript(), clean_latex_text()
- tex_to_word: _match_balanced_braces(), _clean_latex_text()
- latex_table_parser: _match_balanced_braces(), _to_subscript()
"""

import re

_SUBSCRIPT_MAP = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'k': 'ₖ',
    'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ', 'p': 'ₚ',
    'r': 'ᵣ', 's': 'ₛ', 't': 'ₜ', 'x': 'ₓ',
    '+': '₊', '-': '₋', 'o': 'ₒ',
}

_SUPERSCRIPT_MAP = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
    'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'i': 'ⁱ', 'j': 'ʲ',
    'k': 'ᵏ', 'l': 'ˡ', 'm': 'ᵐ', 'n': 'ⁿ', 'o': 'ᵒ',
    'p': 'ᵖ', 'r': 'ʳ', 's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ',
    'v': 'ᵛ', 'w': 'ᵂ', 'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ',
    '+': '⁺', '-': '⁻', 'o': 'ᵒ',
}


def match_balanced_braces(text, start):
    """从start位置匹配平衡的大括号内容, 返回内容字符串"""
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


_ESCAPE_MAP = {
    chr(92): chr(92) + "textbackslash{}",   # backslash
    "%": chr(92) + "%",
    "_": chr(92) + "_",
    "&": chr(92) + "&",
    "#": chr(92) + "#",
    "$": chr(92) + "$",
    "{": chr(92) + "{",
    "}": chr(92) + "}",
    "~": chr(92) + "textasciitilde{}",
    "^": chr(92) + "textasciicircum{}",
}


def escape_latex(text):
    r"""转义LaTeX特殊字符（用于表格单元格文本）。

    逐字符映射，反斜杠最先处理；注入命令中的花括号不会被二次转义。
    花括号不会被二次转义。
    """
    return "".join(_ESCAPE_MAP.get(ch, ch) for ch in text)


def to_subscript(s):
    """将字符串转为 Unicode 下标，如 '2' → '₂'"""
    return ''.join(_SUBSCRIPT_MAP.get(c, c) for c in s)


def to_superscript(s):
    """将字符串转为 Unicode 上标，如 '2' → '²'"""
    return ''.join(_SUPERSCRIPT_MAP.get(c, c) for c in s)


def citation_marker(key, mode):
    """Create a Pandoc-safe marker that preserves a citation key."""
    encoded_key = str(key).encode('utf-8').hex()
    return f'CITELINK{mode}{encoded_key}END'


def clean_latex_text(text):
    """清理 LaTeX 文本中的命令，转为纯文本显示

    CO$_2$ → CO₂, \\textbf{word} → word 等
    """
    text = re.sub(r'CO\$_2\$', 'CO₂', text)
    text = re.sub(r'XCO\$_2\$', 'XCO₂', text)
    text = re.sub(r'\$_\{([^}]+)\}\$', lambda m: to_subscript(m.group(1)), text)
    text = re.sub(r'\$_([^$]+)\$', lambda m: to_subscript(m.group(1)), text)
    text = re.sub(r'\$\^\{([^}]+)\}\$', lambda m: to_superscript(m.group(1)), text)
    text = re.sub(r'\$\^([^$]+)\$', lambda m: to_superscript(m.group(1)), text)
    text = re.sub(r'\\textbf\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\textit\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\emph\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\textsubscript\{([^}]+)\}', lambda m: to_subscript(m.group(1)), text)
    text = re.sub(r'\\underline\{([^}]+)\}', r'\1', text)
    text = re.sub(r'\\\w+\{([^}]*)\}', r'\1', text)
    text = re.sub(r'[${}\\]', '', text)
    return text.strip()
