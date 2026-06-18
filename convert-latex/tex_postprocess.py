#!/usr/bin/env python3
"""后处理 + 文件复制 + 编译

包含：
- postprocess_tex: 后处理tex文件（化学式、URL、引用等修复）
- copy_support_files: 复制支撑文件到输出目录
- compile_tex: 按 cls 选择引擎编译（copernicus→lualatex，其余→xelatex，失败兜底 lualatex）
"""

import re
import shutil
import subprocess
import os
from pathlib import Path

try:
    from skeleton_builder import sanitize_bib_filename
except Exception:
    def sanitize_bib_filename(value, default='references'):
        raw = str(value or '').strip().strip('{}').strip()
        if not raw:
            return default
        raw = raw.split(',')[0].strip().strip('"\'')
        raw = raw.replace('\\', '/').split('/')[-1].strip()
        if raw.lower().endswith('.bib'):
            raw = raw[:-4]
        lowered = raw.lower()
        if (not raw or raw in ('.', '..') or
                any(token in lowered for token in ('your', 'bibdatabase', 'bib database',
                                                   'bibfile', 'bib file', 'filename', 'database')) or
                re.search(r'[<>:"/\\|?*\s]', raw)):
            return default
        return raw


def _strip_latex_comments(tex_content):
    """Remove comments while preserving escaped percent signs."""
    lines = []
    for line in (tex_content or "").splitlines():
        cut = len(line)
        for match in re.finditer('%', line):
            pos = match.start()
            backslashes = 0
            i = pos - 1
            while i >= 0 and line[i] == '\\':
                backslashes += 1
                i -= 1
            if backslashes % 2 == 0:
                cut = pos
                break
        lines.append(line[:cut])
    return "\n".join(lines)


def _extract_latex_citation_keys(tex_content):
    """Return citation keys used by common LaTeX citation commands."""
    cleaned = _strip_latex_comments(tex_content)
    keys = []
    cite_commands = (
        'cite', 'citep', 'citet', 'citealp', 'citealt', 'citeauthor',
        'citeyear', 'citeyearpar', 'citepos', 'citepalias', 'citetalias',
        'onlinecite', 'supercite', 'parencite', 'textcite', 'autocite',
        'footcite', 'nocite',
    )
    cite_re = re.compile(
        r'\\(?P<cmd>' + '|'.join(cite_commands) + r')\*?'
        r'\s*(?:\[[^\]]*\]\s*){0,2}\{(?P<keys>[^}]*)\}'
    )
    for match in cite_re.finditer(cleaned):
        for key in match.group('keys').split(','):
            key = key.strip()
            if key and key != '*':
                keys.append(key)
    return list(dict.fromkeys(keys))


def _extract_bib_entry_keys(bib_content):
    """Return entry keys declared by a BibTeX database."""
    keys = []
    for match in re.finditer(r'@\s*([A-Za-z]+)\s*\{\s*([^,\s]+)', bib_content or ""):
        entry_type = match.group(1).lower()
        if entry_type in {'comment', 'preamble', 'string'}:
            continue
        keys.append(match.group(2).strip())
    return set(keys)


def _bibliography_paths_for_tex(tex_content, tex_path, output_dir):
    cleaned = _strip_latex_comments(tex_content)
    names = []
    for match in re.finditer(r'\\bibliography\{([^}]*)\}', cleaned):
        names.extend(name.strip() for name in match.group(1).split(',') if name.strip())
    if not names:
        return []

    paths = []
    tex_dir = Path(tex_path).parent
    output_dir = Path(output_dir)
    for name in names:
        bib_name = name if name.endswith('.bib') else f'{name}.bib'
        candidate = Path(bib_name)
        if candidate.is_absolute():
            paths.append(candidate)
        else:
            paths.append(output_dir / candidate)
            paths.append(tex_dir / candidate)
    return list(dict.fromkeys(paths))


def validate_bibliography_keys(tex_path, output_dir):
    """Validate that citation keys in tex exist in the referenced BibTeX files."""
    tex_path = Path(tex_path)
    output_dir = Path(output_dir)
    try:
        tex_content = tex_path.read_text(encoding='utf-8', errors='ignore')
    except OSError as exc:
        return {'ok': False, 'reason': f'tex-read-failed: {exc}', 'missing': []}

    citation_keys = _extract_latex_citation_keys(tex_content)
    if not citation_keys:
        return {'ok': True, 'citation_keys': [], 'bib_keys': set(), 'missing': []}

    bib_paths = _bibliography_paths_for_tex(tex_content, tex_path, output_dir)
    if not bib_paths:
        return {
            'ok': False,
            'reason': 'bibliography-not-declared',
            'citation_keys': citation_keys,
            'bib_keys': set(),
            'missing': citation_keys,
        }

    existing_bib_paths = [path for path in bib_paths if path.exists()]
    if not existing_bib_paths:
        return {
            'ok': False,
            'reason': 'bib-file-not-found',
            'citation_keys': citation_keys,
            'bib_keys': set(),
            'missing': citation_keys,
            'bib_paths': [str(path) for path in bib_paths],
        }

    bib_keys = set()
    for bib_path in existing_bib_paths:
        try:
            bib_keys.update(_extract_bib_entry_keys(
                bib_path.read_text(encoding='utf-8', errors='ignore')))
        except OSError:
            continue

    missing = [key for key in citation_keys if key not in bib_keys]
    return {
        'ok': not missing,
        'reason': 'missing-bib-keys' if missing else '',
        'citation_keys': citation_keys,
        'bib_keys': bib_keys,
        'missing': missing,
        'bib_paths': [str(path) for path in existing_bib_paths],
    }


def _inject_layout_probe(tex_content):
    """Write the effective body dimensions to the LaTeX log for Word export."""
    if "SKILL-LAYOUT-COLUMNWIDTH" in tex_content:
        return tex_content
    probe = "\n".join([
        r"\typeout{SKILL-LAYOUT-TEXTWIDTH=\the\textwidth}",
        r"\typeout{SKILL-LAYOUT-COLUMNWIDTH=\the\columnwidth}",
        r"\typeout{SKILL-LAYOUT-COLUMNSEP=\the\columnsep}",
    ])
    body_start = re.search(
        r"(?m)^\\(?:section\{|introduction\b|subsection\{|paragraph\{)",
        tex_content,
    )
    if body_start:
        return tex_content[:body_start.start()] + probe + "\n\n" + tex_content[body_start.start():]
    return tex_content.replace(r"\end{document}", probe + "\n" + r"\end{document}", 1)


def _restore_identifier_subscripts(tex_content):
    """Undo chemical subscript markup when it appears inside identifier tokens."""
    marker = r'(?:\$_\{(?P<m1>\d{1,2})\}\$|\$_(?P<m4>\d{1,2})\$|\$\\_\\\{(?P<m5>\d{1,2})\\\}\$|\\textsubscript\{(?P<m2>\d{1,2})\}|_\{(?P<m3>\d{1,2})\})'
    token_re = re.compile(r'[A-Za-z][A-Za-z0-9_]*' + marker + r'[A-Za-z_][A-Za-z0-9_]*')

    def repl(m):
        token = m.group(0)
        return re.sub(marker, lambda sm: sm.group('m1') or sm.group('m2') or sm.group('m3') or sm.group('m4') or sm.group('m5'), token)

    return token_re.sub(repl, tex_content)


def _normalize_escaped_chemical_subscripts(tex_content):
    """Repair escaped LaTeX subscript markers produced from reference/plain text escaping."""
    tex_content = re.sub(
        r'(?<![0-9A-Za-z_])CO\$\\_\\\{2\\\}\$(?![0-9A-Za-z_])',
        r'CO$_2$',
        tex_content
    )
    tex_content = re.sub(
        r'(?<![0-9A-Za-z_])XCO\$\\_\\\{2\\\}\$(?![0-9A-Za-z_])',
        r'XCO$_2$',
        tex_content
    )
    return tex_content


def _collapse_double_reference_parentheses(tex_content):
    """Collapse duplicated parentheses only around citation-like text."""
    citation_like = (
        r'[^()\n]{0,200}'
        r'(?:\\cite[a-zA-Z]*\{[^}]+\}|\bet\s+al\.|\b\d{4}[a-z]?\b)'
        r'[^()\n]{0,200}'
    )
    return re.sub(r'\(\((' + citation_like + r')\)\)', r'(\1)', tex_content)


def _fix_mathrm_wrappers(tex_content):
    """Strip mathrm only around operators/Greek while respecting nested braces."""
    greek = (
        r'\\(?:alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|'
        r'lambda|mu|nu|xi|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega)'
    )

    def should_strip(content):
        return bool(
            re.search(r'[+=\-,\s]|\\times|\\div|\\pm', content) or
            re.search(greek, content, re.IGNORECASE)
        )

    out = []
    i = 0
    marker = r'\mathrm{'
    while i < len(tex_content):
        if not tex_content.startswith(marker, i):
            out.append(tex_content[i])
            i += 1
            continue
        start = i + len(marker)
        depth = 1
        j = start
        while j < len(tex_content) and depth:
            ch = tex_content[j]
            if ch == '\\':
                out_skip = 2 if j + 1 < len(tex_content) else 1
                j += out_skip
                continue
            if ch == '{':
                depth += 1
            elif ch == '}':
                depth -= 1
            j += 1
        if depth:
            out.append(tex_content[i])
            i += 1
            continue
        content = tex_content[start:j - 1]
        out.append(content if should_strip(content) else tex_content[i:j])
        i = j
    return ''.join(out)


def postprocess_tex(tex_content, layout_spec=None):
    """后处理tex文件：修复化学式下标、URL溢出等问题"""
    import re as _re

    tex_content = _restore_identifier_subscripts(tex_content)
    tex_content = _normalize_escaped_chemical_subscripts(tex_content)

    # 1. CO2 → CO$_2$（但保留已有的 CO$_2$、CO\textsubscript{2}、CO_{2}）
    # 只匹配独立的 CO2（后面不跟字母，避免误匹配 GCASv2 等）
    tex_content = _re.sub(
        r'(?<![0-9a-zA-Z_])CO(?!\$|_|\\textsubscript|\\mathrm)(2)(?![0-9a-zA-Z_])',
        r'CO$_2$',
        tex_content
    )
    # XCO2 → XCO$_2$
    tex_content = _re.sub(
        r'(?<![0-9a-zA-Z_])XCO(?!\$|_|\\textsubscript)(2)(?![0-9a-zA-Z_])',
        r'XCO$_2$',
        tex_content
    )
    # OCO2 → OCO-2 (卫星名称，不是化学式)
    # OCO3 → OCO-3 (卫星名称)
    tex_content = _re.sub(r'(?<![0-9A-Za-z_])OCO2(?![0-9A-Za-z_])', r'OCO-2', tex_content)
    tex_content = _re.sub(r'(?<![0-9A-Za-z_])OCO3(?![0-9A-Za-z_])', r'OCO-3', tex_content)
    tex_content = _restore_identifier_subscripts(tex_content)
    tex_content = _normalize_escaped_chemical_subscripts(tex_content)

    # 2. URL用\url{}包裹（避免超出页面边界）
    tex_content = _re.sub(
        r'(?<!\\url\{)(https?://[^\s\)]+)',
        r'\\url{\1}',
        tex_content
    )

    # 3. 压缩连续3个以上空行为2个空行
    tex_content = _re.sub(r'\n{4,}', '\n\n\n', tex_content)

    # 4. 修复引用括号问题
    # 4a. 连续 \citep{N}\citep{M} → \citep{N,M}（合并同一处引用）
    while _re.search(r'\\citep\{([^}]+)\}\\citep\{', tex_content):
        tex_content = _re.sub(
            r'\\citep\{([^}]+)\}\\citep\{([^}]+)\}',
            r'\\citep{\1,\2}',
            tex_content
        )
    # 4b. (\citep{N}) → \citep{N}（citep 本身已产生括号，外层括号多余）
    tex_content = _re.sub(r'\(\\citep\{([^}]+)\}\)', r'\\citep{\1}', tex_content)
    # 4c. 修复双重括号引用: ((Gui et al., 2024)) → (Gui et al., 2024)
    tex_content = _collapse_double_reference_parentheses(tex_content)

    # 5. 连续公式间距优化
    # 注意：Copernicus等模板中每个equation需要独立编号和label，
    # gather环境不支持每行单独编号，因此不再合并为gather。
    # 保持每个 \begin{equation}...\end{equation} 独立。

    # 6. 清理caption中重复的编号前缀
    # 延迟导入，避免循环
    from shared.caption_utils import clean_caption_prefix_in_tex
    tex_content = clean_caption_prefix_in_tex(tex_content)

    # 7. 禁用unicode-math的prime扫描
    # bbl文件中撇号(')如O'Dell会触发unicode-math的\__um_scanprime_collect
    # 导致TeX capacity exceeded。修复：在preamble中禁用该扫描。
    if 'unicode-math' in tex_content and '\\__um_scanprime_collect' not in tex_content:
        # 用ExplSyntaxOn/Off安全地重定义内部命令
        tex_content = tex_content.replace(
            '\\begin{document}',
            '\\ExplSyntaxOn\n\\cs_set_eq:NN \\__um_scanprime_collect:N \\use_none:n\n\\ExplSyntaxOff\n\\begin{document}',
            1
        )

    # 8. 修复公式中的 \mathrm 误用
    # \mathrm 只应用于罗马体文本，不应包裹运算符和希腊字母
    # 例: \mathrm{+\lambda\times} → +\lambda\times
    # 例: \mathrm{\delta} → \delta
    # 例: \mathrm{\lambda} → \lambda
    tex_content = _fix_mathrm_wrappers(tex_content)

    # 8. 修复 Unicode 下标字符
    # XeLaTeX 不能直接处理 Unicode 下标（如 CO₂ 中的 ₂）
    # 在数学模式外: CO₂ → CO\textsubscript{2}
    # 在数学模式内: 已由 unicode-math 处理，跳过
    unicode_sub_map = {
        '₀': '0', '₁': '1', '₂': '2', '₃': '3',
        '₄': '4', '₅': '5', '₆': '6', '₇': '7',
        '₈': '8', '₉': '9',
    }
    def _fix_unicode_sub(m):
        prefix = m.group(1)
        sub_char = m.group(2)
        digit = unicode_sub_map.get(sub_char, sub_char)
        return f'{prefix}\\textsubscript{{{digit}}}'
    for uc in unicode_sub_map:
        tex_content = re.sub(
            r'([A-Za-z])(' + re.escape(uc) + r')',
            _fix_unicode_sub,
            tex_content
        )

    # 9. 合并被浮动体(table/figure)截断的连续段落
    tex_content = _merge_float_adjacent_paragraphs(tex_content)

    # 10. 长公式自动换行（将过长的equation转为split环境）
    tex_content = _wrap_long_equations(tex_content, layout_spec=layout_spec)

    tex_content = _restore_identifier_subscripts(tex_content)
    tex_content = _normalize_escaped_chemical_subscripts(tex_content)
    return _inject_layout_probe(tex_content)


def _merge_float_adjacent_paragraphs(tex_content):
    """合并被浮动体(table/figure)截断的连续段落

    LaTeX中浮动体(table/figure)前后如果文本是同一段话的延续：
    - 前段不以句末标点(.!?。！？)结尾
    - 后段不以大写字母/数字/\\命令开头
    则移除浮动体前后的空行，使文本连续排版。

    例如:
      "...所以\\n\\n\\begin{table}...\\end{table}\\n\\n导致年平均..."
    →
      "...所以\\n\\begin{table}...\\end{table}\\n导致年平均..."
    """
    sentence_end_chars = set('.!?。！？')
    merged = 0

    # 匹配 \begin{table}[htbp]...\end{table} 或 \begin{figure}[htbp]...\end{figure}
    float_pattern = re.compile(
        r'(\n{2,})'                           # 前面的空行(2+)
        r'(\\begin\{(?:table|figure)\*?\}(?:\[[^\]]*\])?\n.*?\\end\{(?:table|figure)\*?\})'
        r'(\n{2,})',                           # 后面的空行(2+)
        re.DOTALL
    )

    def _check_and_merge(m):
        nonlocal merged
        before_blanklines = m.group(1)
        float_block = m.group(2)
        after_blanklines = m.group(3)

        # 找到前面的文本内容(最后一个非空行)
        before_pos = m.start()
        # 向前扫描找到前段最后一个非空字符
        i = before_pos - 1
        while i >= 0 and tex_content[i] in '\n\r':
            i -= 1
        if i < 0:
            return m.group(0)  # 文件开头，不合并
        prev_char = tex_content[i]

        # 找到后面的文本内容(第一个非空行的第一个字符)
        after_pos = m.end()
        j = after_pos
        while j < len(tex_content) and tex_content[j] in '\n\r':
            j += 1
        if j >= len(tex_content):
            return m.group(0)  # 文件结尾，不合并
        next_char = tex_content[j]

        # 跳过\end{document}等结构性命令
        if tex_content[j:j+5] == '\\end{' or tex_content[j:j+1] == '\\':
            # 如果以\开头但不是续接文字，检查是否是命令
            cmd_match = re.match(r'\\[a-zA-Z]', tex_content[j:])
            if cmd_match:
                return m.group(0)  # LaTeX命令开头，不合并

        # 合并条件：前段不以句末标点结尾 AND 后段不以大写/数字开头
        should_merge = (
            prev_char not in sentence_end_chars and
            not next_char.isupper() and
            not next_char.isdigit()
        )

        if should_merge:
            merged += 1
            # 保留1个换行而不是2+个换行(不分段)
            return '\n' + float_block + '\n'

        return m.group(0)

    tex_content = float_pattern.sub(_check_and_merge, tex_content)

    if merged:
        print(f'  [float-merge] 合并了 {merged} 处浮动体截断段落(Word→LaTeX)')

    return tex_content


def _wrap_long_equations(tex_content, layout_spec=None):
    """检测过长的 \\begin{equation} 并包装为 split 环境

    在两栏布局中，公式宽度受限于单栏（约85mm）。
    超长的公式会导致 overfull hbox 警告并超出边界。
    此函数将超长公式包装为 \\begin{equation}\\begin{split}...\\end{split}\\end{equation}
    并在主要运算符（+、-、=）处自动插入换行 \\\\ 。

    阈值：公式内容长度 > 120 个可见字符时触发换行。
    """
    # 匹配 \begin{equation}...\end{equation}（不含嵌套equation）
    eq_pattern = re.compile(
        r'(\\begin\{equation\*?\})\s*'
        r'(.*?)'
        r'\s*(\\end\{equation\*?\})',
        re.DOTALL,
    )

    # 阈值：120字符（单栏约85mm可容纳约45个字符，宽松设置以避免误判）
    line_capacity = _equation_line_capacity(layout_spec) or 120

    wrapped = 0

    def _try_wrap(m):
        nonlocal wrapped
        begin = m.group(1)
        body = m.group(2).strip()
        end = m.group(3)

        # 如果已经是split/aligned/multline等环境，跳过
        if any(env in body for env in ('\\begin{split}', '\\begin{aligned}',
                                        '\\begin{multline}', '\\begin{gather')):
            return m.group(0)

        # 如果只有一个 \label 和很短的内容，跳过
        content_no_label = re.sub(r'\\label\{[^}]+\}', '', body)
        content_no_label = re.sub(r'\\tag\{[^}]+\}', '', content_no_label).strip()
        if not _equation_exceeds_line(content_no_label, line_capacity):
            return m.group(0)

        # 在主要运算符处插入换行
        # 策略：在 = 或 + 前插入 \\ （保留运算符在行首）
        lines = _split_equation_balanced(content_no_label, line_capacity)

        if len(lines) <= 1:
            return m.group(0)

        # 提取label
        label_m = re.search(r'(\\label\{[^}]+\})', body)
        label = label_m.group(1) if label_m else ''
        tag_m = re.search(r'(\\tag\{[^}]+\})', body)
        tag = tag_m.group(1) if tag_m else ''

        # 构建 split 环境体
        split_body = ' \\\\\n  '.join(lines)
        equation_metadata = ''.join(value for value in (tag, label) if value)
        if equation_metadata:
            equation_metadata = '\n' + equation_metadata

        wrapped += 1
        return f'{begin}\n\\begin{{split}}\n  {split_body}\n\\end{{split}}{equation_metadata}\n{end}'

    result = eq_pattern.sub(_try_wrap, tex_content)
    if wrapped:
        print(f'  [eq-wrap] 包装了 {wrapped} 个超长公式(split环境)')
    return result


def _equation_line_capacity(layout_spec):
    """Estimate display equation capacity from extracted page geometry."""
    page = (layout_spec or {}).get('page_geometry', {}) if layout_spec else {}
    doc = (layout_spec or {}).get('document', {}) if layout_spec else {}
    try:
        columns = int(page.get('column_count', 1) or 1)
    except (TypeError, ValueError):
        columns = 1
    if doc.get('is_twocolumn') and columns < 2:
        columns = 2
    width_mm = page.get('columnwidth_mm') if columns > 1 else None
    if not width_mm:
        try:
            textwidth = float(page.get('textwidth_mm') or 0)
            column_sep = float(page.get('column_sep_mm') or 0)
            width_mm = (textwidth - column_sep) / 2 if columns > 1 else textwidth
        except (TypeError, ValueError):
            width_mm = 0
    try:
        width_mm = float(width_mm)
    except (TypeError, ValueError):
        return None
    if width_mm <= 0:
        return None
    return max(32.0, width_mm / 2.35)


def _equation_visible_units(eq_body):
    """Estimate compact math width without changing the equation."""
    text = re.sub(r'\\(?:label|tag)\{[^{}]*\}', '', eq_body)
    weighted_commands = {
        r'\\times': 1.0,
        r'\\quad': 2.0,
        r'\\,': 0.3,
        r'\\;': 0.6,
    }
    units = 0.0
    for pattern, value in weighted_commands.items():
        matches = len(re.findall(pattern, text))
        units += matches * value
        text = re.sub(pattern, ' ', text)
    text = re.sub(r'\\(?:mathbf|mathrm|boldsymbol|text|operatorname)\{([^{}]*)\}', r'\1', text)
    text = re.sub(r'\\[A-Za-z]+', 'x', text)
    i = 0
    while i < len(text):
        ch = text[i]
        if ch in '{} \t\r\n':
            i += 1
            continue
        if ch in '^_':
            if i + 1 < len(text) and text[i + 1] == '{':
                depth = 1
                j = i + 2
                while j < len(text) and depth:
                    if text[j] == '{':
                        depth += 1
                    elif text[j] == '}':
                        depth -= 1
                    j += 1
                units += max(0.8, 0.45 * _equation_visible_units(text[i + 2:j - 1]))
                i = j
            else:
                units += 0.6
                i += 2
            continue
        if ch.isalpha() or ch.isdigit():
            units += 0.9
        elif ch in '+-=,.:;':
            units += 0.9
        else:
            units += 0.7
        i += 1
    return units


def _equation_exceeds_line(eq_body, line_capacity):
    return _equation_visible_units(eq_body) > float(line_capacity)


def _split_equation_at_operators(eq_body):
    """在 = 或 + 运算符处拆分公式为多行

    策略：找到 = 作为第一个断点，然后在最长的 + 处断开。
    尽量只产生2-3行，避免过多断行。
    """
    # 在 = 处第一次分割（保留 = 在左边）
    eq_m = re.search(r'(?<![{\\])=', eq_body)
    if not eq_m:
        return [eq_body]

    eq_pos = eq_m.start()
    left_side = eq_body[:eq_pos + 1].rstrip()
    right_side = eq_body[eq_pos + 1:].strip()

    if not right_side:
        return [eq_body]

    # 在 + 处找最佳断点：选最长的 + 位置分段，只分2段
    # 找到所有 + 的位置
    plus_positions = [m.start() for m in re.finditer(r'(?<![{\\])\+', right_side)]

    if not plus_positions:
        # 没有+，整行放不下就只能保持原样
        return [eq_body]

    # 找中点的+位置，分成大致相等的两半
    mid = len(right_side) / 2
    best_pos = min(plus_positions, key=lambda p: abs(p - mid))

    part1 = right_side[:best_pos].strip()
    part2 = right_side[best_pos:].strip()  # 以+开头

    # 只产生2行：左边=右边前半 和 +后半
    lines = [left_side, part1 + ' \\\\\n  ' + part2]

    # 如果合并后仍然太长(>200字符)，尝试3段
    total_len = len(part1) + len(part2)
    if total_len > 200 and len(plus_positions) >= 2:
        # 分3段：找两个断点
        third1 = len(right_side) / 3
        third2 = 2 * len(right_side) / 3
        bp1 = min(plus_positions, key=lambda p: abs(p - third1))
        remaining_plus = [p for p in plus_positions if p > bp1 + 5]
        if remaining_plus:
            bp2 = min(remaining_plus, key=lambda p: abs(p - third2))
            s1 = right_side[:bp1].strip()
            s2 = right_side[bp1:bp2].strip()
            s3 = right_side[bp2:].strip()
            lines = [left_side, s1, s2, s3]
        else:
            lines = [left_side, part1, part2]

    # 过滤空段
    lines = [l for l in lines if l.strip()]
    if len(lines) <= 1:
        return [eq_body]

    return lines


def _split_equation_balanced(eq_body, line_capacity=None):
    """Split a long equation into column-safe lines at top-level plus signs."""
    eq_match = re.search(r'(?<![{\\])=', eq_body)
    if not eq_match:
        return [eq_body]
    left_side = eq_body[:eq_match.start()].rstrip()
    right_side = eq_body[eq_match.end():].strip()
    plus_positions = [m.start() for m in re.finditer(r'(?<![{\\])\+', right_side)]
    if not right_side or not plus_positions:
        return [eq_body]

    def visible_length(value):
        value = re.sub(r'\\(?:mathbf|mathrm|text|operatorname)\{([^{}]*)\}', r'\1', value)
        value = re.sub(r'\\[A-Za-z]+', 'x', value)
        return len(re.sub(r'[{}\s]', '', value))

    terms = []
    start = 0
    for pos in plus_positions:
        if pos > start:
            terms.append(right_side[start:pos].strip())
        start = pos
    terms.append(right_side[start:].strip())
    terms = [term for term in terms if term]
    if len(terms) < 2:
        return [eq_body]

    capacity = float(line_capacity or 0)
    if capacity <= 0:
        capacity = max(visible_length(eq_body) / 2.0, 35.0)
    line_limit = capacity * 0.85

    groups, current = [], []
    for term in terms:
        candidate = ''.join(current + [term])
        probe = (
            f'{left_side} ={candidate}'
            if not groups else f'\\quad {candidate}'
        )
        if current and (
            len(current) >= 2
            or _equation_visible_units(probe) > line_limit
        ):
            groups.append(current)
            current = [term]
        else:
            current.append(term)
    if current:
        groups.append(current)

    if len(groups) == 1:
        candidates = []
        for pos in plus_positions:
            first = f'{left_side} ={right_side[:pos].strip()}'
            second = right_side[pos:].strip()
            first_len, second_len = visible_length(first), visible_length(second)
            candidates.append((first_len < second_len, abs(first_len - second_len), pos))
        best_pos = min(candidates)[2]
        groups = [[right_side[:best_pos].strip()], [right_side[best_pos:].strip()]]

    lines = []
    for idx, group in enumerate(groups):
        joined = ''.join(group).strip()
        if idx == 0:
            lines.append(f'{left_side} &={joined.lstrip("+").strip()}')
        else:
            if not joined.startswith('+'):
                joined = '+' + joined
            lines.append(f'&\\quad {joined}')
    return lines


def copy_support_files(template_result, bib_path, output_dir, docx_path=None, skeleton_info=None):
    """复制支撑文件到输出目录

    优先使用docx同目录下的修改版cls(含xelatex兼容补丁)，
    其次使用模板目录下的原始版本。
    """
    src_dir = Path(template_result['output_dir'])
    output_dir = Path(output_dir)
    for f in src_dir.iterdir():
        if f.suffix in ('.cls', '.bst', '.cfg', '.sty') and f.is_file():
            dst = output_dir / f.name
            # 优先从docx所在目录复制修改版cls
            if docx_path and f.suffix == '.cls':
                docx_dir = Path(docx_path).parent
                modified_cls = docx_dir / f.name
                if modified_cls.exists():
                    shutil.copy2(str(modified_cls), str(dst))
                    continue
            if not dst.exists():
                shutil.copy2(str(f), str(dst))

    # bib 文件名从模板骨架动态提取
    bib_filename = sanitize_bib_filename((skeleton_info or {}).get('bib_filename', 'references'))
    bib_dst = output_dir / f'{bib_filename}.bib'
    if not bib_dst.exists():
        shutil.copy2(str(bib_path), str(bib_dst))


def _sanitize_bbl_urls(bbl_path):
    """Repair URL-like eprint fields emitted as unbreakable arXiv links."""
    bbl_path = Path(bbl_path)
    if not bbl_path.exists():
        return False
    text = bbl_path.read_text(encoding='utf-8', errors='ignore')

    def _bad_arxiv_href(match):
        url = match.group(2) or match.group(1)
        return r'\url{' + url + '}'

    fixed = re.sub(
        r'\\href\{http://arxiv\.org/abs/(https?://[^}]+)\}'
        r'\{\{\\tt\s+arXiv:(https?://[^}]+)\}\}',
        _bad_arxiv_href,
        text,
    )
    fixed = re.sub(
        r'\{\\tt\s+arXiv:(https?://[^}\s]+)\}',
        lambda m: r'\url{' + m.group(1) + '}',
        fixed,
    )
    fixed = re.sub(
        r'\\doi\{https?://(?:dx\.)?doi\.org/([^}]+)\}',
        lambda m: r'\doi{' + m.group(1) + '}',
        fixed,
    )
    if fixed == text:
        return False
    bbl_path.write_text(fixed, encoding='utf-8')
    return True


def _warn_bbl_long_urls(bbl_path, column_width_chars=70):
    """检测 bbl 里可能超栏的超长/重复 URL，打印 warning。

    \\url{} 默认不在内部断行，单栏（约 70 字符宽）里超过该长度的 URL
    会物理超栏。同时单条文献重复填 url+doi+href 等价 URL 会加剧超栏。
    根因在 bib 源数据，此处只做预警，不改写 bbl。
    """
    bbl_path = Path(bbl_path)
    if not bbl_path.exists():
        return
    text = bbl_path.read_text(encoding='utf-8', errors='ignore')
    # 按 bibitem 切分条目
    entries = re.split(r'(\\bibitem(?:\[[^\]]*\])?\{)', text)
    # entries[0] 是 preamble，之后成对出现 (标记, 正文)
    long_hits = 0
    dup_hits = 0
    for i in range(1, len(entries) - 1, 2):
        body = entries[i] + entries[i + 1]
        urls = re.findall(r'\\(?:url|doi|href)\{([^}]+)\}', body)
        if not urls:
            continue
        # 重复等价 URL 检测（去除协议差异后比较）
        norm = []
        for u in urls:
            s = re.sub(r'^https?://(?:dx\.)?doi\.org/', '', u)
            s = re.sub(r'^https?://', '', s).rstrip('/')
            norm.append(s.lower())
        if len(set(norm)) < len(urls):
            dup_hits += 1
        # 超长 URL 检测
        for u in urls:
            if len(u) > column_width_chars:
                long_hits += 1
                break
    if long_hits or dup_hits:
        print(f'  [bib-url-warn] 检测到可能超栏的 URL：'
              f'{long_hits} 条含>{column_width_chars}字符超长URL，'
              f'{dup_hits} 条含重复等价URL。建议清洗 references.bib 去重/缩短。')


def _fallback_strip_fullwidth_to_floats(tex_path):
    """Rewrite incompatible strip full-width blocks back to template floats."""
    tex_path = Path(tex_path)
    text = tex_path.read_text(encoding='utf-8', errors='ignore')
    if '\\begin{strip}' not in text:
        return False

    strip_re = re.compile(
        r'(?:% SKILL-FULLWIDTH-FLOAT\s+'
        r'(?P<comment_kind>figure|table)\s+pos=(?P<pos>[^\r\n]+)\s*)?'
        r'\\begin\{strip\}'
        r'(?P<body>(?:(?!\\end\{strip\}).)*?)'
        r'\\end\{strip\}',
        re.DOTALL,
    )

    def repl(match):
        body = match.group('body')
        captype = re.search(
            r'\\@captype\s*\{(figure|table)\}|'
            r'\\def\s*\\@captype\s*\{(figure|table)\}',
            body,
        )
        kind = (
            match.group('comment_kind') or
            (captype.group(1) or captype.group(2) if captype else '')
        )
        if kind not in ('figure', 'table'):
            return match.group(0)
        pos = (match.group('pos') or 't').strip() or 't'
        body = re.sub(r'\\begingroup\s*', '', body)
        body = re.sub(r'\\endgroup\s*', '', body)
        body = re.sub(
            r'\\makeatletter\\def\\@captype\{(?:figure|table)\}\\makeatother\s*',
            '',
            body,
        )
        return f'\\begin{{{kind}*}}[{pos}]\n{body.strip()}\n\\end{{{kind}*}}'

    fixed = strip_re.sub(repl, text)
    if fixed == text:
        return False
    tex_path.write_text(fixed, encoding='utf-8')
    print('  [compile-fallback] strip跨栏块回退为模板星号浮动体后重试')
    return True


def _remove_stale_pdf(tex_path):
    pdf_path = Path(tex_path).with_suffix('.pdf')
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            pass


# 需 lualatex 编译的 cls 白名单（仅含已验证的 Copernicus 系列类名）。
_LUALATEX_CLS = {
    'copernicus', 'acp', 'amt', 'angeo', 'bg', 'cp', 'esd', 'essd',
    'gmd', 'hess', 'nhess', 'os', 'se', 'tc', 'wcd',
}


def _select_engine(cls_name):
    """根据 doc class 选择编译引擎。copernicus 系列需 lualatex（cls 未适配 xelatex 的 \\pdfoutput / inputenc）。"""
    cls_lower = Path(str(cls_name or '')).stem.lower()
    if 'copernicus' in cls_lower or cls_lower in _LUALATEX_CLS:
        return 'lualatex'
    return 'xelatex'


def _latex_subprocess_env(output_dir, engine):
    if engine != 'lualatex':
        return None
    env = os.environ.copy()
    cache_dir = Path(output_dir) / '.texmf-var'
    cache_dir.mkdir(parents=True, exist_ok=True)
    env['TEXMFVAR'] = str(cache_dir.resolve())
    return env


def _run_latex_steps(tex_name, stem, output_dir, engine='xelatex', timeout=120):
    steps = [engine, 'bibtex', engine, engine]
    latex_env = _latex_subprocess_env(output_dir, engine)
    for i, step in enumerate(steps):
        print(f'  [{i+1}/{len(steps)}] {step}...')
        try:
            if step == 'bibtex':
                subprocess.run(
                    ['bibtex', stem],
                    capture_output=True,
                    cwd=str(output_dir),
                    timeout=timeout,
                )
                _sanitize_bbl_urls(output_dir / f'{stem}.bbl')
                _warn_bbl_long_urls(output_dir / f'{stem}.bbl')
            else:
                subprocess.run(
                    [engine, '-interaction=nonstopmode', tex_name],
                    capture_output=True,
                    cwd=str(output_dir),
                    env=latex_env,
                    timeout=timeout,
                )
        except subprocess.TimeoutExpired:
            print(f'  [compile-timeout] {step} 超时')
            return False
    return True


def compile_tex(tex_path, output_dir, cls_name=None):
    """按 cls 选择引擎编译（分步执行，每步打印进度）。

    引擎选择：copernicus 系列用 lualatex（cls 未适配 xelatex），其余用 xelatex。
    兜底：首选引擎失败且非 lualatex 时，自动用 lualatex 重试一次，覆盖未知 cls 的兼容问题。
    使用相对路径避免中文路径导致引擎崩溃。
    """
    tex_path = Path(tex_path)
    output_dir = Path(output_dir)
    tex_name = tex_path.name
    stem = tex_path.stem
    primary_engine = _select_engine(cls_name)
    try:
        bib_check = validate_bibliography_keys(tex_path, output_dir)
        if not bib_check.get('ok'):
            missing = bib_check.get('missing') or []
            preview = ', '.join(missing[:12])
            if len(missing) > 12:
                preview += f', ... (+{len(missing) - 12})'
            reason = bib_check.get('reason') or 'unknown'
            print(f'  [bib-check] failed: {reason}; missing keys: {preview}')
            return None

        def _attempt(engine):
            _remove_stale_pdf(tex_path)
            if _run_latex_steps(tex_name, stem, output_dir, engine=engine):
                pdf_path = tex_path.with_suffix('.pdf')
                if pdf_path.exists():
                    return str(pdf_path)
            if _fallback_strip_fullwidth_to_floats(tex_path):
                _remove_stale_pdf(tex_path)
                if _run_latex_steps(tex_name, stem, output_dir, engine=engine):
                    pdf_path = tex_path.with_suffix('.pdf')
                    if pdf_path.exists():
                        return str(pdf_path)
            return None

        pdf = _attempt(primary_engine)
        if pdf:
            return pdf
        # 兜底：首选引擎失败时换 lualatex 重试（覆盖未知 cls 的 xelatex 兼容问题）
        if primary_engine != 'lualatex':
            print(f'  [engine-fallback] {primary_engine} 失败，尝试 lualatex 重试...')
            pdf = _attempt('lualatex')
            if pdf:
                return pdf
        return None
    except Exception as e:
        print(f'编译异常: {e}')
        return None
