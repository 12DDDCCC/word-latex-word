"""OMML (Office Math Markup Language) → LaTeX 转换器

从 Word (.docx) 文档中提取公式并转换为 LaTeX 格式。
支持：上下标、分数、N元运算(∑∏∫)、横线、定界符、根号、重音、矩阵、方程组。
"""

from docx import Document
from docx.oxml.ns import qn
from pathlib import Path


# Unicode 希腊字母 → LaTeX 命令
GREEK_MAP = {
    'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma',
    'δ': '\\delta', 'ε': '\\epsilon', 'ζ': '\\zeta',
    'η': '\\eta', 'θ': '\\theta', 'ι': '\\iota',
    'κ': '\\kappa', 'λ': '\\lambda', 'μ': '\\mu',
    'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi',
    'ρ': '\\rho', 'ς': '\\varsigma', 'σ': '\\sigma',
    'τ': '\\tau', 'υ': '\\upsilon', 'φ': '\\phi',
    'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
    'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta',
    'Λ': '\\Lambda', 'Ξ': '\\Xi', 'Π': '\\Pi',
    'Σ': '\\Sigma', 'Φ': '\\Phi', 'Ψ': '\\Psi',
    'Ω': '\\Omega',
}

NARY_MAP = {
    '∑': '\\sum', '∏': '\\prod', '∫': '\\int',
    '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint',
    '⋀': '\\bigwedge', '⋁': '\\bigvee',
    '⋂': '\\bigcap', '⋃': '\\bigcup',
}

ACCENT_MAP = {
    '̂': '\\hat', '̃': '\\tilde', '̄': '\\bar',
    '̇': '\\dot', '̈': '\\ddot', '⃗': '\\vec',
    '̀': '\\grave', '́': '\\acute', '̑': '\\breve',
}

DELIM_MAP = {'{': '\\{', '}': '\\}', '|': '\\|'}

# 公式内特殊符号 → LaTeX
FORMULA_SYMBOL_MAP = {
    '×': '\\times', '°': '\\circ',
    '≤': '\\leq', '≥': '\\geq', '±': '\\pm',
    '≈': '\\approx', '≠': '\\neq', '∞': '\\infty',
    '∝': '\\propto', '·': '\\cdot',
}


def _escape_math_run_text(text):
    """Keep math syntax from OMML runs while escaping hard LaTeX breaks."""
    return (text or '').replace('&', r'\&').replace('%', r'\%').replace('#', r'\#')


def omml_to_latex(omath_elem):
    """将单个 m:oMath 元素转换为 LaTeX 字符串"""
    parts = []
    for child in omath_elem:
        result = _convert_element(child)
        if result:
            parts.append(result)
    return ''.join(parts)


def _convert_element(elem):
    """递归转换单个 OMML 元素"""
    tag = elem.tag.split('}')[-1]

    converters = {
        'r': _convert_run,
        'f': _convert_fraction,
        'sSup': _convert_superscript,
        'sSub': _convert_subscript,
        'sSubSup': _convert_subsuperscript,
        'nary': _convert_nary,
        'bar': _convert_bar,
        'd': _convert_delimiter,
        'rad': _convert_radical,
        'acc': _convert_accent,
        'm': _convert_matrix,
        'eqArr': _convert_eqarray,
        'sPre': _convert_pre_sub_sup,
        'groupChr': _convert_groupchr,
    }

    converter = converters.get(tag)
    if converter:
        return converter(elem)

    # 容器/属性元素：递归子元素
    if tag in ('oMath', 'oMathPara', 'e', 'num', 'den', 'sub', 'sup',
               'oMathParaPr', 'sSubSupPr', 'sSubPr', 'sSupPr', 'sPrePr',
               'fPr', 'naryPr', 'barPr', 'dPr', 'radPr', 'accPr',
               'mPr', 'eqArrPr', 'rPr', 'ctrlPr', 'groupChrPr'):
        return _convert_children(elem)

    return ''


def _get_text(elem):
    """提取元素中的文本"""
    return ''.join(t.text for t in elem.findall('.//' + qn('m:t')) if t.text)


def _convert_children(elem):
    """递归转换子元素"""
    if elem is None:
        return ''
    return ''.join(_convert_element(c) for c in elem)


def _convert_run(elem):
    """转换 m:r"""
    rPr = elem.find(qn('m:rPr'))
    style = None
    if rPr is not None:
        sty = rPr.find(qn('m:sty'))
        if sty is not None:
            style = sty.get(qn('m:val'))

    text = _get_text(elem)
    if not text:
        return ''

    for uc, latex in GREEK_MAP.items():
        text = text.replace(uc, latex)
    for uc, latex in FORMULA_SYMBOL_MAP.items():
        text = text.replace(uc, latex)
    text = _escape_math_run_text(text)

    if style == 'p':
        has_letter = any(c.isalpha() and c not in GREEK_MAP for c in text)
        if has_letter:
            return f'\\mathrm{{{text}}}'
        return text
    elif style == 'b':
        return f'\\mathbf{{{text}}}'
    elif style == 'bi':
        return f'\\boldsymbol{{{text}}}'
    elif style == 'i':
        return f'\\mathit{{{text}}}'
    return text


def _convert_fraction(elem):
    """m:f → \\frac{num}{den}"""
    num = _convert_children(elem.find(qn('m:num')))
    den = _convert_children(elem.find(qn('m:den')))
    return f'\\frac{{{num}}}{{{den}}}'


def _convert_superscript(elem):
    """m:sSup → base^{sup}"""
    base = _convert_children(elem.find(qn('m:e')))
    sup = _convert_children(elem.find(qn('m:sup')))
    return f'{base}^{{{sup}}}' if sup else base


def _convert_subscript(elem):
    """m:sSub → base_{sub}"""
    base = _convert_children(elem.find(qn('m:e')))
    sub = _convert_children(elem.find(qn('m:sub')))
    return f'{base}_{{{sub}}}' if sub else base


def _convert_subsuperscript(elem):
    """m:sSubSup → base_{sub}^{sup}"""
    base = _convert_children(elem.find(qn('m:e')))
    sub = _convert_children(elem.find(qn('m:sub')))
    sup = _convert_children(elem.find(qn('m:sup')))
    r = base
    if sub:
        r += f'_{{{sub}}}'
    if sup:
        r += f'^{{{sup}}}'
    return r


def _convert_nary(elem):
    """m:nary → \\sum_{sub}^{sup} expr"""
    naryPr = elem.find(qn('m:naryPr'))
    chr_val = '∑'
    if naryPr is not None:
        c = naryPr.find(qn('m:chr'))
        if c is not None:
            chr_val = c.get(qn('m:val'), '∑')

    op = NARY_MAP.get(chr_val, chr_val)
    sub_t = _convert_children(elem.find(qn('m:sub')))
    sup_t = _convert_children(elem.find(qn('m:sup')))
    expr = _convert_children(elem.find(qn('m:e')))

    r = op
    if sub_t:
        r += f'_{{{sub_t}}}'
    if sup_t:
        r += f'^{{{sup_t}}}'
    return r + expr


def _convert_bar(elem):
    """m:bar → \\overline{} 或 \\underline{}"""
    barPr = elem.find(qn('m:barPr'))
    pos = 'top'
    if barPr is not None:
        p = barPr.find(qn('m:pos'))
        if p is not None:
            pos = p.get(qn('m:val'), 'top')

    base = _convert_children(elem.find(qn('m:e')))
    return f'\\overline{{{base}}}' if pos == 'top' else f'\\underline{{{base}}}'


def _convert_delimiter(elem):
    """m:d → \\left(\\right)"""
    dPr = elem.find(qn('m:dPr'))
    beg, end = '(', ')'
    if dPr is not None:
        b = dPr.find(qn('m:begChr'))
        e = dPr.find(qn('m:endChr'))
        if b is not None:
            beg = b.get(qn('m:val'), '(') or '.'
        if e is not None:
            end = e.get(qn('m:val'), ')') or '.'

    beg_l = DELIM_MAP.get(beg, beg)
    end_l = DELIM_MAP.get(end, end)
    parts = [_convert_children(e) for e in elem.findall(qn('m:e'))]
    return f'\\left{beg_l}{" ".join(parts)}\\right{end_l}'


def _convert_radical(elem):
    """m:rad → \\sqrt{} 或 \\sqrt[n]{}"""
    deg = _convert_children(elem.find(qn('m:deg')))
    base = _convert_children(elem.find(qn('m:e')))
    if deg:
        return f'\\sqrt[{deg}]{{{base}}}'
    return f'\\sqrt{{{base}}}'


def _convert_accent(elem):
    """m:acc → \\hat{} 等"""
    accPr = elem.find(qn('m:accPr'))
    chr_val = '̂'
    if accPr is not None:
        c = accPr.find(qn('m:chr'))
        if c is not None:
            chr_val = c.get(qn('m:val'), '̂')

    base = _convert_children(elem.find(qn('m:e')))
    accent = ACCENT_MAP.get(chr_val)
    if not accent:
        return f'\\overset{{{chr_val}}}{{{base}}}' if chr_val else base
    return f'{accent}{{{base}}}'


def _convert_matrix(elem):
    """m:m → \\begin{matrix}...\\end{matrix}"""
    rows = []
    for mr in elem.findall(qn('m:mr')):
        cells = [_convert_children(e) for e in mr.findall(qn('m:e'))]
        rows.append(' & '.join(cells))
    return f'\\begin{{matrix}}{" \\\\ ".join(rows)}\\end{{matrix}}'


def _convert_eqarray(elem):
    """m:eqArr → \\begin{aligned}...\\end{aligned}"""
    rows = [_convert_children(e) for e in elem.findall(qn('m:e'))]
    return f'\\begin{{aligned}}{" \\\\ ".join(rows)}\\end{{aligned}}'


def _convert_pre_sub_sup(elem):
    """m:sPre → {}_{sub}{}^{sup}base"""
    base = _convert_children(elem.find(qn('m:e')))
    sub = _convert_children(elem.find(qn('m:sub')))
    sup = _convert_children(elem.find(qn('m:sup')))
    r = ''
    if sub:
        r += f'_{{{sub}}}'
    if sup:
        r += f'^{{{sup}}}'
    return r + base


def _convert_groupchr(elem):
    """m:groupChr → \\underbrace{} 等"""
    groupChrPr = elem.find(qn('m:groupChrPr'))
    chr_val = '⏟'
    if groupChrPr is not None:
        c = groupChrPr.find(qn('m:chr'))
        if c is not None:
            chr_val = c.get(qn('m:val'), '⏟')

    group_map = {'⏟': '\\underbrace', '⏞': '\\overbrace', '⏜': '\\widehat', '⏝': '\\widetilde'}
    cmd = group_map.get(chr_val, '\\underbrace')
    base = _convert_children(elem.find(qn('m:e')))
    return f'{cmd}{{{base}}}'


def _analyze_formula_positions(para, pi):
    """分析段落中公式的精确位置

    返回公式在段落内的偏移信息：
    - before_text: 公式前的文本
    - after_text: 公式后的文本
    - child_index: 公式在段落子节点中的序号
    - total_children: 段落子节点总数
    - type: display 或 inline
    """
    elem = para._element
    results = []

    # 收集段落直接子节点信息
    children = list(elem)
    total = len(children)

    # 计算每个公式前的文本长度偏移
    text_offset = 0
    for ci, child in enumerate(children):
        tag = child.tag.split('}')[-1]

        if tag == 'oMathPara':
            # display 公式
            for om in child.findall(qn('m:oMath')):
                latex = omml_to_latex(om)
                if latex:
                    results.append({
                        'para_index': pi,
                        'child_index': ci,
                        'total_children': total,
                        'type': 'display',
                        'before_text': para.text[:text_offset] if text_offset <= len(para.text) else '',
                        'after_text': '',
                        'latex': latex,
                    })
            # 更新偏移（oMathPara 不产生可见文本偏移）
            continue

        if tag == 'oMath':
            # inline 公式
            latex = omml_to_latex(child)
            if latex:
                results.append({
                    'para_index': pi,
                    'child_index': ci,
                    'total_children': total,
                    'type': 'inline',
                    'before_text': para.text[:text_offset] if text_offset <= len(para.text) else '',
                    'after_text': '',
                    'latex': latex,
                })

        # 累加文本偏移
        if tag == 'r':
            t = child.find(qn('w:t'))
            if t is not None and t.text:
                text_offset += len(t.text)

    # 回填 after_text：每个公式的 after_text = 下一个公式前的 before_text 减去当前 before_text
    for i in range(len(results)):
        if i + 1 < len(results) and results[i]['para_index'] == results[i+1]['para_index']:
            before_next = results[i+1]['before_text']
            before_cur = results[i]['before_text']
            results[i]['after_text'] = para.text[len(before_cur):len(before_next)] if len(para.text) >= len(before_next) else ''
        else:
            results[i]['after_text'] = para.text[len(results[i]['before_text']):] if results[i]['before_text'] else para.text

    return results


def extract_formulas(docx_path):
    """从 Word 文档提取所有公式及其精确位置

    Returns:
        list: [{
            'para_index': 段落序号,
            'child_index': 公式在段落子节点中的序号,
            'total_children': 段落子节点总数,
            'type': 'display' 或 'inline',
            'before_text': 公式前的文本,
            'after_text': 公式后的文本,
            'latex': LaTeX 代码,
            'context': 段落前30字符,
        }]
    """
    doc = Document(docx_path)
    results = []

    for pi, para in enumerate(doc.paragraphs):
        pos_list = _analyze_formula_positions(para, pi)
        for p in pos_list:
            p['context'] = para.text[:30] if para.text else ''
            results.append(p)

    return results


def generate_latex_doc(formulas, output_path):
    """生成完整的 LaTeX 文件，编译显示所有公式

    Args:
        formulas: extract_formulas() 的返回结果
        output_path: 输出 .tex 文件路径
    """
    lines = [
        r'\documentclass{article}',
        r'\usepackage[UTF8,fontset=windows]{ctex}',
        r'\usepackage{amsmath,amssymb}',
        r'\usepackage{geometry}',
        r'\geometry{a4paper,margin=2cm}',
        r'\begin{document}',
        '',
    ]

    eq_count = 0
    for f in formulas:
        eq_count += 1
        latex = f['latex']
        ftype = f['type']

        if ftype == 'display':
            lines.append(f'% 公式 {eq_count} (段落{f["para_index"]} 子节点{f["child_index"]}/{f["total_children"]})')
            lines.append(r'\begin{equation}')
            lines.append(f'  {latex}')
            lines.append(r'\end{equation}')
        else:
            lines.append(f'% 公式 {eq_count} (段落{f["para_index"]} 子节点{f["child_index"]}/{f["total_children"]}, 行内)')
            before = f.get('before_text', '')[-20:]
            after = f.get('after_text', '')[:20]
            lines.append(f'行内公式 {eq_count}: ...{before}$ {latex} ${after}...')
        lines.append('')

    lines.append(r'\end{document}')

    Path(output_path).write_text('\n'.join(lines), encoding='utf-8')
    return output_path


if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("用法: python omml_to_latex.py <docx路径> [输出目录]")
        print("提取 Word 文档中的公式，生成 LaTeX 文件编译显示")
        sys.exit(1)

    docx_path = sys.argv[1]
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path('.')

    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. 提取公式
    formulas = extract_formulas(docx_path)
    print(f"提取到 {len(formulas)} 个公式")

    # 2. 保存公式列表（含位置信息）
    list_path = out_dir / 'formulas_list.txt'
    with open(list_path, 'w', encoding='utf-8') as f:
        for i, fm in enumerate(formulas):
            ftype = '独立' if fm['type'] == 'display' else '行内'
            f.write(f"[{i+1}] 段落{fm['para_index']} 子节点{fm['child_index']}/{fm['total_children']} ({ftype})\n")
            f.write(f"    LaTeX: {fm['latex']}\n")
            if fm['type'] == 'inline':
                f.write(f"    前文: ...{fm['before_text'][-30:]}\n")
                f.write(f"    后文: {fm['after_text'][:30]}...\n")
            f.write('\n')
    print(f"公式列表: {list_path}")

    # 3. 生成 LaTeX 文件
    tex_path = out_dir / 'formulas.tex'
    generate_latex_doc(formulas, tex_path)
    print(f"LaTeX 文件: {tex_path}")

    # 4. 编译
    import subprocess
    try:
        result = subprocess.run(
            ['xelatex', '-interaction=nonstopmode', str(tex_path)],
            capture_output=True, text=True, encoding='utf-8',
            errors='replace', cwd=str(out_dir), timeout=60
        )
    except subprocess.TimeoutExpired:
        print(f"编译超时（>60s），请检查 {tex_path}")
        sys.exit(1)
    pdf_path = out_dir / 'formulas.pdf'
    if pdf_path.exists():
        print(f"PDF 输出: {pdf_path}")
    else:
        print(f"编译失败，请检查 {tex_path}")
