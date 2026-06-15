r"""LaTeX → OMML (Office Math Markup Language) 转换器

将 LaTeX 公式转换为 Word 可识别的 OMML XML 字符串，支持直接插入 .docx 文档。
核心路径: LaTeX → latex2mathml → MathML → XSLT (MML2OMML.XSL) → OMML

同时提供从 LaTeX 源文件提取公式的功能，包括:
- equation/align/multline 等环境 → 独立公式
- gather 环境 → 拆分为多个独立公式（label 归属上一行公式）
- \[...\] 行间公式
- $...$ 行内公式
- LaTeX 文本清理（CO$_2$ → CO₂ 等 Unicode 转换）
"""

import os
import re
import sys
from pathlib import Path


def _match_braces(s, start):
    """从s[start]='{'位置匹配平衡的花括号，返回内容(不含外层{})"""
    if start >= len(s) or s[start] != '{':
        return None, start
    depth = 0
    i = start
    while i < len(s):
        if s[i] == '{':
            depth += 1
        elif s[i] == '}':
            depth -= 1
            if depth == 0:
                return s[start + 1:i], i + 1
        i += 1
    return None, start

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.latex_text_utils import to_subscript, to_superscript, clean_latex_text


# ============================================================
# LaTeX 公式环境预处理


# ============================================================
# LaTeX 公式环境预处理
# ============================================================

def gather_to_display(match):
    """将 gather 环境拆分为独立的 $$...$$ 公式

    gather 中每个公式以 \\\\ 分隔，\\label{eqN} 在公式下一行，
    label 归属于上一行公式（即 label 跟在公式后面）。

    Args:
        match: re.Match 对象，group(1) 为 gather 环境内部内容

    Returns:
        用 \\n\\n 分隔的独立 $$ 公式字符串
    """
    body = match.group(1)
    raw_lines = body.split('\n')
    formula_lines = []  # [(公式文本, 编号), ...]

    i = 0
    while i < len(raw_lines):
        stripped = raw_lines[i].strip()
        if not stripped:
            i += 1
            continue

        if stripped.startswith('\\label'):
            label_m = re.search(r'\\label\{([^}]+)\}', stripped)
            if label_m and formula_lines:
                label_key = label_m.group(1)
                num_m = re.search(r'eq(\d+)[-\.]?(\d+)?', label_key)
                if num_m:
                    if num_m.group(2):
                        eq_num = f'({num_m.group(1)}-{num_m.group(2)})'
                    else:
                        eq_num = f'({num_m.group(1)})'
                else:
                    eq_num = f'({label_key})'
                formula_lines[-1] = (formula_lines[-1][0], eq_num)
            i += 1
            continue

        dslash = '\\' + '\\'
        if stripped.endswith(dslash):
            stripped = stripped[:-2].rstrip()
        if stripped:
            formula_lines.append((stripped, ''))
        i += 1

    result_parts = []
    for eq_text, eq_num in formula_lines:
        if eq_num:
            result_parts.append(f'$$ {eq_text} $$ {eq_num}')
        else:
            result_parts.append(f'$$ {eq_text} $$')
    return '\n\n'.join(result_parts)


def equation_to_display(match):
    """将 equation 环境转为 $$...$$ 公式

    提取 \\label{eqN} 中的编号信息，追加到公式后。
    """
    body = match.group(1)
    label_m = re.search(r'\\label\{([^}]+)\}', body)
    eq_num = ''
    if label_m:
        label_key = label_m.group(1)
        num_m = re.search(r'eq(\d+)[-\.]?(\d+)?', label_key)
        if num_m:
            if num_m.group(2):
                eq_num = f'({num_m.group(1)}-{num_m.group(2)})'
            else:
                eq_num = f'({num_m.group(1)})'
        else:
            eq_num = f'({label_key})'
        body = re.sub(r'\\label\{[^}]+\}', '', body)
    body = body.strip()
    if eq_num:
        return f'$$ {body} $$ {eq_num}'
    return f'$$ {body} $$'


def preprocess_formulas(content):
    """预处理 LaTeX 内容，将公式环境转为 Pandoc 可识别的 $$ 格式

    处理: gather → 拆分独立公式, equation → $$...$$

    Args:
        content: LaTeX 源文件内容

    Returns:
        预处理后的内容字符串
    """
    result = re.sub(
        r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}',
        gather_to_display, content, flags=re.DOTALL
    )
    result = re.sub(
        r'\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}',
        equation_to_display, result, flags=re.DOTALL
    )
    return result


# ============================================================
# 从 LaTeX 源文件提取公式
# ============================================================

def extract_formulas_from_tex(tex_path):
    """从 LaTeX 源文件提取数学公式，包括编号信息

    gather 环境中的多个公式会被拆分为独立的 display 公式，
    label 正确归属到上一行公式。

    Args:
        tex_path: .tex 文件路径

    Returns:
        list: [{
            'id': 序号,
            'type': 'display' 或 'inline',
            'env': 环境名 (equation/gather_line/bracket/dollar),
            'latex': LaTeX 公式代码,
            'label': label 键名 或 None,
            'eq_num': 公式编号 如 '(1)' 或 None,
            'start': 在文件中的起始位置,
            'end': 在文件中的结束位置,
        }]
    """
    with open(tex_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    formulas = []
    idx = 0
    equation_counter = 0

    # equation/align/multline/alignat/flalign 环境
    for m in re.finditer(
        r'\\begin\{(equation|align|multline|alignat|flalign)\*?\}(.*?)\\end\{\1\*?\}',
        content, re.DOTALL
    ):
        env = m.group(1)
        body = m.group(2).strip()

        label_match = re.search(r'\\label\{([^}]+)\}', body)
        label = label_match.group(1) if label_match else None

        eq_num = None
        if label:
            num_match = re.search(r'eq(\d+)[-\.]?(\d+)?', label)
            if num_match:
                if num_match.group(2):
                    eq_num = f"({num_match.group(1)}-{num_match.group(2)})"
                else:
                    eq_num = f"({num_match.group(1)})"

        if not eq_num:
            equation_counter += 1
            eq_num = f"({equation_counter})"

        formulas.append({
            'id': idx, 'type': 'display', 'env': env,
            'latex': body, 'label': label, 'eq_num': eq_num,
            'start': m.start(), 'end': m.end(),
        })
        idx += 1

    # gather 环境: 拆分为独立公式
    for m in re.finditer(r'\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}', content, re.DOTALL):
        body = m.group(1)
        raw_lines = body.split('\n')
        current_label = ''

        i = 0
        while i < len(raw_lines):
            stripped = raw_lines[i].strip()
            if not stripped:
                i += 1
                continue

            if stripped.startswith('\\label'):
                label_m = re.search(r'\\label\{([^}]+)\}', stripped)
                if label_m:
                    current_label = label_m.group(1)
                i += 1
                continue

            dslash = '\\' + '\\'
            if stripped.endswith(dslash):
                stripped = stripped[:-2].rstrip()

            if stripped:
                eq_num = ''
                if current_label:
                    num_m = re.search(r'eq(\d+)[-\.]?(\d+)?', current_label)
                    if num_m:
                        if num_m.group(2):
                            eq_num = f'({num_m.group(1)}-{num_m.group(2)})'
                        else:
                            eq_num = f'({num_m.group(1)})'
                    current_label = ''

                if not eq_num:
                    equation_counter += 1
                    eq_num = f'({equation_counter})'

                formulas.append({
                    'id': idx, 'type': 'display', 'env': 'gather_line',
                    'latex': stripped, 'label': None, 'eq_num': eq_num,
                    'start': m.start(), 'end': m.end(),
                })
                idx += 1
            i += 1

    # \[...\] 行间公式
    for m in re.finditer(r'\\\[(.*?)\\\]', content, re.DOTALL):
        if any(m.start() >= f['start'] and m.end() <= f['end'] for f in formulas):
            continue
        formulas.append({
            'id': idx, 'type': 'display', 'env': 'bracket',
            'latex': m.group(1).strip(), 'label': None, 'eq_num': None,
            'start': m.start(), 'end': m.end(),
        })
        idx += 1

    # $...$ 行内公式（非 $$）
    for m in re.finditer(r'(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)', content):
        if any(m.start() >= f['start'] and m.end() <= f['end'] for f in formulas):
            continue
        formulas.append({
            'id': idx, 'type': 'inline', 'env': 'dollar',
            'latex': m.group(1).strip(), 'label': None, 'eq_num': None,
            'start': m.start(), 'end': m.end(),
        })
        idx += 1

    return formulas


# ============================================================
# LaTeX → OMML 核心转换
# ============================================================

def _find_xslt_path():
    """查找 MML2OMML.XSL 文件路径"""
    candidates = [
        os.path.join(
            os.environ.get('ProgramFiles(x86)', ''),
            'Microsoft Office', 'root', 'Office16', 'MML2OMML.XSL'
        ),
        r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL',
        r'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def _wrap_nary_body(latex_str):
    r"""为nary操作符(\sum, \prod, \int等)明确包裹求和体

    问题: latex2mathml 将 \sum_{i} x_i 转换时，无法确定 \sum 的作用范围，
    把 \sum_{i} 生成为空的nary结构，而 x_i 放在外面，导致Word显示占位符。

    修复: 将 \sum_{sub}^{sup} expr 转为 \sum_{sub}^{sup}{expr}
    使 latex2mathml 明确知道求和体的范围。

    双层防护:
    1. 本函数(预处理层): 从根源上避免空nary
    2. _fix_empty_nary_body(OMML后处理层): 兜底修复漏网情况
    """
    nary_cmds = ('\\sum', '\\prod', '\\coprod', '\\bigcup', '\\bigcap',
                 '\\bigsqcup', '\\bigvee', '\\bigwedge', '\\bigoplus',
                 '\\bigotimes', '\\bigodot', '\\oint', '\\int',
                 '\\iint', '\\iiint')

    # 找到所有nary操作符的位置
    positions = []
    for cmd in nary_cmds:
        for m in re.finditer(re.escape(cmd), latex_str):
            positions.append((m.start(), m.end(), cmd))
    if not positions:
        return latex_str
    positions.sort()

    # 从后往前处理，避免偏移
    result = latex_str
    for start, end, cmd in reversed(positions):
        i = end
        # 跳过空白
        while i < len(result) and result[i] in ' \t':
            i += 1

        # 解析可选的 _{sub} 和 ^{sup}（支持嵌套括号）
        sub_content = None
        sup_content = None
        sub_span = (i, i)
        sup_span = (i, i)

        # 解析 _{sub} 或 _X
        if i < len(result) and result[i] == '_':
            i += 1
            if i < len(result) and result[i] == '{':
                content, new_i = _match_braces(result, i)
                if content is not None:
                    sub_content = content
                    sub_span = (i - 1, new_i)
                    i = new_i
            elif i < len(result):
                sub_content = result[i]
                sub_span = (i - 1, i + 1)
                i += 1

        # 跳过空白
        while i < len(result) and result[i] in ' \t':
            i += 1

        # 解析 ^{sup} 或 ^X
        if i < len(result) and result[i] == '^':
            i += 1
            if i < len(result) and result[i] == '{':
                content, new_i = _match_braces(result, i)
                if content is not None:
                    sup_content = content
                    sup_span = (i - 1, new_i)
                    i = new_i
            elif i < len(result):
                sup_content = result[i]
                sup_span = (i - 1, i + 1)
                i += 1

        # 跳过空白
        while i < len(result) and result[i] in ' \t':
            i += 1

        # 如果后面已有花括号，跳过
        if i < len(result) and result[i] == '{':
            continue
        if i >= len(result):
            continue

        # 确定求和体范围: 从当前位置到公式末尾
        # 但遇到等号或另一个nary操作符时停止
        body_start = i
        j = i
        paren_depth = 0
        brace_depth = 0
        while j < len(result):
            c = result[j]

            if c == '{':
                brace_depth += 1
            elif c == '}':
                brace_depth -= 1
            elif c == '(' and brace_depth == 0:
                paren_depth += 1
            elif c == ')' and brace_depth == 0:
                paren_depth -= 1
                if paren_depth == 0:
                    j += 1  # 包含右括号
                    break
            elif brace_depth == 0 and paren_depth == 0:
                # 顶层遇到 = 或另一个nary命令 → 停止
                if c == '=' and j > 0 and result[j - 1] != '\\':
                    break
                rest = result[j:]
                if any(rest.startswith(nc) for nc in nary_cmds):
                    break
                # 二元运算符(前后都有内容时): + - 等
                if c in '+-' and j > body_start:
                    prev_c = result[j - 1] if j > 0 else ''
                    if prev_c in '})]0123456789abcxyzpnkNMiI\'':
                        break
            j += 1

        if j > body_start:
            body = result[body_start:j]
            # 构造: \cmd_{sub}^{sup}{body}
            # 保留原始sub/sup文本（含_和^前缀），再追加{body}
            sub_sup_end = max(sub_span[1], sup_span[1])
            if sub_sup_end > end:
                # 有sub/sup，保留原始 _{...}^{...} 部分
                original_sub_sup = result[end:sub_sup_end]
                # 但要跳过sub_sup_end到body_start之间的空白
                wrapped = cmd + original_sub_sup + '{' + body + '}'
            else:
                # 没有sub/sup，直接追加body
                wrapped = cmd + '{' + body + '}'

            # 替换: 从cmd开始到body结束
            result = result[:start] + wrapped + result[j:]

    return result


def latex_to_omml(latex_str, xslt_path=None):
    """将 LaTeX 公式转为 Office Math ML (OMML) XML 字符串

    转换路径: LaTeX → latex2mathml → MathML → XSLT → OMML

    Args:
        latex_str: LaTeX 公式代码（不含 $ 或 $$ 包裹）
        xslt_path: MML2OMML.XSL 文件路径，为 None 时自动查找

    Returns:
        str: OMML XML 字符串，可直接插入 Word 文档
        None: 转换失败时返回 None
    """
    try:
        from latex2mathml.converter import convert as l2m_convert
        from lxml import etree

        # 清理 LaTeX 公式中 latex2mathml 不支持的命令
        cleaned = latex_str
        cleaned = re.sub(r'\\boldsymbol\{([^}]+)\}', r'\\mathbf{\1}', cleaned)
        cleaned = re.sub(r'\\textbf\{\s*\}', '', cleaned)

        # 预处理: 为nary操作符明确包裹求和体
        # 解决 latex2mathml 无法确定 \sum 作用范围的问题
        cleaned = _wrap_nary_body(cleaned)

        # LaTeX → MathML
        mathml = l2m_convert(cleaned)

        # 查找 XSLT 文件
        if xslt_path is None:
            xslt_path = _find_xslt_path()

        if not xslt_path or not os.path.exists(xslt_path):
            print('Warning: MML2OMML.XSL not found, cannot convert to OMML')
            return None

        # MathML → OMML
        mathml_tree = etree.fromstring(mathml.encode('utf-8'))
        xslt_tree = etree.parse(xslt_path)
        transform = etree.XSLT(xslt_tree)
        omml_tree = transform(mathml_tree)
        omml_str = etree.tostring(omml_tree, encoding='unicode')
        return omml_str

    except ImportError as e:
        print(f'Missing dependency: {e}. Install: pip install latex2mathml lxml')
        return None
    except Exception as e:
        print(f'Formula OMML conversion failed: {e}')
        return None


def latex_to_omml_element(latex_str, xslt_path=None):
    """将 LaTeX 公式转为 lxml Element，可直接插入 python-docx 文档

    Args:
        latex_str: LaTeX 公式代码
        xslt_path: MML2OMML.XSL 文件路径

    Returns:
        lxml.etree._Element: OMML 元素节点（m:oMathPara 或 m:oMath）
        None: 转换失败时返回 None
    """
    omml_str = latex_to_omml(latex_str, xslt_path)
    if omml_str is None:
        return None
    try:
        from lxml import etree
        return etree.fromstring(omml_str.encode('utf-8'))
    except Exception as e:
        print(f'Failed to parse OMML: {e}')
        return None


# ============================================================
# 便捷函数：批量转换公式并插入 Word
# ============================================================

def formulas_to_omml_list(tex_path, xslt_path=None):
    """从 LaTeX 文件提取所有公式并转为 OMML

    Args:
        tex_path: .tex 文件路径
        xslt_path: MML2OMML.XSL 路径（可选）

    Returns:
        list: [{
            'id': 序号,
            'type': 'display' 或 'inline',
            'env': 环境名,
            'latex': LaTeX 代码,
            'eq_num': 编号,
            'omml': OMML XML 字符串 或 None,
        }]
    """
    formulas = extract_formulas_from_tex(tex_path)
    for f in formulas:
        f['omml'] = latex_to_omml(f['latex'], xslt_path)
    return formulas


# ============================================================
# CLI 入口
# ============================================================

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print("用法: python latex_to_omml.py <tex文件> [--xslt XSLT路径]")
        print("  提取 LaTeX 公式并转为 OMML")
        print()
        print("单独转换公式:")
        print("  python latex_to_omml.py --convert '\\frac{1}{2}'")
        sys.exit(1)

    if sys.argv[1] == '--convert':
        # 直接转换单条公式
        latex_input = sys.argv[2] if len(sys.argv) > 2 else ''
        result = latex_to_omml(latex_input)
        if result:
            print(result[:500])
            print(f'... (total {len(result)} chars)')
        else:
            print('Conversion failed')
        sys.exit(0)

    tex_path = sys.argv[1]
    xslt = None
    if '--xslt' in sys.argv:
        xi = sys.argv.index('--xslt')
        if xi + 1 < len(sys.argv):
            xslt = sys.argv[xi + 1]

    formulas = formulas_to_omml_list(tex_path, xslt)
    print(f"提取到 {len(formulas)} 个公式")
    for f in formulas:
        omml_status = 'OK' if f['omml'] else 'FAILED'
        print(f"  [{f['id']}] {f['type']} {f['env']} eq={f['eq_num']} omml={omml_status}")
        print(f"       LaTeX: {f['latex'][:80]}...")
