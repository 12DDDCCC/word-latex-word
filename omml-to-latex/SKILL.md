---
name: omml-to-latex
description: OMML ↔ LaTeX 双向转换：从 Word 提取公式转 LaTeX，从 LaTeX 公式转 Word OMML
version: 3.0.0
author: 小孟同学
tags: [docx, formula, omml, latex, math, python, position, symbol, latex2mathml, xslt]
---

# OMML ↔ LaTeX 双向转换 Skill

**提取方向**: Word (.docx) → OMML → LaTeX（含精确位置信息）
**转换方向**: LaTeX → MathML → OMML → Word（无损公式插入）

## 文件结构

```
skill/omml-to-latex/
├── SKILL.md           ← 本文档
├── omml_to_latex.py   ← Word → LaTeX 提取器
└── latex_to_omml.py   ← LaTeX → Word 转换器
```

---

## 方向一：Word → LaTeX 提取（omml_to_latex.py）

从 Word (.docx) 提取公式，转换为 LaTeX，**并记录每个公式在原文中的精确位置**。

### 核心功能

1. **公式提取**: 从 Word OMML 递归解析 14 种元素 → LaTeX
2. **位置定位**: 记录公式在段落中的子节点序号、前后文本
3. **编译验证**: 自动生成 LaTeX 文件并 xelatex 编译为 PDF

### 使用方法

```bash
python omml_to_latex.py 小论文.docx test/formula
```

### Python API

```python
from omml_to_latex import extract_formulas, generate_latex_doc

formulas = extract_formulas('小论文.docx')
for f in formulas:
    if f['type'] == 'inline':
        line = f['before_text'] + '$ ' + f['latex'] + ' $' + f['after_text']
    else:
        line = f'\\begin{{equation}}\n  {f["latex"]}\n\\end{{equation}}'
```

### 位置信息

| 字段 | 类型 | 说明 |
|------|------|------|
| `para_index` | int | 段落序号 |
| `child_index` | int | 子节点序号 |
| `total_children` | int | 子节点总数 |
| `type` | str | `display`/`inline` |
| `before_text` | str | 公式前的文本 |
| `after_text` | str | 公式后的文本 |
| `latex` | str | LaTeX 代码 |
| `context` | str | 段落前30字符 |

### OMML → LaTeX 映射

| OMML | LaTeX |
|------|-------|
| `m:f` | `\frac{num}{den}` |
| `m:sSup` | `base^{sup}` |
| `m:sSub` | `base_{sub}` |
| `m:sSubSup` | `base_{sub}^{sup}` |
| `m:nary` | `\sum_{sub}^{sup}` |
| `m:bar` | `\overline{e}` |
| `m:d` | `\left(...\right)` |
| `m:acc` | `\hat{x}` |
| `m:rad` | `\sqrt{e}` |
| `m:m` | `\begin{matrix}` |
| `m:eqArr` | `\begin{aligned}` |
| `m:sPre` | `_{sub}^{sup}base` |
| `m:groupChr` | `\underbrace{e}` |

### Run 属性

| `m:sty` | LaTeX |
|----------|-------|
| `p` | `\mathrm{}` (字母) 或 直出 (数字) |
| `b` | `\mathbf{}` |
| `bi` | `\boldsymbol{}` |
| `i` | `\mathit{}` |

---

## 方向二：LaTeX → Word 转换（latex_to_omml.py）

将 LaTeX 公式无损转换为 Word OMML，可直接插入 .docx 文档。

### 转换路径

```
LaTeX → latex2mathml → MathML → MML2OMML.XSL → OMML → python-docx 插入
```

### 核心功能

1. **公式转换**: `latex_to_omml()` — LaTeX → OMML XML 字符串
2. **元素生成**: `latex_to_omml_element()` — LaTeX → lxml Element（直接插入 docx）
3. **公式提取**: `extract_formulas_from_tex()` — 从 .tex 文件提取所有公式
4. **环境预处理**: `gather_to_display()` / `equation_to_display()` — 公式环境拆分
5. **文本清理**: `clean_latex_text()` — LaTeX 标记 → Unicode 纯文本
6. **批量转换**: `formulas_to_omml_list()` — 提取 + 转换一步完成

### 使用方法

```bash
# 提取 .tex 文件中所有公式并转为 OMML
python latex_to_omml.py paper.tex

# 直接转换单条公式
python latex_to_omml.py --convert '\frac{1}{2}'
```

### Python API

```python
from latex_to_omml import latex_to_omml, extract_formulas_from_tex, clean_latex_text

# 1. 单条公式转换
omml_xml = latex_to_omml(r'\frac{1}{n-1}')

# 2. 从 .tex 文件提取所有公式
formulas = extract_formulas_from_tex('paper.tex')
for f in formulas:
    print(f"  {f['type']} {f['env']} eq={f['eq_num']}")
    print(f"  LaTeX: {f['latex'][:80]}")

# 3. 清理 LaTeX 文本（用于 Word caption 显示）
clean = clean_latex_text(r'CO$_2$ concentration')  # → 'CO₂ concentration'
```

### 插入 Word 文档示例

```python
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from latex_to_omml import latex_to_omml

doc = Document('output.docx')
para = doc.add_paragraph()

# 转换公式
omml_str = latex_to_omml(r'\sum_{i=1}^{n} x_i^2')
if omml_str:
    omml_elem = etree.fromstring(omml_str.encode('utf-8'))
    para._element.append(omml_elem)
```

### gather 环境拆分规则

gather 环境中多个公式以 `\\` 分隔，`\label{eqN}` 在公式下一行。

**关键规则**: label 归属于**上一行**公式（不是下一行）。

```latex
\begin{gather}
  X_0^b + \lambda \times \delta_i \times X_0^b \\
  \label{eq1}
  P^b = \frac{1}{n-1} \sum (X_i^b - \hat{X}^b)(X_i^b - \hat{X}^b)^T \\
  \label{eq2}
  \overline{X^a} = \overline{X^b} + K(y - H\overline{X^b})
  \label{eq3}
\end{gather}
```

拆分结果:
- 公式1: `X_0^b + ...` 编号 (1)
- 公式2: `P^b = ...` 编号 (2)
- 公式3: `\overline{X^a} = ...` 编号 (3)

### LaTeX 文本清理

| 输入 | 输出 | 说明 |
|------|------|------|
| `CO$_2$` | `CO₂` | 常见化学式下标 |
| `XCO$_2$` | `XCO₂` | 柱浓度下标 |
| `$^o$` | `ᵒ` | 上标度数 |
| `\textbf{word}` | `word` | 去粗体命令 |
| `\textit{word}` | `word` | 去斜体命令 |

### 公式提取返回格式

```python
[{
    'id': 0,           # 序号
    'type': 'display',  # display 或 inline
    'env': 'equation',  # equation/gather_line/bracket/dollar
    'latex': '...',     # LaTeX 公式代码
    'label': 'eq1',     # label 键名 或 None
    'eq_num': '(1)',    # 公式编号 或 None
    'start': 1234,      # 文件中起始位置
    'end': 1300,        # 文件中结束位置
}]
```

---

## 依赖

### omml_to_latex.py（提取方向）

```bash
pip install python-docx
```

### latex_to_omml.py（转换方向）

```bash
pip install latex2mathml lxml
```

**XSLT 文件**: 需要 Microsoft Office 的 `MML2OMML.XSL`，自动查找路径:
- `C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL`
- `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`

---

## 注意事项

| 问题 | 说明 |
|------|------|
| `\boldsymbol{X}` | latex2mathml 不支持，自动转为 `\mathbf{X}` |
| `\textbf{ }` | 空粗体命令自动移除 |
| `sty="p"` | 下标数字默认直体，字母需 `\mathrm{}` |
| 希腊字母 | OMML 用 Unicode (λ=U+03BB)，映射到 `\lambda` |
| 公式编号 | Word 中编号是段落文本，不是 OMML 属性 |
| gather label | label 在公式**下一行**，归属**上一行**公式 |
| Python `\\` | 匹配 LaTeX 换行符需用 `'\\' + '\\'`，避免转义混淆 |
