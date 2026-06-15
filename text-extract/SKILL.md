---
name: text-extract
description: 从 Word (.docx) 无损提取文本信息（段落/标题/格式/公式/引用/化学式），不含表格和图片，输出 JSON 结构
version: 2.0.0
author: 小孟同学
tags: [docx, text, extract, heading, format, formula, citation, chemistry, python, json]
---

# Word 文本无损提取 Skill

从 Word (.docx) 提取所有文本信息，保留完整的格式细节，用于后续 LaTeX 精确转换。

**不提取表格和图片**，公式随文本一起提取确保文本统一性。

## 核心功能

1. **段落提取**: 样式、标题级别、对齐、缩进、行距、段前段后间距
2. **Run 级别格式**: 粗体、斜体、下划线、删除线、上标、下标、字体、字号、颜色
3. **公式集成**: 行内/独立公式随文本一起提取，交替拼接 LaTeX
4. **引用检测**: 红色(EE0000)编号 → `\citep{refN}`
5. **超链接**: URL + 文本
6. **化学式下标**: `CO2` → `CO$_{2}$`，`XCO2` → `XCO$_{2}$`
7. **占位符机制**: 保护 Unicode/化学式转换结果不被 LaTeX 转义破坏

## 输出结构

```json
{
  "source": "文件路径",
  "total_paragraphs": 100,
  "headings": [{"para_index": 0, "level": 1, "text": "...", "style": "Heading 1", "latex": "..."}],
  "paragraphs": [{
    "para_index": 0,
    "style": "Normal",
    "heading_level": null,
    "alignment": "justify",
    "first_line_indent_pt": 24.0,
    "left_indent_pt": null,
    "line_spacing": 1.5,
    "space_before_pt": null,
    "space_after_pt": 6.0,
    "runs": [
      {"type": "text", "text": "CO2", "bold": false, "latex": "CO$_{2}$", "color_rgb": "000000", "is_cite": false},
      {"type": "formula", "latex": "X_{0}^{b}", "formula_type": "inline"},
      {"type": "hyperlink", "text": "链接", "url": "https://...", "latex": "\\url{...}"}
    ],
    "text": "纯文本",
    "latex": "LaTeX格式文本",
    "has_formula": true
  }],
  "statistics": {"headings": 12, "citations": 71, "formulas": 47, "bold_runs": 20, "italic_runs": 5}
}
```

## 段落级别字段

| 字段 | 类型 | 说明 | 用途 |
|------|------|------|------|
| `para_index` | int | 段落序号 | 定位段落 |
| `style` | str | Word 样式名 | 判断段落类型 |
| `heading_level` | int/null | 标题级别 1-4 | 生成 `\section{}` 等 |
| `alignment` | str/null | left/center/right/justify | `\begin{flushleft}` 等 |
| `first_line_indent_pt` | float/null | 首行缩进(pt) | `\indent` 或 `\parindent` |
| `left_indent_pt` | float/null | 左缩进(pt) | 列表项偏移 |
| `line_spacing` | float/null | 行距倍数 | `\linespread` |
| `space_before_pt` | float/null | 段前间距(pt) | `\vspace{}` |
| `space_after_pt` | float/null | 段后间距(pt) | `\vspace{}` |
| `text` | str | 纯文本 | 查找/对照 |
| `latex` | str | LaTeX 格式 | 直接用于 LaTeX 文件 |
| `has_formula` | bool | 是否含公式 | 快速筛选 |

## Run 级别字段

| 字段 | 类型 | 说明 | LaTeX 映射 |
|------|------|------|------------|
| `type` | str | text/formula/hyperlink | — |
| `text` | str | 原始文本 | — |
| `bold` | bool | 粗体 | `\textbf{}` |
| `italic` | bool | 斜体 | `\textit{}` |
| `underline` | bool | 下划线 | `\underline{}` |
| `strike` | bool | 删除线 | `\sout{}` |
| `superscript` | bool | 上标 | `\textsuperscript{}` |
| `subscript` | bool | 下标 | `\textsubscript{}` |
| `font_name` | str/null | 字体名 | 字体选择 |
| `size_pt` | float/null | 字号(pt) | 字号映射 |
| `color_rgb` | str/null | 颜色HEX | 角色判断 |
| `color_role` | str/null | cite/颜色命令 | 引用或着色 |
| `is_cite` | bool | 是否引用编号 | `\citep{}` |
| `latex` | str | LaTeX 格式文本 | 直接输出 |

## 颜色 → 角色映射

| 颜色 | RGB | 角色 | LaTeX 处理 |
|------|-----|------|------------|
| 红色 | EE0000 | cite | 数字→`\citep{N}` |
| 黑色 | 000000 | 默认 | 无特殊处理 |
| 亮红 | FF0000 | `\textcolor{red}` | 红色着色 |
| 蓝色 | 0000FF | `\textcolor{blue}` | 蓝色着色 |
| 绿色 | 008000 | `\textcolor{green}` | 绿色着色 |

## 引用检测逻辑

Word 中引用编号以红色(EE0000)标记，如 "[13]" 或 "13"：

1. 检测 `color_rgb == 'EE0000'`
2. 提取文本中的数字 `\d+`
3. 映射为 `\citep{1},\citep{13}` 等（数字key直接匹配bib文件）
4. 非数字字符（逗号等）丢弃

**注意**: 数字key需与 `references.bib` 中的 `@Article{1,` 等数字key一致。

## 化学式下标识别 (v2.0)

Word 中化学式常见两种写法：
- `CO₂` — Unicode 下标字符
- `CO2` — 纯 ASCII 文本

两者都需要转为 `CO$_{2}$`。

### 识别规则

```python
chem_pattern = re.compile(r'([A-Z][A-Za-z]*?)(\d{1,2})(?=[^0-9a-zA-Z_]|$)')
```

| 输入 | 匹配 | 输出 | 说明 |
|------|------|------|------|
| `CO2` | CO+2 | `CO$_{2}$` | 化学式 |
| `XCO2` | XCO+2 | `XCO$_{2}$` | 化学式 |
| `OCO2` | — | `OCO2` | 卫星名，排除 |
| `XCO2variable_name` | — | `XCO2variable_name` | 数字后跟字母，不匹配 |
| `GlobalGriddedDailyCO2EmissionsDataset` | — | `GlobalGriddedDailyCO2EmissionsDataset` | 数据集/变量标识符内部，不下标化 |
| `GlobalGriddedDailyCO₂EmissionsDataset` | — | `GlobalGriddedDailyCO2EmissionsDataset` | Unicode下标夹在英文标识符内部，还原为普通数字 |
| `H2O` | H+2 | `H$_{2}$`O | 氢元素 |

### 排除列表

```python
if prefix in ('OCO',):  # OCO-2/OCO-3 卫星
    continue
```

### 正向前瞻 vs 负向前瞻

- 旧版: `(?![0-9a-zA-Z\$_\-])` — 负向前瞻，排除特定字符
- 新版: `(?=[^0-9a-zA-Z_]|$)` — 正向前瞻，数字后必须是非字母数字下划线或行尾

新版更严格：`XCO2quality_flag` 中 `2` 后跟 `q`，不匹配 ✓
标识符保护：`CO2`/`XCO2` 前后如果仍属于同一个英文、数字或下划线 token，则视为数据集名、变量名、版本号或文件名，不做化学式下标；已经由 Unicode 下标或 Word 下标生成的 `CO$_{2}$`/`\textsubscript{2}` 也要在英文 token 内还原为 `CO2`。

## 占位符机制 (v2.0 核心)

### 问题背景

Unicode 转换（₂→`$_{2}$`）和 LaTeX 转义（`_`→`\_`）在同一文本上先后执行，导致：
- `CO₂` → `CO$_{2}$` → `CO$\_\{2\}$` (错误！)

### 解决方案

用 `\x00` 包裹的占位符保护已转换的片段：

```
步骤1: Unicode → 占位符
  CO₂ → CO\x00AA\x00  (映射: \x00AA\x00 → $_{2}$)

步骤2: 化学式 → 占位符
  CO2 → CO\x00BB\x00  (映射: \x00BB\x00 → $_{2}$)

步骤3: LaTeX 特殊字符转义（占位符不受影响）
  _ → \_, { → \{, } → \}

步骤4: 恢复占位符
  CO\x00AA\x00 → CO$_{2}$ ✓
  CO\x00BB\x00 → CO$_{2}$ ✓
```

### 占位符格式

```python
ph_chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
ph = f'\x00{c1}{c2}\x00'  # 如 \x00AA\x00, \x00AB\x00
```

**为什么不含数字？** 因为化学式正则会匹配 `CO` + 数字，如果占位符是 `\x00PH0\x00`，正则可能把 `PH0` 中的 `0` 当成化学式数字匹配。

## 公式处理

公式通过 `omml_to_latex` 模块转换，在段落遍历时与文本 run 交替拼接：

- `m:oMath` (行内公式) → `$ latex $`
- `m:oMathPara` (独立公式) → `\begin{equation} latex \end{equation}`

段落子节点遍历顺序: `w:r` → `m:oMath` → `w:r` → `m:oMathPara` → ...

最终 `paragraph['latex']` = 拼接所有 run 的 latex 字段。

### 公式内符号转换 (omml_to_latex.py v2.0)

```python
FORMULA_SYMBOL_MAP = {
    '×': '\\times', '°': '^\\circ',
    '≤': '\\leq', '≥': '\\geq', '±': '\\pm',
    '≈': '\\approx', '≠': '\\neq', '∞': '\\infty',
    '∝': '\\propto', '·': '\\cdot',
}
```

### 公式内下划线转义

公式文本中的 `_`（非 OMML 结构控制的下标）需转义为 `\_`，避免 LaTeX 解释为下标命令：
```python
if '_' in text:
    text = text.replace('_', '\\_')
```

例: `XCO_{2}\_quality\_flag` — `_{2}` 是下标，`\_quality\_flag` 是字面下划线

## Unicode → LaTeX 映射

### 文本模式 (加 `$...$`)

| Unicode | LaTeX | 示例 |
|---------|-------|------|
| ₂ ₃ ₄ ₁ | `$_{N}$` | CO₂ → CO$_{2}$ |
| ² ³ ⁰ | `$^{N}$` | m² → m$^{2}$ |
| λ δ α | `$\lambda$` 等 | λ → $\lambda$ |
| ° | `$^{\circ}$` | 30° → 30$^{\circ}$ |
| × ≤ ≥ ± ≈ | `$\times$` 等 | — |

### 公式模式 (不加 `$...$`)

| Unicode | LaTeX | 用途 |
|---------|-------|------|
| ₂ ₃ ₁ | `_{N}` | 公式内下标 |
| ⁰ ² ³ | `^{N}` | 公式内上标 |
| λ δ | `\lambda` `\delta` | 公式内希腊字母 |
| × ° | `\times` `^\circ` | 公式内运算符 |
| ≤ ≥ ± ≈ ≠ | 对应 LaTeX 命令 | 公式内关系符 |

## 尺寸转换

Word 使用 EMU (English Metric Units)：

- 1 pt = 12700 EMU
- `emu_to_pt(emu) = round(emu / 12700, 1)`

常见字号对应:

| Word 字号 | EMU | pt |
|-----------|-----|----|
| 小四 | 152400 | 12 |
| 五号 | 203200 | 16 |
| 三号 | 304800 | 24 |

## 使用方法

### 命令行

```bash
python text_extract.py 小论文.docx test/output
```

输出:
- `text_extract.json` — 完整 JSON 数据
- `text_extract_summary.txt` — 可读摘要

### Python API

```python
from text_extract import extract_docx_text, generate_summary

# 提取
result = extract_docx_text('小论文.docx', 'output.json')

# 查看摘要
print(generate_summary(result))

# 遍历段落
for p in result['paragraphs']:
    if p['heading_level']:
        print(f"H{p['heading_level']}: {p['text']}")
    elif p['has_formula']:
        print(f"公式段落: {p['latex'][:50]}")
```

## 依赖

- Python >= 3.10
- `python-docx` >= 0.8.11
- `omml_to_latex` (同级 skill 模块)

```bash
pip install python-docx
```

## 注意事项

| 问题 | 说明 |
|------|------|
| 表格段落跳过 | `para._element.getparent()` 为 `w:tc` 时跳过 |
| 引用 key 映射 | `refN` 需与 bib 文件 key 对齐，目前为自动编号 |
| 公式空白字符 | 部分公式末尾有空白 `\boldsymbol{   }`，需后处理清理 |
| LaTeX 特殊字符 | `\ & % # { } _ ~` 自动转义 |
| Unicode 上下文 | 文本模式加 `$...$`，公式模式不加，避免双重包裹 |
| 化学式排除 | OCO2/OCO3 是卫星名不下标化；数字后跟字母/下划线不下标化；英文标识符内部的 CO2/XCO2 不下标化 |
| 占位符保护 | `\x00` 包裹的占位符确保 Unicode/化学式转换不被 LaTeX 转义破坏 |

## 编译验证

| 版本 | 错误 | Missing char | Overfull | PDF页数 | 说明 |
|------|------|-------------|----------|---------|------|
| v1.0 | 68 | 1119 | 4 | 11 | 首版：`_`未转义、`$`被破坏 |
| v1.1 | 2 | 2 | 0 | 11 | 修复 `_`转义、占位符机制 |
| v2.0 | 0 | 0 | 0 | 11 | 修复 ×→\times、°→^\circ、化学式识别、公式内_\转义 |

## 文件清单

```
skill/text-extract/
├── SKILL.md               ← 本文档
├── text_extract.py         ← 提取器 Python 模块
└── troubleshooting.md      ← 问题与解决方案
```

## 与其他 Skill 的关系

```
document-extract (图片提取)
    ↓
text-extract (文本提取) ← 使用 omml-to-latex，化学式识别，占位符机制
    ↓
omml-to-latex (公式提取) ← FORMULA_SYMBOL_MAP，公式内_\转义
    ↓
convert_direct.py (LaTeX 生成)
```
