---
name: text-extract-troubleshooting
description: 文本提取过程中遇到的所有问题及解决方案，供后续参考
version: 1.0.0
author: 小孟同学
tags: [troubleshooting, fix, co2, subscript, unicode, latex, formula]
---

# 文本提取问题与解决方案

开发过程中遇到的每个问题、根因分析、修复方案，以及验证结果。

---

## 问题 1: CO₂ 下标被 LaTeX 转义破坏

**现象**: `CO₂` 转换后输出 `CO$\_\{2\}$`，LaTeX 编译报错 Missing $，PDF 中显示为 `CO_\{2\}` 而非 CO₂

**根因**: `_extract_run()` 中先做 Unicode→LaTeX（₂→`$_{2}$`），再做 LaTeX 特殊字符转义（`_`→`\_`、`{`→`\{`）。转义步骤把 `$_{2}$` 中的 `_`、`{`、`}` 也替换了，变成 `$\_\{2\}$`

**修复**: 改用 **占位符机制**：
1. Unicode 转换时，先用 `\x00` 包裹的占位符替换（如 ₂→`\x00AA\x00`，映射到 `$_{2}$`）
2. 化学式识别也用占位符（CO2→CO + `\x00BB\x00`，映射到 `$_{2}$`）
3. LaTeX 特殊字符转义只处理纯文本，占位符中 `\x00` 不受影响
4. 最后恢复所有占位符为真实 LaTeX 值

**关键代码** (`text_extract.py` _extract_run):
```python
placeholders = {}
for uc, ltx in UNICODE_TO_LATEX.items():
    ph = f'\x00{c1}{c2}\x00'  # 无数字，不会被后续正则匹配
    placeholders[ph] = ltx
    latex = latex.replace(uc, ph)
# ... 化学式、转义 ...
for ph, ltx in placeholders.items():
    latex = latex.replace(ph, ltx)
```

**验证**: 编译后 0 错误，`CO$_{2}$` 正确渲染

---

## 问题 2: 纯文本 CO2 不识别为化学式下标

**现象**: Word 中 `CO2`（纯 ASCII，无 Unicode 下标 ₂）输出为 `CO2`，LaTeX 中 2 不是下标

**根因**: `_extract_run()` 只做了 Unicode 转换，没有识别纯文本中的化学式模式

**修复**: 添加化学式正则识别：
```python
chem_pattern = re.compile(r'([A-Z][A-Za-z]*?)(\d{1,2})(?=[^0-9a-zA-Z_]|$)')
```
- `CO2` → `CO` + `2` → `CO$_{2}$`
- `XCO2` → `XCO` + `2` → `XCO$_{2}$`
- `OCO2` → 跳过（卫星名，非化学式）
- `H2O` → `H` + `2` → `H$_{2}$`O

**排除规则**: OCO2/OCO3 是卫星名，不下标化。数字后面紧跟字母或 `_` 时不下标化（如 `XCO2quality_flag`）

**验证**: `CO2` → `CO$_{2}$` ✓，`OCO-2` 保持不变 ✓

---

## 问题 3: XCO₂quality_flag 连续下标 (Double subscript)

**现象**: Word 中 `XCO₂variable_name` 在公式内转为 `XCO_{2}_variable_name`，LaTeX 报错 Double subscript

**根因**: OMML 公式中 `₂` 被转为 `_{2}`（正确的下标），但紧跟的 `_variable_name` 中 `_` 在公式环境被 LaTeX 解释为下标命令，形成连续下标 `_{2}_variable`

**修复**: 在 `omml_to_latex.py` 的 `_convert_run()` 中，公式文本的 `_` 转义为 `\_`：
```python
if '_' in text:
    text = text.replace('_', '\\_')
```
结果: `XCO_{2}\_variable\_name` — `_{2}` 是下标，`\_variable\_name` 是普通下划线字符

**验证**: 编译后 0 Double subscript 错误 ✓

---

## 问题 4: 公式中 × 和 ° 未转换

**现象**: `1°×1°` 在公式中输出为 `1^{o}×1^{o}` 和 `52^{°}`，LaTeX 报 Missing character（× 和 ° 无法渲染）

**根因**: `omml_to_latex.py` 的 `_convert_run()` 只做了希腊字母映射（GREEK_MAP），没有处理 × (乘号) 和 ° (度) 等数学符号

**修复**: 添加 `FORMULA_SYMBOL_MAP`：
```python
FORMULA_SYMBOL_MAP = {
    '×': '\\times', '°': '^\\circ',
    '≤': '\\leq', '≥': '\\geq', '±': '\\pm',
    '≈': '\\approx', '≠': '\\neq', '∞': '\\infty',
    '∝': '\\propto', '·': '\\cdot',
}
```
在 `_convert_run()` 中 GREEK_MAP 替换后追加：
```python
for uc, latex in FORMULA_SYMBOL_MAP.items():
    text = text.replace(uc, latex)
```

**验证**: `1^{\\circ}\\times1^{\\circ}` ✓，`52^{\\circ}` ✓，0 Missing character ✓

---

## 问题 5: 文本中 `_` 和 `~` 未转义导致 LaTeX 报错

**现象**: `obspack_co2_1_GLOBALVIEWplus` 中 `_` 被 LaTeX 解释为下标命令，报 Missing $ inserted

**根因**: `_extract_run()` 的 LaTeX 特殊字符转义列表缺少 `_` 和 `~`

**修复**: 扩展转义列表：
```python
for old, new in [('\\', '\\textbackslash{}'), ('&', '\\&'), ('%', '\\%'),
                 ('#', '\\#'), ('{', '\\{'), ('}', '\\}'),
                 ('_', '\\_'), ('~', '\\textasciitilde{}')]:
    latex = latex.replace(old, new)
```

**验证**: `obspack\_co2\_1\_GLOBALVIEWplus` ✓，0 Missing $ inserted ✓

---

## 问题 6: URL 超出页面宽度 (Overfull hbox)

**现象**: `https://example.com/dataset/dataaccess/` 等长 URL 导致段落超出页面边界

**根因**: LaTeX 默认断行规则无法在 URL 中间断行

**修复**: 在 LaTeX 预览模板中添加：
```latex
\usepackage{url}
\usepackage{breakurl}
\sloppy  % 允许更宽松的断行
```

**验证**: 0 Overfull hbox ✓

---

## 编译验证历史

| 版本 | 错误 | Missing char | Overfull | PDF页数 | 说明 |
|------|------|-------------|----------|---------|------|
| v1.0 | 68 | 1119 | 4 | 11 | 首版：`_`未转义、`$`被破坏 |
| v1.1 | 2 | 2 | 0 | 11 | 修复 `_`转义、占位符机制 |
| v1.2 | 0 | 0 | 0 | 11 | 修复 ×→\times、°→^\circ、公式内_\转义 |

---

## 转换处理流水线

```
原始文本 text
    ↓
[1] Unicode → 占位符 (₂→\x00AA\x00, 映射到 $_{2}$)
    ↓
[2] 化学式 → 占位符 (CO2→CO + \x00BB\x00, 映射到 $_{2}$)
    ↓
[3] LaTeX 特殊字符转义 (\ _ { } & % # ~ → 对应转义)
    ↓
[4] 恢复占位符 → 真实 LaTeX 值
    ↓
[5] 格式包裹 (\textbf \textit \textsuperscript ...)
    ↓
最终 latex 输出
```

**关键**: 步骤1-2用占位符保护，步骤3只转义纯文本，步骤4恢复时 LaTeX 值不再被转义破坏。