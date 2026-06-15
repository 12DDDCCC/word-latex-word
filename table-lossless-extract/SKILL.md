---
name: table-lossless-extract
description: Word↔LaTeX 双向无损表格转换。Word→JSON→TikZ(正向) + LaTeX→JSON→Word(反向)。支持booktabs风格、合并单元格、精确边框还原。
version: 3.4.1
triggers:
  - 提取表格
  - word表格
  - 表格提取
  - lossless table
  - docx表格
  - tikz表格
  - latex表格
  - latex转word表格
  - tikz转word
---

# Word↔LaTeX 双向无损表格转换 Skill

## 功能概述

**双向**无损表格转换：

| 方向 | 路径 | 无损度 |
|------|------|--------|
| Word → LaTeX | docx → JSON → TikZ LaTeX → PDF | 100% |
| LaTeX → Word | TikZ/tabular → JSON → docx | 100%(TikZ) / 部分(tabular) |

还原内容包括：
1. **结构**：合并单元格(gridSpan/vMerge)、列宽、行高
2. **边框**：每条边框线独立绘制，粗线/细线/无线的精确还原
3. **格式**：底纹(shading)、垂直对齐(vAlign)、加粗
4. **文本**：段落对齐、字体、字号、加粗/斜体/上下标
5. **位置**：表格在文档中的章节标题、前导段落文本、序号

## 为什么用TikZ而不是tabular

标准LaTeX tabular的`\hline`/`\cline`只能画整行或指定列的横线，无法实现：
- 一行中部分列粗线、部分列无线（如Jinsha行只有col0有粗底边）
- 不同行不同列的垂直线粗细不同
- 精确的部分边框控制

TikZ通过逐条线段`\draw`实现100%边框还原。

## Template-Driven Rule Policy

When `layout_spec.table` is provided by `convert-latex`, table borders must follow the extracted template rule policy:

- `no_vertical_rules=true`: do not draw Word source vertical borders.
- `rule_style=template_hlines`: preserve Word source horizontal group separators, but map top/bottom to thick template rules and internal separators to thin template rules.
- `rule_style=booktabs`: no vertical rules; preserve source horizontal separators with booktabs-like top/mid/bottom weights.
- `rule_style=default`: preserve source horizontal and vertical borders.

Never infer no-vertical behavior from a journal name. Use `layout_spec.table.no_vertical_rules`, `vertical_rules`, and `rule_style`.

## 使用方法

### 正向: Word → LaTeX

```bash
# 完整Pipeline（推荐）
python run_pipeline.py <input.docx> [output_dir] [--no-pdf]

# 分步执行
python extract_all_tables.py <input.docx> <output.json>
python tikz_table_gen.py <input.json> [output_dir] [--no-pdf]
```

### 反向: LaTeX → Word

```bash
# TikZ → JSON → Word
python latex_table_parser.py <input.tex> [output.json]

# Python API
from latex_table_parser import tikz_to_json, tabular_to_json
from gen_table_from_json import generate_docx

json_data = tikz_to_json(tikz_body, table_env)
generate_docx(json_data, 'output.docx')
```

### 测试

```bash
# tabular 转换测试
python test_tabular_conversion.py [output_dir]

# TikZ 回环测试 (JSON→TikZ→JSON→Word)
python test_tikz_roundtrip.py <all_tables_complete.json> [output_dir]
```

## 正向工作流程

### Step 1: Word XML零损失提取 → JSON

脚本: `extract_all_tables.py`

关键点：
- 使用`zipfile`直接读取docx内的`word/document.xml`
- 用`xml.etree.ElementTree`解析
- 对每个`<w:tc>`（包括vMerge=continue的）**必须**解析`<w:tcBorders>`
- 用`[c for c in tr if tag_local(c) == 'tc']`获取直接子cell（避免嵌套表格干扰）

### Step 2: 边框继承规则

Word XML边框优先级：
1. **单元格级`tcBorders`** > 表格级`tblBorders`
2. 没有tcBorders的cell继承tblBorders
3. `val="nil"`表示明确无边框（不继承）
4. 纵向合并(vMerge)的单元格：
   - `restart`行的cell是合并后的实际cell
   - `continue`行不再有独立tc
   - **底边框取continue行的数据**（关键！）

### Step 3: JSON → TikZ LaTeX生成

脚本: `tikz_table_gen.py`

核心算法：
1. 遍历每个cell的4个方向边框(top/right/bottom/left)
2. 将每条边框转换为坐标线段(y, x_start, x_end, width)或(x, y_start, y_end, width)
3. 合并相邻同宽度线段(减少\draw数量)
4. 映射到TikZ坐标系统：
   - x: 列宽累积(twips → cm, 1cm ≈ 567 twips)
   - y: 行高累积(twips → cm)，Y轴向下取负
5. 生成`\draw[line width=Xpt]`和`\node[anchor=center]`

边框宽度映射(Word sz → LaTeX pt)：
- sz >= 8: 1.2pt (粗线)
- sz >= 4: 0.4pt (细线)
- sz < 4: 不绘制

### Step 4: LaTeX编译

使用xelatex编译(支持中文ctex包)：
```bash
xelatex -interaction=nonstopmode tikz_tables.tex
```

## 反向工作流程 (LaTeX → Word)

### Step 1: TikZ → JSON (`tikz_to_json`)

**100%无损**：从TikZ源码恢复所有信息。

解析步骤：
1. 提取`\draw[line width=Npt] (x1,y1) -- (x2,y2);` → 水平/垂直边框线段
2. 提取`\node[attrs] at (x,y) {text};` → 单元格文本+位置+字体
3. **列边界推导** — 频率投票算法：
   - 节点按y坐标聚类为逻辑行
   - 统计每个x坐标出现的行频率
   - 阈值 = max(max_freq × 0.4, 2)，排除跨列噪声
   - 从列中心推导列边界(x_pos)
4. **行边界推导** — draw y坐标 + 节点行中心补充
5. **节点映射** — 列中心节点直接映射，跨列候选用最小距离法确定gridSpan
6. **vMerge检测** — 空行连续检测，gs>1也参与vMerge
7. **边框映射** — 水平线段→cell top/bottom borders（booktabs风格，无竖线）
8. 组装JSON

### Step 2: tabular → JSON (`tabular_to_json`)

**部分无损**：从LaTeX格式规格推导边框/列宽。

解析步骤：
1. 解析col_format: `|l|c|r|p{3cm}|` → 列对齐+竖线+列宽
2. 拆分行：按`\\`分割，记录每行前的规则（\hline/\toprule等）
3. 拆分单元格：用平衡大括号匹配，处理\multicolumn/\multirow
4. 构建边框和文本
5. 组装JSON

### Step 3: JSON → Word (`gen_table_from_json.py`)

直接import使用：
```python
from gen_table_from_json import generate_docx
generate_docx(json_data, 'output.docx')
```

## JSON数据结构

```json
{
  "source_file": "xxx.docx",
  "total_tables": 3,
  "tables": [
    {
      "table_index": 1,
      "position": {
        "paragraph_index": 106,
        "current_heading": "3 Results",
        "table_caption": "Table 1: Site Compare...",
        "caption_full": "Table 1: Site Compare... (完整表例文本)",
        "legend_paragraphs": ["表例段落1", "表例段落2"],
        "context_above": "前一段落最后50字",
        "context_below": "后一段落前50字",
        "context_above_text": "表格上方最近的正文段落全文",
        "context_below_text": "表格下方表例之后的正文段落全文"
      },
      "grid_cols": [{"width_twips": 1700}],
      "rows": [
        {
          "row_height": "360",
          "row_height_rule": "atLeast",
          "cells": [
            {
              "col_start": 0,
              "text": "Site",
              "gridSpan": 2,
              "vMerge": "restart",
              "bold": true,
              "borders": {
                "top": {"val": "single", "sz": "8", "color": "000000"},
                "bottom": {"val": "nil", "sz": "0", "color": "auto"}
              },
              "paragraphs": [{
                "runs": [{
                  "format": {"bold": true, "size_pt": 9.0}
                }]
              }]
            }
          ]
        }
      ]
    }
  ]
}
```

## 边框宽度映射表

| LaTeX 来源 | 含义 | Word sz (1/8pt) | Word 线宽 |
|-----------|------|-----------------|----------|
| TikZ ≥1.0pt | 粗线(toprule/bottomrule) | 8 | 1.0pt |
| TikZ ≥0.7pt | 中粗(midrule) | 6 | 0.75pt |
| TikZ <0.7pt | 细线 | 4 | 0.5pt |
| \toprule | 顶线 | 12 | 1.5pt |
| \midrule | 中线 | 6 | 0.75pt |
| \bottomrule | 底线 | 12 | 1.5pt |
| \hline | 普通线 | 4 | 0.5pt |
| \| (竖线) | 列分隔 | 4 | 0.5pt |

## 已知陷阱

1. **vMerge=continue的边框丢失**：必须解析continue行的tcBorders，底边框通常在continue行而非restart行
2. **python-docx的merge()会重置tcPr**：必须在所有merge操作完成后再设置边框，否则边框会被覆盖
3. **Table Grid样式自动加边框**：还原时必须先设全部none，再逐cell设tcBorders
4. **没有tcBorders的cell**：应继承tblBorders，不能简单忽略
5. **esc()不要转义{}**：否则\textbf{}会被破坏为\textbf\{\}
6. **行高为0或无效**：默认400 twips
7. **标题行检测**：首行首cell的gridSpan == num_cols时为标题行
8. **TikZ booktabs无竖线**：LaTeX编译后无竖线，TikZ→JSON不生成竖直边框
9. **merge后边框传播产生杂线**：对主tc设置边框会传播到合并区域所有底层tc，必须用"主tc集合"方案清除非主tc边框
10. **row_height_rule可能为None**：使用 `or 'atLeast'` 而非 `.get('key', default)` 防止None值

## 文件清单

| 文件 | 职责 |
|------|------|
| `SKILL.md` | 本文件 |
| `extract_all_tables.py` | Word XML零损失提取 → JSON |
| `tikz_table_gen.py` | JSON → TikZ LaTeX生成(+PDF编译) |
| `latex_table_parser.py` | TikZ/tabular → JSON 解析器(反向转换) |
| `gen_table_from_json.py` | JSON还原为Word文档(v3.4.0: meta边框优先+合并后物理tc映射) |
| `test_tikz_roundtrip.py` | TikZ回环测试(JSON→TikZ→JSON→Word) |
| `test_tabular_conversion.py` | tabular转换测试 |
| `run_pipeline.py` | 完整Pipeline入口(Word → JSON → LaTeX → PDF) |
| `to_latex.py` | JSON转标准LaTeX tabular(不支持部分粗线) |

## v3.0 更新内容 — LaTeX→Word 反向转换

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 新增 `latex_table_parser.py` | TikZ→JSON + tabular→JSON 双解析器 | 实现LaTeX→Word反向转换 |
| 新增 `test_tikz_roundtrip.py` | JSON→TikZ→JSON→Word 回环测试 | 验证无损性 |
| 新增 `test_tabular_conversion.py` | tabular→JSON→Word 转换测试 | 验证5种典型表格 |
| 频率投票阈值 | max_freq × 0.4 替代中位数 | 中位数在极端频率分布时排除真正列中心 |
| 跨列gridSpan最小优先 | 距离相近时优先选最小gs | MICASA gs=4→gs=2，避免过度合并 |
| vMerge居中节点检测 | 节点数≤2且只占1-2列的行映射到上方 | Tap/Uum居中节点分到独立行 |
| gs>1支持vMerge | 移除gs>1排除vMerge逻辑 | Site gs=2 也是vMerge起始行 |
| skip_cell删除 | gs>1不生成覆盖列的cell | Word中出现空单元格 |
| vMerge continue继承gridSpan | continue行gs=restart行gs | Word中前两列是独立空单元格而非合并延续 |
| 线宽分段映射 | ≥1.0pt→sz=8, ≥0.7pt→sz=6, <0.7pt→sz=4 | 旧公式round(pt*8)与Word booktabs标准不一致 |
| booktabs无竖线 | v_segments不修正x_pos，不生成竖直边框 | LaTeX编译后无竖线 |
| 标题行y边界补充 | 节点行中心在首条横线上方时插入边界 | BEPS/MICASA/MEAN标题行丢失 |
| _find_interval超出范围 | val<positions[0]→0, val>positions[-1]→最后一行 | 节点y在边界外返回None被跳过 |

## v3.4.1 更新内容 — 后置表例隔短正文识别

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 后置表例补检 | 当表格后出现一个很短的正文续句，再出现显式`TableN.`/`表N，`表例时，允许跳过这个短续句继续识别表例 | 源Word中表2为“表格 -> 通量很难矫正过来。 -> Table2...”布局，旧逻辑遇到短正文即停止，导致表2表例缺失 |
| 严格表例标签 | 新增严格caption标签正则，要求编号后出现空格或`.`/`,`/`，`/`：`等分隔符 | 避免把“表2中可以看到...”这类正文误判为表例 |
| 验证标准 | 表2的`table_caption`和`caption_full`必须包含原始`Table2...`文本；导出Word仍保持5张图、3个真实表格、无TikZ/FIGURE占位符、表格无竖线 | 确保修复表例不破坏图片、表格和模板边框结构 |

## v3.4.0 更新内容 — 模板横线与合并单元格边框修复

| 修复项 | 说明 | 原因 |
|--------|------|------|
| `meta:json` 边框优先 | TikZ反向解析时，若节点携带`meta:json`，必须读取其中的`bd`边框和`va`垂直对齐信息，再映射为JSON的`borders`/`vAlign` | 仅从draw线段推断会在跨行、跨列和分组横线场景吞掉部分横线 |
| 只补横线、不加竖线 | 当前模板规则为无竖线时，解析层必须显式把`left/right`设为`nil`；横向`top/bottom`只来自模板/TikZ/元数据 | ACP/Copernicus是无竖线风格，但横向分组线仍应保留 |
| 合并后物理tc映射 | `gen_table_from_json.py`必须先完成所有merge，再建立`(row, col_start) -> main <w:tc>`映射，之后再写入top/bottom边框 | `table.cell(ri, cs)`在合并区域可能返回非预期tc，导致边框写到错误位置或被清除 |
| 边框责任单一 | `tex-to-word`的Word样式层不得再自动补整行横线；表格边框完全由`table-lossless-extract`根据LaTeX/TikZ/模板提取结果生成 | 避免二次猜测覆盖真实模板编译后的表格结构 |
| 验证标准 | Table 1/2/3导出Word后必须满足：真实表格数量正确、`left/right=0`、源表中的分组横线存在、首行顶线和末行底线存在 | 确保“根据模板编译格式绘制表格”，而不是固定3条线或吞线 |

## v3.3.1 更新内容 — 杂线彻底修复

| 修复项 | 说明 | 原因 |
|--------|------|------|
| Step 6b 重写 | 用"主 tc 集合"方案替代 vMerge/hMerge continue 检查 | 旧方案只清理 vMerge continue 和 hMerge continue 的 tc，遗漏了 gridSpan merge 后的其他非主 tc |
| 主 tc 集合 | 收集所有 JSON cell 对应的 `table.cell(ri, cs)._tc` 的 id，遍历底层 tc 时不在集合中的全部清除边框 | merge 后主 tc 设边框会传播到合并区域所有底层 tc，需确保非主 tc 边框为 nil |
| row_height_rule None 修复 | `row.get('row_height_rule', 'atLeast')` → `row.get('row_height_rule') or 'atLeast'` | JSON 中 row_height_rule 可能为 None，直接传给 set() 触发 TypeError |
| TikZ 输出去除表例 | `process_table()` 不再生成 `\begin{table}`/`\caption`/`\end{table}`，仅输出 `\begin{tikzpicture}...\end{tikzpicture}` | 用户要求只绘制表格，不保留表例 |

## v3.3 更新内容 — merge 操作与边框设置时序修复

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 重写 gen_table_from_json.py | 将执行顺序改为：merge → 设全nil边框 → 逐cell设边框 → 写文本 | python-docx merge() 会重置 tcPr，覆盖之前设置的边框 |
| Site行top边框修复 | gridSpan=2 的 Site 单元格 top=single/8 正确写入 Word | 旧版边框在 merge 前设置，被 merge 覆盖 |
| Jinsha行bottom边框修复 | vMerge restart 行的 bottom=single/8 正确写入 Word | vMerge merge() 覆盖了之前设置的边框 |
| landcover行数据修复 | 元数据覆盖正确填充 text/gridSpan/vMerge | 之前版本 landcover 行无数据 |

## v3.2 更新内容 — 元数据辅助逆向+draw边框映射

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 元数据cell覆盖 | 用meta cells覆盖cell_map(gridSpan/vMerge/text/bold) | 推断路径产生不完整的cell数据 |
| 元数据y_pos修正 | 推断行数不匹配时用meta.yp替换 | Table 2推断22行vs实际21行 |
| draw边框映射=LaTeX格式 | 边框只从draw水平线段推断，不恢复原始Word竖线 | 输出必须匹配LaTeX编译后格式(booktabs无竖线) |
| 删除_meta_to_json | 废弃的快路径函数，会恢复原始Word格式(含竖线) | 与LaTeX编译格式要求矛盾 |
| draw坐标校准x_pos | draw线段xs/xe校准推断的x_pos(容差0.15cm) | Table 2 xe=14.658 vs x_pos[-1]=14.597差距0.061导致0边框 |
| tikz_table_gen Path修复 | 添加`from pathlib import Path` | NameError: Path not defined |

## v2.1 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| position.paragraph_index | 新增字段，表格在文档中的段落索引 | assemble_tex 需要此字段进行表格位置映射 |
| position.caption_full | 新增字段，完整表例文本（含多行合并） | 单行 table_caption 可能截断 |
| position.legend_paragraphs | 新增字段，表例段落列表 | 表格下方的说明段落 |
| layout_spec 传递 | process_table() 接收 layout_spec 参数 | 表格格式从 layout_spec 动态获取 |

## 测试验证记录

| 测试 | Table 1 | Table 2 | Table 3 | Tabular |
|------|---------|---------|---------|---------|
| 列数 | 6 ✓ | 7 ✓ | 7 ✓ | 3/3/3/4/3 ✓ |
| 行数 | 15(14+caption) ✓ | 21 ✓ | 6 ✓ | 3/3/3/3/2 ✓ |
| 边框 | Site top=8 ✓ Jinsha bot=8 ✓ | landcover ✓ | 28条(0%差异) | 5/5 PASS |
| 杂线 | 0 ✓ | 0 ✓ | 0 ✓ | — |
| 竖线 | 0 ✓ | 0 ✓ | 0 ✓ | — |
| gridSpan | Site gs=2 ✓ | BEPS/MICASA/MEAN gs=2 ✓ | — | \multicolumn ✓ |
| vMerge | Tap/Uum/Hkg等 ✓ | Marginal-tropical ✓ | — | — |

## 依赖

- Python >= 3.10
- `python-docx` >= 0.8.11
- XeLaTeX (TeX Live / MiKTeX)
- BibTeX

```bash
pip install python-docx
```
