---
name: tex-to-word
description: 将LaTeX源文件(.tex)转换为Word文档(.docx)，确保文字内容与PDF编译结果显示一致，引用为author-year格式，公式可编辑
version: 1.7.0
triggers:
  - tex转word
  - latex转word
  - tex to word
  - latex to docx
  - 论文转word
---

# LaTeX → Word 转换 Skill

## 功能概述

将LaTeX源文件(.tex)转换为Word文档(.docx)，确保文字与PDF编译结果显示一致。

**核心特性**:
- **引用格式与PDF一致**: 正文引用显示为author-year格式(如 `(Kondo et al., 2020)`)，`\citep`→`(Author, Year)`，`\citet`→`Author (Year)`
- **公式可编辑**: LaTeX公式转为Word OMML格式，可在Word中直接编辑
- **章节结构保留**: `\introduction`/`\conclusions`等Copernicus自定义命令正确转为Word标题层级
- **参考文献完整**: 从编译生成的.bbl文件解析条目，Word后处理直接生成References区，引用显示为author-year格式(如 `(Kondo et al., 2020)`)
- **文献交叉引用可跳转**: 参考文献条目→Word Bookmark, `\citep`/`\citet`→内部HYPERLINK域，显示文本保持与LaTeX编译一致，点击即可跳转到参考文献条目
- **图表分页稳定**: 从LaTeX模板派生页面/列数/表格字号；Word表格行、段落、图片和caption写入keep标记，避免图表被拆页或与正文列状态混乱

## v1.7 更新内容（2026-06-12）

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 模板模式推断 | standalone 转 Word 时从 `\documentclass` 选项推断 final/classic/manuscript，并传递给页面几何提取 | 修复 ACP final Word 误用单栏几何 |
| 可编辑 Word keep-together | 表格行写入 `cantSplit`，非末行段落写入 `keepNext`，所有表格/caption段落写入 `keepLines` | 防止表格、图片和caption在Word中跨页拆开 |
| TikZ表格字号 | fallback表格字号从 `layout_spec.table.body_size` 或模板小字号派生 | 修复 ACP final 表2字号过小 |
| PDF视觉一致Word | visual-exact Word按PDF整页图像锚定，页面尺寸使用PDF本身，避免单栏/双栏混排和图表覆盖 | 保证提交检查时Word外观与PDF一致 |

## 使用方法

```bash
python tex_to_word.py <tex文件> [-o 输出.docx] [--bib 参考文献.bib] [--ref-doc 模板.docx]
```

### 参数

| 参数 | 必需 | 说明 |
|------|------|------|
| `tex_file` | 是 | LaTeX源文件路径 |
| `-o, --output` | 否 | 输出Word文件路径(默认同目录同名.docx) |
| `--bib` | 否 | BibTeX参考文献文件(默认自动查找tex同目录.bib) |
| `--ref-doc` | 否 | Word样式模板(可选) |

### 示例

```bash
# 基本转换(自动查找references.bib)
python tex_to_word.py paper.tex

# 指定输出路径和参考文献
python tex_to_word.py paper.tex -o output.docx --bib refs.bib

# 使用Word样式模板
python tex_to_word.py paper.tex --ref-doc template.docx
```

## 转换管道

```
LaTeX源文件(.tex)
  ↓ [Step 1] 预处理: Copernicus自定义命令 → 标准LaTeX
  ↓            平衡大括号匹配处理嵌套\url{}等
  ↓ [Step 2] 提取: 公式(LaTeX源码) + 表格(结构化JSON) + 图片路径 + TikZ表格
  ↓ [Step 3] Pandoc: LaTeX → Word (citeproc生成author-year引用)
  ↓ [Step 4] 后处理: 验证OMML公式 + 表格边框重建 + 字体设置
  ↓ [Step 5] 嵌入图片: 将\includegraphics图片嵌入Word
  ↓ [Step 6] TikZ表格: 解析TikZ绘图命令 → 重建标准Word表格
  ↓ [Step 7] 模板样式回填: 从.tex/spec/cls派生字体、字号、行距并应用到Word
  ↓
Word文档(.docx)
```

## 预处理映射表

### 无参数命令 → Section

| Copernicus命令 | 转换结果 |
|---------------|---------|
| `\introduction` | `\section{Introduction}` |
| `\conclusions` | `\section{Conclusions}` |

### Statement命令(带平衡大括号) → Section + 内容

| 命令 | 转换结果 |
|------|---------|
| `\dataavailability{...}` | `\section{Data Availability}` + 内容 |
| `\codeavailability{...}` | `\section{Code Availability}` + 内容 |
| `\codedataavailability{...}` | `\section{Code and Data Availability}` + 内容 |
| `\authorcontribution{...}` | `\section{Author Contributions}` + 内容 |
| `\competinginterests{...}` | `\section{Competing Interests}` + 内容 |
| `\sampleavailability{...}` | `\section{Sample Availability}` + 内容 |
| `\disclaimer{...}` | `\section{Disclaimer}` + 内容 |
| `\copyrightstatement{...}` | `\section{Copyright Statement}` + 内容 |

### 其他命令

| 命令 | 处理方式 |
|------|---------|
| `\Author[affil][email]{First}{Last}` | → `\author{First Last}` |
| `\affil[n]{text}` | → 注释掉 |
| `\begin{acknowledgements}` | → `\begin{acknowledgment}` |
| `\runningtitle{...}` | → 注释掉 |
| `\runningauthor{...}` | → 注释掉 |
| `\received{...}/\published{...}` 等 | → 注释掉 |
| `\bibliography{../references}` | → `\bibliography{references}` |

## 质量指标(实测)

| 元素 | 保留率 | 说明 |
|------|--------|------|
| 纯文本 | 90-98% | Pandoc高质量转换 |
| 引用格式 | 95%+ | citeproc生成author-year,与PDF显示一致 |
| 数学公式 | 60-70% | OMML可编辑,复杂公式(align/multiline)可能简化 |
| 章节结构 | 95%+ | 所有自定义命令正确映射 |
| 参考文献 | 90%+ | BibTeX完整处理 |
| 图片 | 100% | 自动嵌入Word |
| TikZ表格 | 90%+ | 解析TikZ绘图命令重建标准Word表格 |
| 标准表格 | 50-60% | 简单表格好,复杂合并单元格表格可能简化 |
| 交叉引用 | 85%+ | Bookmark+内部HYPERLINK域,可点击跳转,显示文本不被目标参考文献覆盖

## 关键技术点

### 1. 平衡大括号匹配

`\dataavailability{\url{https://...}}` 中 `\url{}` 嵌套在大括号内,
简单正则 `(.*?)` 无法匹配。使用 `_match_balanced_braces()` 函数逐字符
跟踪大括号深度,正确提取完整内容。

### 2. TikZ表格解析

TikZ表格使用`\draw`和`\node`命令绘制,非标准`tabular`环境。
使用正则提取`\node[anchor=center] at (x,y) {text}`模式,
按y坐标分组为行,按x坐标排序为列,重建标准Word表格。

### 3. Author命令双方括号

`\Author[][email]{First}{Last}` 有两个可选方括号参数,
正则 `\\Author(\[.*?\])?\{.*?\}\{.*?\}` 只匹配一个。
改为 `(?:\[[^\]]*\]){0,2}` 支持零到两个方括号。

### 3. Bibliography路径

`\bibliography{../references}` 使用相对路径指向父目录,
Pandoc在工作目录中找不到。预处理将路径规范化为
`\bibliography{references}`, 并复制bib文件到工作目录。

### 4. citeproc引用格式

Pandoc的 `--citeproc` 将 `\citep{key}` 转为author-year格式
(如 `(Kondo et al. 2020)`), 而非LaTeX默认的数字编号 `[1]`。
这与PDF中Copernicus期刊的引用显示格式一致。

## 依赖

| 工具 | 版本 | 说明 |
|------|------|------|
| Python | 3.8+ | 运行脚本 |
| Pandoc | 3.x | LaTeX→Word转换引擎 |
| python-docx | 1.x | Word文档后处理 |
| latex2mathml | 可选 | 公式备用转换 |

## 输出文件

- `{name}.docx` — 转换后的Word文档
- `_tex2word_work/` — 调试工作目录(保留中间文件)

调试目录包含:
- `prepared.tex` — 预处理后的LaTeX文件
- `extract_info.json` — 提取的公式和表格信息
- `pandoc_output.docx` — Pandoc原始输出
- `references.bib` — 复制的参考文献文件

## 已知陷阱

1. **嵌套大括号**: `\dataavailability{\url{...}}` 必须用平衡大括号匹配
2. **TikZ表格**: `\node`命令的位置坐标需精确解析,复杂布局可能错位
3. **Author双方括号**: `\Author[1][email]{...}{...}` 需要支持0-2个方括号
4. **bib路径**: `\bibliography{../references}` 需路径规范化
5. **cls/sty缺失**: Pandoc需要cls/sty文件才能解析某些命令,必须复制到工作目录
6. **复杂公式**: align/multiline环境转OMML质量有限
7. **复杂表格**: 合并单元格/多行表头可能简化
8. **图片路径**: `\includegraphics`中的相对路径需正确解析

## 适配范围

适用于使用Copernicus模板的期刊论文(ACP, AMT, BG等)。
通用article类论文也可转换,预处理步骤会跳过不存在的Copernicus命令。

## 文件清单

- `SKILL.md` — 本文件
- `tex_to_word.py` — 主转换脚本

## 更新日志

### v1.6.0 (2026-06-04)
- 修复LaTeX到Word图片回写：预处理阶段用`[FIGURE_N]`保护完整figure环境，后处理阶段按LaTeX中的`\includegraphics`路径插入真实图片，并保留原始`\thefigure`编号与caption文本。
- 修复图例重复：caption解析使用平衡大括号，legend从完整`\caption{...}`命令之后开始截取，避免`Figure 1.`或正文说明重复进入图例。
- 修复TikZ表格回写边框职责：Word样式层不再自动猜测或补齐整行横线；所有表格top/bottom横线由`table-lossless-extract`从LaTeX/TikZ/meta信息生成，left/right竖线按模板规则保持nil。
- 修复合并单元格表格边框错位：反向生成Word表格时依赖`table-lossless-extract`的合并后物理tc映射，避免跨行/跨列后的分组横线丢失。
- 验证标准：导出Word必须包含LaTeX中的全部图片、全部TikZ表格、原始图表编号；表格必须无竖线，且保留模板编译结果中的首行顶线、末行底线和源表分组横线。

### v1.5.0 (2026-06-04)
- 新增 LaTeX 模板派生的 Word 样式回填步骤：`tex_to_word.py` 在全部内容插入后执行 `Step 7`，调用 `apply_template_word_styles()`。
- Word 正文、标题、图例/表例、参考文献和表格文字的字体、字号、行距从生成的 `.tex` 前导区、`*_template_spec.json` 和期刊 `.cls` 解析获得，不再使用 Pandoc 默认样式或固定期刊硬编码。
- 当前 ACP/Copernicus manuscript 模板实测派生为：正文 Times New Roman 11 pt，中文 SimSun，标题 Arial 11 pt 加粗，caption/table/reference 10 pt，正文行距 1.4。
- 更换期刊模板时，优先读取 `\setmainfont`、`\setsansfont`、`\setmonofont`、`\setCJKmainfont`、`\documentclass` 选项、`\@setfontsize` 和 `\baselinestretch`；模板未声明时才使用通用字体兜底。
- 验证标准：导出的 docx 中 `Normal`、`Heading 1-3`、`Caption`、`Bibliography` 样式和段落直接格式必须与模板派生值一致，表格 run 字号必须使用模板 small size。

### v1.3.0 (2026-06-02)
- 新增交叉引用无损转换: `\label`→Word Bookmark, `\ref`/`\eqref`→REF域代码(\h超链接开关)
- 交叉引用可点击跳转到目标位置，编号随文档编辑自动更新
- 集成 `cross_ref_builder.py` (citation-extract skill)，在预处理阶段提取label-ref映射
- `prepare_tex_for_pandoc()` 返回值增加 `label_ref_map`
- `postprocess_docx()` 新增交叉引用插入步骤

### v1.2.0 (2026-06-01)
- 公式编号修复：仅对独立公式段落(text≤10且含OMML)添加编号，避免内联公式误编号
- TikZ表头解析重写：支持多行表头，按x坐标分配列位置，加粗行自动识别为表头
- 图片定位修复：在Pandoc已有图片位置旁插入caption/legend，不再追加到文档末尾
- TikZ表格定位修复：替换[TIKZ_TABLE_N]占位符，不再追加到文档末尾
- Caption分隔符：冒号":"改为句点"."，匹配Copernicus格式(Figure 1.)
- Caption字号：9pt→10pt，匹配Copernicus \small定义

### v1.4.0 (2026-06-04)
- 修复LaTeX到Word文献交叉引用跳转: `.bbl` 不再注入给 pandoc 生成参考文献列表。
- `prepare_tex_for_pandoc()` 将 `\bibliography{...}` 替换为 `[REFERENCES_PLACEHOLDER]`，只保留参考文献区位置。
- `postprocess_docx()` 解析 `.bbl` 条目，用 python-docx 直接生成 `References` 段落、悬挂缩进、DOI/URL链接和 `_Bib_key` bookmark。
- `insert_bib_cross_references()` 复用已有 `_Bib_key` bookmark，只替换正文引用为 `HYPERLINK \l "_Bib_key"` 域，避免 `REF` 域把整条参考文献显示到正文。
- 验证标准: bookmark数量应等于`.bbl`条目数，所有HYPERLINK域目标必须存在，`REF _Bib_*` 和 `[REFERENCES_PLACEHOLDER]`不能残留。

### v1.1.0 (2025-05-31)
- 新增TikZ表格解析功能,自动重建为标准Word表格
- 新增图片自动嵌入功能,支持`\includegraphics`路径解析
- 优化表格边框重建逻辑

### v1.0.0 (2025-05-29)
- 初始版本,支持基础LaTeX→Word转换
