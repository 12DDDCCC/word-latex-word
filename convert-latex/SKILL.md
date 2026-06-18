---
name: convert-latex
description: "Word→LaTeX 无损转换整合 skill。一键将 .docx 转为可直接 xelatex+bibtex 编译的 LaTeX 论文。所有排版格式从模板动态提取，支持任意期刊模板。"
version: 3.9.0
author: 小孟同学
triggers:
  - word转latex
  - docx转tex
  - latex转换
  - 无损转换
  - convert latex
  - word to latex
  - docx to tex
metadata:
  task_type: pipeline
  input_type: docx
  output_type: tex+pdf
  related_skills:
    - text-extract
    - omml-to-latex
    - document-extract
    - table-lossless-extract
    - citation-extract
    - journal-template-extract
---

# Word→LaTeX 无损转换 Skill v3.9

一键将 Word 文档 (.docx) 转换为可直接 xelatex+bibtex 编译的 LaTeX 论文文件。
**所有排版格式从模板动态提取**，支持任意期刊模板（Copernicus、IEEE、Elsevier 等）。

## Quick Start

```bash
python convert_direct.py <docx> <模板目录> <bib> <期刊名> [输出目录] [--no-pdf]
```

**示例：**
```bash
python convert_direct.py 小论文.docx 目标模板/ test/references.bib acp convert_output/
```

**输出：** `convert_output/acp_full.tex` + `convert_output/acp_full.pdf`

---

## Pipeline 流程

```
[1/6] text-extract          → text_extract.json (段落/格式/公式/引用)
[2/6] document-extract       → fig/*.png + 位置信息
[3/6] table-lossless-extract → all_tables_complete.json
[4/6] citation-extract       → citations.json (验证用)
[5/6] journal-template-extract → 期刊骨架 + layout_spec + 支撑文件
[6/6] assemble_tex()         → {journal}_full.tex
      ↓ copy_support_files() → .cls/.bst/.bib/图片
      ↓ postprocess_tex()    → CO₂下标 + URL包裹
      ↓ compile_tex()        → xelatex×3 + bibtex → PDF
```

## 参数说明

| 参数 | 必填 | 说明 | 示例 |
|------|------|------|------|
| docx_path | 是 | Word 文档路径 | `小论文.docx` |
| template_dir | 是 | 含 .cls/.cfg/.bst 的目录 | `目标模板/` |
| bib_path | 是 | references.bib 路径 | `test/references.bib` |
| journal | 是 | 期刊缩写 | `acp` |
| output_dir | 否 | 输出目录（默认 docx 同级 convert_output/） | `convert_output/` |
| --no-pdf | 否 | 不编译 PDF | |

## 6 个子 Skill

| # | Skill | 版本 | 职责 | 输出 |
|---|-------|------|------|------|
| 1 | text-extract | v2.0 | 段落/格式/公式/引用/化学式 | JSON |
| 2 | omml-to-latex | v2.1 | OMML 公式→LaTeX | 被 text-extract 内部调用 |
| 3 | document-extract | v1.0 | 图片+位置(para_index) | 图片文件 + JSON |
| 4 | table-lossless-extract | v3.0 | 表格+TikZ还原(双向) | JSON + TikZ 代码 |
| 5 | citation-extract | v1.0 | 引用标记详情 | JSON(验证用) |
| 6 | journal-template-extract | v2.0 | 期刊模板骨架 | .tex 骨架 + 支撑文件 |

## v3.9 更新内容（2026-06-12）

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 双栏全宽图表原子化 | 双栏模板中的全宽图片/表格使用模板派生的 `figure*`/`table*` 与 `[!t]`，并用 `minipage` 保持图表和 caption 不被拆开 | 避免图表跨页断裂、末尾独占整页或造成大面积空白 |
| 模板 float 策略提取 | 从 `.cls`/模板探针提取 `topfraction`、`dbltopfraction` 等参数，图片高度按模板阈值限制 | 不死编码 ACP/Elsevier/NSR 的浮动阈值 |
| 表格高度估算 | 根据模板文本区、列宽和单元格换行估算 TikZ 表格实际高度，只在真正超页时使用多页表格环境 | 防止表格被错误压缩、截断或拆成两页 |
| 双栏长公式换行 | 按模板列宽对长公式进行 `split` 包装，保留 `\tag`/`\label` 顺序 | 修复 ACP final 中公式跨栏覆盖正文 |

## v3.1 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| FloatBarrier + placeins | section/subsection 后插入 `\FloatBarrier`，加载 `placeins` 包 | 图片浮动到文档末尾 |
| [htbp] 默认浮动 | 图片浮动位置从 `[t]` 改为 `[htbp]` | `[t]` 过于严格导致图片漂移 |
| equation→gather 合并 | 连续 `\begin{equation}` 合并为 `\begin{gather}` | 连续 equation 叠加 abovedisplayskip/belowdisplayskip 间距过大 |
| 双重括号修复 | `((Gui et al., 2024))` → `(Gui et al., 2024)` | 引用转换产生多余括号 |
| Author 智能拆分 | 单词→surname=该词；多词→最后一个为 surname | `\Author[][EMAIL]{given}{surname}` 格式要求 |
| empty 段落跳过映射 | 图片映射时跳过 `empty` 段落 | empty 段落无 LaTeX 输出，图片应映射到有输出的段落 |
| 公式间距由 .cls 控制 | 不在 preamble 中手动设置 abovedisplayskip 等 | copernicus.cls 已定义 `abovedisplayskip=11pt plus 3pt minus 6pt`，手动覆盖会冲突 |

## v3.2 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 表格映射改用 position.paragraph_index | 不再用 `get_table_positions` 从 docx XML 读取，直接用 table_result 的 `position.paragraph_index` | docx XML 的 para_index 与 text_para_indices 不一致，导致表格丢失 |
| 表格映射跳过无效段落 | 表格映射时跳过 `figure_caption`/`table_caption`/`empty`，与图片映射逻辑一致 | 无效段落无 LaTeX 输出，表格应映射到有输出的段落 |
| equation→gather 状态机修复 | 用 IDLE→IN_EQ→AFTER_END 三态状态机替代旧逻辑 | 旧逻辑在遇到公式内容行时错误刷新缓冲，导致合并失效 |
| citep 合并 | `\citep{N}\citep{M}` → `\citep{N,M}` | 同一处引用应合并为一个 citep |
| citep 去括号 | `(\citep{N})` → `\citep{N}` | citep 本身已产生括号，外层括号多余 |
| 插入完整性检查 | 整合完成后检查所有表格/图片是否都已插入，未插入的发出 WARNING | 防止内容静默丢失 |

## v3.3 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 删除死代码 | 移除 `get_table_positions` 和 `build_table_map` 函数 | v3.2 重构后已无调用点 |
| inserted_tbl_ids 统一 | 表格去重从 `id(tbl_data)` 改为 `table_index`，与完整性检查一致 | 两者 ID 类型不一致导致误报 WARNING |
| 公式自动编号 | 不再依赖 Word 中的编号(可能不正确)，使用 `eq_counter` 自动递增生成 `eq1`-`eqN` | Word 编号可能重复或错误(如公式6被标记为(5)) |
| 保留公式后文本 | `, i = 1, 2, ... , N` 等公式行尾部文本不再删除 | 该文本是公式行的一部分，删除会导致公式不完整 |
| 清理空命令 | `\boldsymbol{   }`、`\mathbf{   }` 等空命令自动移除 | 空命令在 gather 中产生空白公式行(幽灵编号) |
| gather 尾部 `\\` 修复 | 从末尾向前查找最后一个带 `\\` 的行并删除，而非只检查最后一行 | label 行在公式行之后，旧逻辑无法删除公式行末尾的 `\\`，导致 gather 产生多余空白公式 |

## v3.8 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| FloatBarrier条件化 | `\FloatBarrier`仅在placeins包加载时才插入，通过`_has_placeins`标志控制 | 无placeins时FloatBarrier未定义导致编译错误 |
| placeins加载放宽 | 有浮动体（图片/表格）时统一加载placeins，不再仅依赖spec.float_position | 多数学术论文都需要浮动控制，spec不声明float_position时也应加载 |
| 多模板验证 | 5个模板(ACP/IEEE/Elsevier/Springer/Nature)全部通过硬编码消除验证 | 确认所有排版差异均来自spec动态提取 |

## v3.9 Table Rule Policy

| Fix | Rule | Why |
|-----|------|-----|
| Template-driven table rules | Derive `layout_spec.table.rule_style`, `hline_commands`, `vertical_rules`, and `no_vertical_rules` from extracted template signals such as `\tophline`, `\middlehline`, `\bottomhline`, `\toprule`, `\midrule`, `\bottomrule`, and sample tabular column specs. | Table rendering must follow the active template, not journal-name heuristics. |
| Preserve source horizontal separators | For no-vertical-rule templates, keep Word source horizontal group separators while mapping line weights to template style. | A template can require no vertical rules without requiring all source horizontal grouping lines to be discarded. |
| No journal-name table fallback | Do not infer table rule style from `journal == acp/copernicus`. Use extracted `table_format` fields or template content only. | Changing to another journal template must not inherit ACP-specific behavior. |

## v3.7 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| FONT_PKG_MAP spec覆盖 | spec.fonts中提取的字体名优先于FONT_PKG_MAP映射 | 不同模板字体不同，FONT_PKG_MAP仅作回退 |
| 关键词集合spec扩展 | 从spec.section_commands/special_envs动态扩展INTRO/CONCLUSION/APPENDIX/DECL关键词 | 不同模板可能有不同的章节别名命令和标题文本 |
| extra_pkgs spec驱动 | tikz/soul/placeins仅在spec声明相关内容时才添加 | 不是所有模板都需要这些包 |
| spacing_map spec扩展 | 行间距命令映射从spec.body_text.spacing_commands扩展 | 不同模板可能定义自定义行间距命令名 |
| 声明默认文本去硬编码 | 声明命令占位文本从模板骨架获取，不再硬编码英文默认文本 | 硬编码英文默认文本不适用于非英文模板 |
| bib_filename spec驱动 | bib_filename优先从layout_spec获取，再回退skeleton_info | 不同模板的bib文件名不同 |
| format_checker spacing扩展 | spacing_cmd_map从spec.body_text.spacing_commands扩展 | 与convert_direct保持一致的spec驱动逻辑 |

## v3.6 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 行间距spec驱动 | SpecAdapter新增 `_derive_body_text_spec()`，从spec.body_text.line_spacing动态设置setspace | 行间距硬编码为single，换模板(如1.5倍行距)无法生效 |
| 公式字体spec驱动 | SpecAdapter新增 `_derive_equation_spec()`，从spec.equation_format.font_size动态推导 | 公式字号未从spec获取 |
| 表格字体spec驱动 | `_derive_table_spec()`增强：header_size/body_size/float_position/alignment从spec获取 | 表格字体硬编码为small，换模板无法适配 |
| Caption位置spec驱动 | `_derive_caption_spec()`增强：figure_position/table_position/font_size从spec获取 | Caption位置和字号未从spec获取 |
| 参考文献spec驱动 | `_derive_bibliography_spec()`增强：natbib_options/bibpunct/font_size从spec获取 | 参考文献格式部分硬编码 |
| Caption格式应用 | preamble中根据spec生成 `\captionsetup{separator/label_weight/font}` | Caption格式spec提取了但未在preamble中应用 |
| 行间距preamble应用 | preamble中根据spec动态加载setspace+spacing命令 | 行间距spec提取了但未在preamble中应用 |
| 参考文献字体应用 | preamble中根据spec生成 `\renewcommand{\bibfont}{\small}` | 参考文献字号spec提取了但未在preamble中应用 |
| bibpunct格式修复 | `\bibpunct` 6参数生成从错误的双大括号改为正确的 `}{` 分隔 | f-string中 `{{{"}{{"` 导致LaTeX编译报错 |
| format_checker增强 | check_paragraph_format/check_caption_format/check_table_format 重写为spec驱动检查 | 6类排版属性(行间距/公式字体/表格字体/Caption字体/Caption位置/参考文献格式)的合规检查 |
| 图片宽度动态化 | 从columns选项推导 `\columnwidth`(双栏)或 `\textwidth`(单栏) | 硬编码 `\columnwidth`，单栏模板不适用 |

## v3.5 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 字体设置动态化 | 字体名从 spec.required_packages 动态推导（FONT_PKG_MAP），不再硬编码 Times New Roman | 不同模板需要不同字体，Springer 不需要 fontspec |
| 附录格式驱动 | 遇到"附录"/"Appendix"关键词时，根据 appendix_format.type 插入 `\appendix` 或 `\begin{appendices}` | appendix_format spec 完全未使用 |
| title_args 驱动 | 支持 `\title[short]{full}` 格式 | 部分模板要求短标题参数 |
| natbib_options/bibpunct | 从 spec.bibliography_format 动态加载 natbib 选项和 bibpunct | cls 未内置 natbib 时需要显式声明 |
| Caption粗体检测修复 | format_checker 检测 `labelfont=bf` 而非 `label=bf` | 消除误报 |

## v3.4 更新内容

| 修复项 | 说明 | 原因 |
|--------|------|------|
| 编号模式检测 | 新增 `detect_template_numbering_mode()` 和 `detect_source_numbering_mode()` | 自动检测模板和源文档的编号模式(simple/sectioned) |
| 编号引用替换 | `convert_numbering_references()` 将正文硬编码编号替换为 `\ref{}`/`\eqref{}` | 正文中"图2.1为"→"图\ref{fig1}为"，LaTeX编译时自动生成正确编号 |
| 预扫描编号映射 | 遍历 figure_caption/table_caption 段落，建立编号→label完整映射 | 确保所有图/表编号都有对应label |
| figure/table 添加 label | 每个 `\begin{figure}` 添加 `\label{figN}`，每个表格添加 `\label{tabN}` | 没有 label 无法使用 `\ref{}` 引用 |
| Caption分隔符修正 | 使用句点"."而非冒号":" | Copernicus格式: `Figure 2. 描述文字` |
| Caption字号修正 | 使用10pt(\small)而非9pt | Copernicus \small = 10pt |
| Caption前缀清理 | 移除 `\caption{}` 内重复的 `Table N:` / `Figure N:` 前缀 | LaTeX自动生成标签，源文档前缀导致双重显示 |

### 编号模式说明

| 模式 | 图片编号 | 表格编号 | 公式编号 | 示例 |
|------|----------|----------|----------|------|
| simple | Figure 1 | Table 1 | (1) | manuscript模式 |
| sectioned | Figure 2.1 | Table 2.1 | (2-1) | final模式 |

### 编号引用替换流程

```
[1] 预扫描 figure_caption/table_caption 段落
    → 提取编号: 图2.1, 图4.3, 表1, ...
    → 按出现顺序映射: 图2.1→fig1, 图4.3→fig5, 表1→tab1, ...

[2] 插入 figure/table 时添加 \label{figN}/\label{tabN}
    → figure: \begin{figure}...\caption{...}\label{fig1}\end{figure}
    → table: \begin{table}...\label{tab1}\end{table}

[3] 插入 equation 时构建公式编号映射
    → 源文档 (3) → \eqref{eq2-3}

[4] 后处理 convert_numbering_references()
    → 图2.1为RegGCAS的流程示意图 → 图\ref{fig1}为RegGCAS的流程示意图
    → 根据公式(3)计算 → 根据公式\eqref{eq2-3}计算
    → 未映射的编号保留原样
```

### Caption格式修正

| 修正项 | 修正前 | 修正后 | 说明 |
|--------|--------|--------|------|
| 分隔符 | `Figure 1:` | `Figure 1.` | Copernicus用句点 |
| 字号 | 9pt | 10pt | \small=10pt |
| 前缀重复 | `\caption{Table 1: xxx}` | `\caption{xxx}` | LaTeX自动生成标签 |

## 核心整合逻辑

### 1. 模板骨架 → 导言区 + 元数据

- 从 `journal-template-extract` 骨架获取 `\documentclass`、`\usepackage`、元数据命令
- `clean_preamble()`: 过滤伪命令行（如 `\usepackage commands included in the copernicus.cls:`）
- 包冲突检测: 移除被 cls 内置的包（babel, natbib），避免 Option clash
- 自动添加缺失包: `tikz`(有表格时)、`soul`、`placeins`(有浮动体时)、`ctex`(中文支持) — 均从spec驱动 (v3.8)
- `extract_skeleton_commands()`: 严格正则提取 `\begin{document}` 后的元数据命令

### 2. 正文内容 → 按段落顺序

以 `text-extract` 的段落为主索引，按 `para_index` 顺序遍历：

| 段落类型 | LaTeX 生成 | 来源 |
|----------|-----------|------|
| 标题 | `\section`/`\subsection`/`\subsubsection`/`\paragraph` | `heading_level` + `latex` |
| 含公式 | 直接用 `latex` 字段 | text-extract (含 omml-to-latex) |
| 含图片 | `\begin{figure}...\includegraphics...\end{figure}` | document-extract |
| 含表格 | TikZ `\begin{tikzpicture}...` | table-lossless-extract |
| 普通文本 | 直接用 `latex` 字段 | text-extract |
| 引用 | `\citep{N}` | text-extract `is_cite` (数字 key) |

### 3. 图片/表格位置匹配 (para_index 映射)

**核心问题**: text-extract 跳过了图片/表格段落，导致 para_index 不连续。

**解决方案**: 双层映射 `img_insert_map` / `tbl_insert_map`，统一使用"找最近有效段落"策略

```
图片 para_index → 找最近的"有效"段落 (前一个 heading/body)
  → 跳过 figure_caption/table_caption/empty 段落
  → img_insert_map[text_para_index] = [img_info_dict]

表格 para_index → 找最近的"有效"段落 (前一个 heading/body)
  → 跳过 figure_caption/table_caption/empty 段落
  → tbl_insert_map[text_para_index] = [tbl_data]
```

- **图片来源**: `build_image_map()` 从 document-extract 结果按 `para_index` 建索引
- **表格来源**: 直接从 table-lossless-extract 结果的 `position.paragraph_index` 获取位置（v3.2）
- 映射时跳过 `figure_caption`/`table_caption`/`empty`，确保映射到有实际输出(heading/body)的段落
- **v3.2**: 表格不再用 `get_table_positions()` 从 docx XML 读取，改用 table_result 自身的 `position.paragraph_index`
- **v3.2**: 兜底检查所有表格/图片是否都已插入，未插入的发出 WARNING 并在末尾补插

### 4. 参考文献段落删除

检测标题含 "References"/"参考文献" 等关键词，其后所有内容跳过。
替换为 `\bibliographystyle{动态样式}` + `\bibliography{动态文件名}`。

### 5. 动态格式提取与应用（v3.0 核心特性，v3.7 全面spec驱动）

所有排版格式从两个来源动态提取，不再硬编码：

#### 来源 1：模板骨架 `parse_template_skeleton()`

| 字段 | 用途 | 默认值 |
|------|------|--------|
| `introduction_cmd` | 引言命令 (`\introduction` 或 `\section`) | `None` → `\section` |
| `conclusions_cmd` | 结论命令 (`\conclusions` 或 `\section`) | `None` → `\section` |
| `statement_cmds` | 声明命令列表 (`\codeavailability{...}` 等) | `{}` |
| `ack_env` | 致谢环境名 (`acknowledgements` / `acknowledgment`) | `acknowledgements` |
| `bib_style` | 参考文献样式 (`copernicus` / `plain` 等) | `plain` |
| `bib_filename` | 参考文献文件名 (`references` 等) | `references` |
| `abstract_env` | 摘要环境名 (`abstract` / `abstract*`) | `abstract` |
| `keywords_cmd` | 关键词命令 (`\keywords{...}` 或环境) | `None` |
| `abstract_after_maketitle` | 摘要在 maketitle 之后 | `True` |

#### 来源 2：排版规格 `layout_spec`（从 .cls 提取）

| 字段 | 用途 | 应用位置 |
|------|------|----------|
| `body_text.line_spacing` | 行距 (single/1.5/double) | preamble setspace |
| `body_text.spacing_commands` | 行间距命令映射（自定义） | preamble spacing命令 (v3.7) |
| `body_text.first_line_indent` | 首行缩进 | preamble \parindent |
| `body_text.paragraph_skip` | 段间距 | preamble \parskip |
| `caption.font_size` | caption 字体 (small/footnotesize) | figure/table caption |
| `caption.separator` | caption 标签分隔符 (. / :) | preamble captionsetup |
| `caption.label_weight` | caption 标签粗体 | preamble captionsetup |
| `caption.figure_position` | 图片 caption 位置 (above/below) | figure 环境 |
| `caption.table_position` | 表格 caption 位置 (above/below) | table 环境 |
| `figure.caption_position` | 图片 caption 位置 | figure 环境 |
| `figure.float_position` | 图片浮动位置 ([t]/[htbp]) | figure 环境 |
| `figure.width` | 图片宽度 (\columnwidth/\textwidth) | includegraphics |
| `table.header_size` | 表头字体 | tikz node font |
| `table.body_size` | 表体字体 | tikz node font |
| `table.caption_position` | 表格 caption 位置 | table 环境 |
| `table.rule_style` | 线型 (template_hlines/booktabs/default) | tikz draw |
| `table.hline_commands` | 模板横线命令 (`tophline`/`middlehline`/`bottomhline` 等) | tikz draw line-weight mapping |
| `table.vertical_rules` | 模板竖线策略 (`none`/`source`) | tikz vertical draw switch |
| `table.no_vertical_rules` | 是否禁用源表竖线 | tikz vertical draw switch |
| `table.float_position` | 表格浮动位置 | table 环境 |
| `table.alignment` | 表格对齐 (center/left/right) | \centering 等 |
| `bibliography.style` | 参考文献样式 | bibliographystyle |
| `bibliography.font_size` | 参考文献字体 | \bibfont |
| `fonts.main_font` | 主字体名 | fontspec \setmainfont (v3.7) |
| `fonts.sans_font` | 无衬线字体名 | fontspec \setsansfont (v3.7) |
| `fonts.mono_font` | 等宽字体名 | fontspec \setmonofont (v3.7) |
| `fonts.math_font` | 数学字体名 | unicode-math \setmathfont (v3.7) |

#### 由 .cls 自身控制的格式（无需在 preamble 重复设置）

- `page_layout` — 页面尺寸、边距（由 .cls 的 \oddsidemargin 等控制）
- `columns` — 栏数（由 \documentclass 选项控制）
- `fonts` — 字体（由 .cls 的字体设置控制；spec.fonts优先覆盖FONT_PKG_MAP，v3.7）
- `title/author/headings` — 标题格式（由 .cls 的 \makeatletter 控制段）
- `footnote` — 脚注格式（由 .cls 控制）
- `numbering` — 编号格式（由 .cls 控制）
- `colors` — 颜色定义（由 .cls 控制）

### 5. 后处理 `postprocess_tex()`

| 修复项 | 正则 | 说明 |
|--------|------|------|
| CO₂ 下标 | `CO(2)` → `CO$_2$` | 保留已有下标格式 |
| XCO₂ 下标 | `XCO(2)` → `XCO$_2$` | 同上 |
| OCO-2 卫星 | `OCO2` → `OCO-2` | 卫星名，非化学式 |
| 标识符保护 | `GlobalGriddedDailyCO_{2}EmissionsDataset` → `GlobalGriddedDailyCO2EmissionsDataset` | 数据集名/变量名/版本号内部不应用化学式下标 |
| OCO-3 卫星 | `OCO3` → `OCO-3` | 同上 |
| URL 包裹 | `https://...` → `\url{https://...}` | 避免超出页面 |
| 双重括号 | `((...))` → `(...)` | 引用转换产生多余括号 (v3.1) |
| citep合并 | `\citep{N}\citep{M}` → `\citep{N,M}` | 同一处引用应合并 (v3.2) |
| citep去括号 | `(\citep{N})` → `\citep{N}` | citep本身已产生括号 (v3.2) |
| 声明占位文本 | 从模板骨架获取，不硬编码英文 | 不同语言模板默认文本不同 (v3.7) |
| 连续公式合并 | `\begin{equation}...\end{equation}`×N → `\begin{gather}...\end{gather}` | 减少叠加 display skip (v3.1) |
| 清理空命令 | `\boldsymbol{   }` / `\mathbf{   }` → 移除 | 空命令产生空白公式行 (v3.3) |
| gather尾部`\\`修复 | 从末尾向前查找最后一个带`\\`的行并删除 | 旧逻辑只检查最后一行(label行)，遗漏公式行的`\\`，导致幽灵编号 (v3.3) |

### 5.0 Inline Identifier Formula Style (v3.10)

Simple inline identifiers extracted from Word equation runs must follow the target template, not LaTeX's default math italic by accident.

Rules:
- If `equation_format.inline_identifier_style` is explicitly extracted as `math_italic`, keep `$...$` math mode.
- If the template has no explicit inline identifier style, render name-like identifiers as upright body text.
- Only safe identifier shapes are converted: `XCO_{2}`, `CO_{2}`, `XCO_{2}\_quality\_flag`, `XCO2\_quality`.
- Real math remains unchanged: `P_{h}`, `X_{i}^{b}`, equations with operators, Greek letters, matrices, or commands.

This keeps variables mathematically correct while preventing dataset/field names such as `XCO2_quality_flag` from becoming unintended italic math.

### 5.1 连续公式合并为 gather (v3.1)

中间仅空行的连续 `\begin{equation}...\end{equation}` 合并为 `\begin{gather}...\end{gather}`。

**原因**: 每个 `equation` 环境都会产生 `abovedisplayskip + belowdisplayskip`，连续公式叠加导致间距过大。
`gather` 环境内部公式间距由 `\jot` 控制，更紧凑。

**规则**:
- 仅空行分隔的连续 equation 才合并（有正文内容则不合并）
- 2个及以上连续 equation 合并为 gather
- 单个 equation 保持不变
- 每行公式末尾加 `\\`（最后一行除外）
- `\label` 保持不变

**重要**: 不在 preamble 中手动设置 `abovedisplayskip`/`belowdisplayskip`，这些值由 `.cls` 文件控制。
例如 copernicus.cls 已定义 `abovedisplayskip=11pt plus 3pt minus 6pt`，手动覆盖会产生冲突。

### 5.2 图片浮动控制 (v3.1, v3.8条件化)

**问题**: `[t]` 浮动位置过于严格，图片会漂移到页面顶部甚至文档末尾。

**解决方案（三重保障）**:
1. **`placeins` 包**: 有浮动体时统一加载 `\usepackage{placeins}` (v3.8)
2. **`[htbp]` 默认浮动**: 从 layout_spec 获取，默认 `htbp`
3. **`\FloatBarrier`**: 仅在 placeins 加载时才在 section/subsection 后插入 (v3.8条件化)

```latex
% 仅在 _has_placeins=True 时插入
\section{Introduction}
\FloatBarrier    % 自动插入
```

**v3.8修复**: `\FloatBarrier` 通过 `_has_placeins` 标志条件化，确保不会在无 placeins 包时产生未定义命令错误。

### 6. 编译 `compile_tex()`

4 步串行编译，每步打印进度：
1. `xelatex` (timeout=300s)
2. `bibtex` (timeout=120s)
3. `xelatex` (timeout=300s)
4. `xelatex` (timeout=300s)

## 输出文件结构

```
output_dir/
├── {journal}_full.tex          ← 完整论文(可编译)
├── {journal}_full.pdf          ← 编译后 PDF
├── references.bib              ← 参考文献
├── copernicus.cls              ← 期刊类文件(优先用修改版)
├── copernicus.bst              ← 引用样式
├── copernicus.cfg              ← 期刊配置
├── fig/                        ← 图片目录
│   ├── fig1.png
│   ├── fig2.emf
│   └── ...
├── text_extract.json           ← 文本提取结果
├── all_tables_complete.json    ← 表格提取结果
├── citations.json              ← 引用提取结果
└── template_info.json          ← 模板信息
```

## 格式审查

转换完成后，应运行排版合规性检查器验证生成的LaTeX文件：

```bash
# 带模板规格检查（推荐）
python format_checker.py <tex_file> --spec <spec.json> [--output report.json]

# 基础检查（无模板规格）
python format_checker.py <tex_file>
```

**示例：**
```bash
python format_checker.py convert_output/acp_full.tex --spec acp_paper/acp_template_spec.json --output compliance_report.json
```

### 检查项目（15类模板规格 + 2类通用检查）

| # | 类别 | 说明 | 检查内容 |
|---|------|------|----------|
| 1 | document_class | 文档类 | 类名是否匹配、选项是否完整 |
| 2 | required_packages | 必需包 | 外部包是否声明（排除cls内置包） |
| 3 | title_format | 标题格式 | \title命令参数结构 |
| 4 | author_format | 作者格式 | \Author或\author命令是否存在 |
| 5 | abstract_format | 摘要格式 | 环境vs命令是否匹配模板 |
| 6 | keywords_format | 关键词格式 | \keywords命令是否存在 |
| 7 | section_commands | 章节命令 | 别名命令（如\introduction）是否使用；alias_of+section_title用于动态扩展关键词 (v3.7) |
| 8 | special_envs | 声明环境 | 必填声明段落是否使用 |
| 9 | figure_format | 图格式 | 位置参数、图片路径、子图包 |
| 10 | table_format | 表格式 | 表格包、table/tikzpicture环境 |
| 11 | caption_format | Caption格式 | 分隔符(./:)、标签粗体、位置 |
| 12 | equation_format | 公式格式 | amsmath、编号、空公式行 |
| 13 | bibliography_format | 参考文献格式 | bibliographystyle、natbib/biblatex |
| 14 | appendix_format | 附录格式 | \appendix声明 |
| 15 | template_specific | 模板特有 | 必填声明、\maketitle |
| 16 | page_layout | 页面布局 | 不应使用geometry/手动设置 |
| 17 | paragraph_format | 段落格式 | 行间距、noindent、ctex |

### 输出

- **控制台**: 分类显示 [OK]/[WARN]/[ERROR]
- **JSON报告** (`--output`): 结构化JSON，含每个类别的pass/issues/warnings列表

## 已知限制

1. **图片 caption**: 暂时为空 `\caption{}`，需手动填写
2. **模板元数据**: 标题已自动替换；作者名智能拆分 given_name/surname（单个词放入 surname）；机构无源数据时保留 `[AFFILIATION PLACEHOLDER]`
3. **`\citep{0}`**: 源 Word 中异常引用编号，bib 中无对应 key
4. **EMF/WMF 图片**: LaTeX 不直接支持，需转为 PDF/PNG
5. **声明命令**: `\codeavailability` 等由模板骨架提供占位符
6. **gather 合并条件**: 仅合并空行分隔的连续 equation，有正文间隔的不合并
7. **公式间距**: 由 .cls 文件控制，不在 preamble 中手动设置 abovedisplayskip 等
8. **公式编号**: 使用自动编号(eq_counter)，不依赖 Word 中的编号（Word编号可能重复或错误）
9. **公式后文本**: `, i = 1, 2, ... , N` 等公式行尾部文本保留在 equation 环境中
10. **编号引用未映射**: 源文档中编号未出现在 caption 中的（如正文单独引用的`图3`），无法建立映射，保留原样
11. **子图编号**: `图4.3d` 中子图字母 `d` 保留在 `\ref{fig7}d` 中，LaTeX编译时 `\ref{fig7}` 生成编号后紧跟 `d`

## 编译验证记录

| 版本 | 错误 | PDF页数 | 图片 | 表格 | CO₂ | URL | 说明 |
|------|------|---------|------|------|-----|-----|------|
| v1.0 | 68 | 11 | 1/5 | 0/3 | — | — | 首版 |
| v1.1 | 0 | 25 | 5/5 | 0/3 | — | — | 修复图片匹配 |
| v2.0 | 0 | 27 | 5/5 | 3/3 | 20处 | 6处 | 修复表格+CO₂+URL |
| v3.1 | 0 | 27 | 5/5 | 3/3 | 20处 | 6处 | FloatBarrier+gather+双括号+empty映射 |
| v3.2 | 0 | 27 | 5/5 | 3/3 | 20处 | 6处 | 表格映射重构+citep合并+去括号+完整性检查 |
| v3.3 | 0 | 23 | 5/5 | 3/3 | 20处 | 6处 | 公式自动编号+保留公式后文本+清理空命令+gather尾部\\修复+参考文献 |
| v3.4 | 0 | 22 | 5/5 | 3/3 | 20处 | 6处 | 编号引用替换(\ref/\eqref)+figure/table label+Caption句点+Caption前缀清理 |
| v3.5 | 0 | 16 | 5/5 | 3/3 | 20处 | 6处 | 字体动态化+附录驱动+title_args+natbib_options+Caption检测修复 |
| v3.6 | 0 | 22 | 5/5 | 3/3 | 20处 | 6处 | 6类排版属性全spec驱动+bibpunct修复+format_checker增强 |
| v3.7 | 0 | 22 | 5/5 | 3/3 | 20处 | 6处 | 硬编码消除:字体/关键词/包/行间距/声明文本/bib名全spec驱动 |
| v3.8 | 0 | 22 | 5/5 | 3/3 | 20处 | 6处 | FloatBarrier条件化+placeins放宽+5模板验证(ACP/IEEE/Elsevier/Springer/Nature) |

## 文件清单

```
skill/convert-latex/
├── SKILL.md                    ← 本文件
├── convert_direct.py           ← 整合脚本(主入口)
├── agents/                     ← 各阶段 agent 说明
│   └── pipeline_stages.md
└── references/                 ← 排查指南
    └── troubleshooting.md
```

## 依赖

- Python >= 3.10
- `python-docx` >= 0.8.11
- XeLaTeX (TeX Live / MiKTeX)
- BibTeX

```bash
pip install python-docx
```
