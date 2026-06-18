# Pipeline Stages — 各阶段职责说明

本文档描述 `convert_direct.py` 中 6 步 Pipeline 的各阶段职责，供 agent 调用时参考。

## Stage 1: text-extract

**入口**: `text_extract.extract_docx_text(docx_path, output_json)`
**子依赖**: omml-to-latex (公式转换)

**职责**:
- 遍历 docx 所有段落，提取文本/格式/公式/引用
- 跳过表格段落 (`w:tc` 父元素) 和图片段落
- 化学式识别: CO₂→CO$_{2}$，排除 OCO2(卫星)
- 占位符机制: 保护 Unicode 转换结果不被 LaTeX 转义破坏
- 引用检测: 红色(EE0000)编号 → `\citep{N}`

**输出**: `text_extract.json`
- `paragraphs[]`: 含 para_index, heading_level, latex, has_formula, runs[]
- `headings[]`: 标题索引
- `statistics`: 段落/公式/引用统计

**关键注意**:
- para_index 在跳过表格/图片段落后不连续
- 公式通过 omml_to_latex 转换，交替拼接

## Stage 2: document-extract

**入口**: `document_extract.extract_all_images_with_position(docx_path, output_dir)`

**职责**:
- ZIP 解压提取 `word/media/` 下所有图片到 `fig/`
- 通过 python-docx 关系映射 `rId → media/imageN.ext`
- 遍历段落 `w:drawing` 定位图片所在 para_index

**输出**: 图片文件列表
- 每项含 `para_index`, `image_file`, `context`
- 图片文件保存到 `output_dir/fig/`

**关键注意**:
- 图片段落的 para_index 在 text-extract 中不存在
- context 可能是 "(图片段落)"（空段落），不可靠

## Stage 3: table-lossless-extract

**入口**: `extract_all_tables.extract_tables(docx_path)`
**渲染**: `tikz_table_gen.process_table(table_data, table_index)`

**职责**:
- ZIP 直接读取 `word/document.xml` 零损失提取表格
- 解析合并单元格 (gridSpan/vMerge)、边框 (tcBorders/tblBorders)
- 边框继承规则: 单元格级 > 表格级，vMerge=continue 底边框
- 生成 TikZ 代码精确还原表格

**输出**: `all_tables_complete.json`
- `tables[]`: 含 table_index, position, grid_cols, rows[]
- 每个 cell 含 text, gridSpan, vMerge, borders, paragraphs

**关键注意**:
- 表格的 position 信息由 `get_table_positions()` 从 docx XML 直接读取
- TikZ 用 `\draw` 逐条线段绘制，支持部分粗线

## Stage 4: citation-extract

**入口**: `extract_citations.extract_citations(docx_path)`

**职责**:
- 提取引用标记详情（验证用）
- 交叉验证 text-extract 的引用 key 映射

**输出**: `citations.json`
- 引用编号、位置、上下文

**关键注意**:
- 此阶段为验证性质，不直接影响 .tex 生成
- 异常引用如 `\citep{0}` 需人工修正

## Stage 5: journal-template-extract

**入口**: `extract_template.generate_latex_file(template_dir, journal, output_dir)`

**职责**:
- 从模板目录提取 .cls/.cfg/.bst 文件
- 生成期刊骨架 .tex（含 documentclass + 导言区 + 元数据占位符）

**输出**: `template_info.json` + 骨架 .tex + 支撑文件

**关键注意**:
- 修改版 .cls（含 xelatex 兼容补丁）优先从 docx 同目录复制
- 骨架中的元数据命令由 `extract_skeleton_commands()` 过滤

## Stage 6: assemble_tex (核心整合)

**入口**: `convert_direct.assemble_tex(text_result, image_result, table_result, template_result, bib_path, output_dir, docx_path)`

**职责**:
1. 构建匹配索引:
   - `build_image_map()`: 图片按 para_index 建索引
   - `build_table_map()`: 表格按 para_index 建索引 (需 docx_path 读取位置)
2. 双层映射:
   - `img_insert_map`: text_para_index → [image_files]
   - `tbl_insert_map`: text_para_index → [table_data]
3. 按段落顺序生成正文
4. 插入图片/表格到对应位置
5. 兜底: 未匹配项在正文末尾插入
6. 后处理: `postprocess_tex()` 修复化学式、URL
7. 编译验证: `compile_tex()` xelatex×3 + bibtex

**关键注意**:
- 参考文献段落检测后跳过所有后续段落
- `clean_preamble()` 过滤伪命令行
- 包冲突检测移除 cls 内置的包
