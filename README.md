# Academic Paper Format Converter (Word ↔ LaTeX)

Word ↔ LaTeX 双向无损转换工具，支持学术论文的完整格式还原。工具会从 Word 提取正文、公式、图片、表格和引用，适配目标期刊 LaTeX 模板，编译 PDF，并可再转换为可编辑 Word 文档。

## 前置要求（必读）

### Word 输入文档要求

1. **参考文献标注方式**（二选一）：
   - **红色数字标注**：引用位置用红色数字标记，例如 `(1)`、`[1]`。
   - **超链接引用**：已使用 Word 插入 → 交叉引用/超链接功能链接到参考文献列表。
   - 两种方式均支持自动提取并还原为 LaTeX `\cite{}` / `\citep{}` / `\citet{}` 命令。

2. **图例/表例格式**：
   - 图例（Figure caption）和表例（Table caption）的**字号必须比正文小**。
   - 系统通过字号差异自动识别 caption，例如正文 12pt，caption 9-10pt。
   - 无需手动标记，系统自动检测。

3. **BibTeX 文件**（必须）：
   - 需要自行提供 `.bib` 格式的参考文献数据库。
   - 文件中的 key 需要与 Word 中引用的作者-年份或引用识别结果对应。
   - 例：`@article{zhang2020, author={Zhang, ...}, ...}`。

4. **高清论文图片**（可选）：
   - 如提供原始高清图片目录，系统会自动匹配并替换 Word 中压缩的图片。
   - 支持 JPEG/PNG 格式。
   - 如不提供，将使用 Word 文档内嵌的图片。

5. **Word 文档结构建议**：
   - 使用 Word 标题样式（标题1/2/3）标记章节层级。
   - 表格使用标准 Word 表格，合并单元格、跨行跨列均支持。
   - 公式使用 Word 公式编辑器（OMML）输入。

### 期刊模板目录

需要提供期刊的 LaTeX 模板文件，目录中应包含：

- `.cls` 文件（文档类，如 `copernicus.cls`、`elsarticle.cls`、`nsr.cls`）
- `.cfg` / `.sty` 文件（配置文件或模板宏包）
- `.bst` 文件（参考文献样式，如 `copernicus.bst`）
- `template.tex` 或期刊示例 `.tex`（用于提取格式规范）

**模板配置模式**：部分期刊模板（如 Copernicus）在 `.cls` 中定义了多种排版配置，系统会自动从模板中提取所有可用模式。

| 模式 | 说明 | 适用场景 |
|------|------|---------|
| `manuscript`（默认） | 经典稿件格式，通常单栏、宽边距 | 投稿、审稿 |
| `final` | 出版终稿格式，通常双栏、窄边距 | 校对、最终版 |
| `discussions` | 讨论版格式 | 预印本、同行讨论 |

默认使用 `manuscript`。如需指定其他模式，可在命令行加入：

```bash
python orchestrator.py input.docx template/ references.bib copernicus output/ --config-mode final
```

系统会根据所选模式自动提取对应的页面尺寸、字体大小、边距、分栏等参数，无需手动配置。

### 软件环境

| 依赖 | 版本要求 | 用途 |
|------|---------|------|
| Python | ≥ 3.10 | 运行环境 |
| Pandoc | ≥ 3.0 | LaTeX → Word 基础转换 |
| XeLaTeX | TeX Live 2022+ 或 MiKTeX | 编译 PDF |
| BibTeX | 随 TeX 发行版安装 | 生成 `.bbl` 参考文献 |
| python-docx | ≥ 0.8 | Word 文件读写 |
| lxml | ≥ 4.9 | XML 解析 |
| Pillow | 推荐安装 | 图片处理 |
| PyMuPDF | 推荐安装 | PDF 位置辅助分析 |
| SimSun 字体 | 系统已安装 | 中文正文字体 |
| Times New Roman | 系统已安装 | 英文正文字体 |

安装 Python 依赖：

```bash
pip install python-docx lxml pillow pymupdf pytest
```

如果 Pandoc 由 WinGet 安装且 Python 无法调用，可额外安装普通版 Pandoc，或设置环境变量 `PANDOC_EXE` 指向 `pandoc.exe`。

## 模块结构

```text
├── orchestrator.py              # 主入口管线（9步流水线）
├── shared/                      # 共享工具
│   ├── caption_detect.py        #   caption 检测（字号+文本模式）
│   ├── caption_utils.py         #   caption 格式化（normalize/clean）
│   ├── latex_text_utils.py      #   LaTeX 文本转义/上下标转换
│   ├── template_config.py       #   模板配置动态提取（多模式）
│   ├── unit_convert.py          #   单位转换（pt/cm/mm/twips）
│   └── word_xml_utils.py        #   Word XML 工具函数
├── text-extract/                # 文本提取（段落/标题/语义分类）
├── document-extract/            # 图片提取（位置+尺寸+caption关联）
├── image-rename/                # 图片顺序匹配（可选高清替换）
├── table-lossless-extract/      # 无损表格（TikZ/tabular → JSON → Word）
├── citation-extract/            # 引用提取 + 交叉引用跳转（bookmark）
├── journal-template-extract/    # 期刊模板提取（CLS/TEX → LaTeX spec）
├── convert-latex/               # Word → LaTeX 转换
├── omml-to-latex/               # 公式转换（OMML ↔ LaTeX 双向）
├── tex-to-word/                 # LaTeX → Word 转换（Pandoc + 后处理）
└── template-extract-lite/       # 轻量模板提取
```

## 使用方法

### Python API

```python
from orchestrator import run_pipeline

result = run_pipeline(
    docx_path="input.docx",           # Word 输入文档
    template_dir="template/",          # 期刊模板目录（含 .cls/.cfg/.bst）
    bib_path="references.bib",         # BibTeX 参考文献
    journal="copernicus",              # 期刊名
    output_dir="output/",              # 输出目录
    compile_pdf=True,                  # 编译 PDF
    convert_word=True,                 # 转换回 Word
    original_images_dir="images/",     # 可选：高清图片目录
    config_mode="final",              # 可选：manuscript/final/discussions
)

print(result.summary())
```

### 命令行

```bash
python orchestrator.py input.docx template/ references.bib copernicus output/
```

选项：

```text
--original-images <dir>   原始高清图片目录
--config-mode <mode>      指定模板模式，如 manuscript/final/discussions
--no-pdf                  不编译 PDF
--no-word                 不转换为 Word
--no-verify               跳过人工确认步骤
--continue-on-error       出错时继续执行
--no-pdf-float-wrap       关闭 final 模式下的 PDF-guided 图片环绕处理
```

单独执行 LaTeX → Word：

```bash
python tex-to-word/tex_to_word.py main.tex -o out.docx --bib references.bib --config-mode final
```

### 输出文件

| 文件 | 说明 |
|------|------|
| `{journal}_full.tex` | 完整 LaTeX 源码 |
| `{journal}_full.pdf` | 编译后的 PDF |
| `{journal}_converted.docx` | 转换回的 Word 文件 |
| `orchestrator_log.json` | 管线执行日志 |
| `*_report.txt/json` | 文本、图片、表格、公式等检查报告 |

## BibTeX 文献处理说明

本工具使用 `.bib` 文件作为参考文献数据源，并在 LaTeX 编译阶段通过 BibTeX 生成 `.bbl`。LaTeX → Word 阶段会优先读取 `.bbl`，用于还原 Word 末尾参考文献和正文引用格式。

### 1. 准备 `.bib` 文件

把所有文献条目放入一个 BibTeX 文件，例如 `references.bib`：

```bibtex
@article{friedlingstein2025global,
  author  = {Friedlingstein, Pierre and others},
  title   = {Global Carbon Budget 2025},
  journal = {Earth System Science Data},
  year    = {2025},
  doi     = {10.xxxx/example}
}
```

正文引用 key 必须与 `.bib` 条目的 key 一致。例如：

```latex
\citep{friedlingstein2025global}
\citet{friedlingstein2025global}
```

### 2. LaTeX 中插入参考文献

生成的 `.tex` 会根据期刊模板和提取结果自动写入类似命令：

```latex
\bibliographystyle{期刊样式}
\bibliography{references}
```

注意：

- `\bibliography{references}` 不写 `.bib` 后缀。
- `references.bib` 必须和主 `.tex` 在同一编译目录，或能被 BibTeX 搜索到。
- 不同期刊使用不同 `.bst`，例如 `copernicus.bst`、`elsarticle-num.bst` 或模板自带样式。

### 3. 编译顺序

BibTeX 正确工作的标准编译顺序是：

```bash
xelatex main.tex
bibtex main
xelatex main.tex
xelatex main.tex
```

`orchestrator.py` 会自动执行这一顺序。第一次 `xelatex` 生成 `.aux`，`bibtex` 根据 `.aux` 和 `.bib` 生成 `.bbl`，后两次 `xelatex` 刷新引用编号和参考文献列表。

### 4. Word 回转时如何使用 `.bbl`

LaTeX → Word 阶段会优先读取同目录的 `.bbl`：

- `.bbl` 用于生成 Word 末尾的 References 区段。
- 若 `.bbl` 中是数字型引用，Word 参考文献会保留 `[1]`、`[2]` 这样的编号。
- Word 正文中的引用会根据 `cite_map` 建立可点击交叉引用。

如果 Word 中参考文献没有编号，优先检查：

1. `.bbl` 是否已生成。
2. `.bbl` 是否与 `.tex` 主文件同名，例如 `NSR_full.tex` 对应 `NSR_full.bbl`。
3. `.bib` key 是否和正文引用 key 完全一致。
4. 期刊模板是否使用数字型 bibliography style。

### 5. 常见 BibTeX 问题

- **引用显示为 `?`**：通常是 `.bib` key 不匹配，或没有运行 BibTeX。
- **没有生成 `.bbl`**：检查 `.aux` 中是否有 `\citation{...}` 和 `\bibdata{...}`。
- **URL 超出栏宽**：尽量不要在 `url`、`doi`、`eprint` 中重复放多个等价长链接；保留 DOI 或一个规范 URL 即可。
- **Word 和 PDF 引用格式不一致**：确认 `.bbl` 已被 LaTeX → Word 阶段读取；工具会从 `.bbl` 判断数字型引用并生成 `[n]` 编号。

## 转换管线

```text
Word (.docx)
  │
  ├─[1] 文本提取 → 段落/标题/语义分类
  ├─[2] 图片提取 → 位置+尺寸+caption
  ├─[3] 图片匹配 → 可选高清替换
  ├─[4] 表格提取 → TikZ/tabular 无损还原
  ├─[5] 引用提取 → 红色数字/超链接 → cite key
  ├─[6] 模板提取 → CLS/TEX → 格式规范
  ├─[7] 整合生成 → 完整 .tex 文件
  ├─[8] LaTeX 编译 → XeLaTeX + BibTeX → PDF
  └─[9] Word 转换 → Pandoc + 后处理 → .docx
```

## 跨栏图表与 Word 空白处理

`final` 模式下默认开启 PDF-guided float wrap，用于缓解双栏 Word 中跨栏图表前后的大片空白：

- 跨栏图片会按 PDF 辅助信息转换为 Word 页顶浮动容器，并设置四周环绕。
- 图片可以压住单栏或双栏区域，让正文围绕其排版，减少强制分页造成的空白。
- 纯表格默认不放入浮动容器，避免 Word 中嵌套浮动表格导致表格缺失或不可见。
- 多个图表在 PDF 中连续排版时，转换会尽量保持连续顺序，但最终仍受 Word 排版引擎影响。

关闭该处理：

```bash
python orchestrator.py input.docx template/ references.bib copernicus output/ --no-pdf-float-wrap
```

## 已验证功能

| 功能 | 状态 |
|------|------|
| Word → LaTeX → PDF → Word 循环 | ✓ |
| 公式 (OMML ↔ LaTeX) | ✓ (display + inline) |
| TikZ/booktabs 表格 | ✓ (含合并单元格) |
| vMerge 合并单元格底线补齐 | ✓ |
| 图片位置还原 | ✓ |
| 文献交叉引用跳转 (bookmark) | ✓ |
| 表格黑色方块消除 | ✓ |
| 模板配置动态提取 | ✓ |
| final 模式跨栏图片四周环绕 | ✓ |
| 数字型参考文献 `[n]` 回填 | ✓ |


- 编译目录中应保留期刊 `.cls`、`.bst`、`.sty`、`.cfg` 等模板支撑文件。
- 最终 Word 视觉效果仍受 Word 排版引擎影响，建议对关键期刊模板进行一次人工打开检查。

## License

MIT
