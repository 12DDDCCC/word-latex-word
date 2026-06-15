# Academic Paper Format Converter (Word ↔ LaTeX)

Word ↔ LaTeX 双向无损转换工具，支持学术论文的完整格式还原。

## 前置要求（必读）

### Word 输入文档要求

1. **参考文献标注方式**（二选一）：
   - **红色数字标注**：引用位置用红色数字标记（如 `(1)`, `[1]`）
   - **超链接引用**：已使用 Word 插入→交叉引用/超链接功能链接到参考文献列表
   - 两种方式均支持自动提取并还原为 LaTeX `\cite{}` 命令

2. **图例/表例格式**：
   - 图例（Figure caption）和表例（Table caption）的**字号必须比正文小**
   - 系统通过字号差异自动识别 caption（例如正文 12pt，caption 9-10pt）
   - 无需手动标记，系统自动检测

3. **BibTeX 文件**（必须）：
   - 需要自行提供 `.bib` 格式的参考文献数据库
   - 文件中的 key 需要与 Word 中引用的作者-年份对应
   - 例：`@article{zhang2020, author={Zhang, ...}, ...}`

4. **高清论文图片**（可选）：
   - 如提供原始高清图片目录，系统会自动匹配并替换 Word 中压缩的图片
   - 支持 JPEG/PNG 格式
   - 如不提供，将使用 Word 文档内嵌的图片

5. **Word 文档结构建议**：
   - 使用 Word 标题样式（标题1/2/3）标记章节层级
   - 表格使用标准 Word 表格（合并单元格、跨行跨列均支持）
   - 公式使用 Word 公式编辑器（OMML）输入

### 期刊模板目录

需要提供期刊的 LaTeX 模板文件，目录中应包含：
- `.cls` 文件（文档类，如 `copernicus.cls`）
- `.cfg` 文件（配置文件，如 `copernicus.cfg`）
- `.bst` 文件（参考文献样式，如 `copernicus.bst`）
- `template.tex`（模板示例文件，用于提取格式规范）

**模板配置模式**：部分期刊模板（如 Copernicus）在 `.cls` 中定义了多种排版配置，系统会自动从模板中提取所有可用模式：

| 模式 | 说明 | 适用场景 |
|------|------|---------|
| `manuscript`（默认） | 经典稿件格式，单栏，宽边距 | 投稿、审稿 |
| `final` | 出版终稿格式，双栏，窄边距 | 校对、最终版 |
| `discussions` | 讨论版格式 | 预印本、同行讨论 |

默认使用 `manuscript`（经典稿件格式）。如需指定其他模式：
```python
result = run_pipeline(
    ...,
    # convert-latex 的 config_mode 参数通过 convert_direct 传递
)
```

系统会根据所选模式自动提取对应的页面尺寸、字体大小、边距、分栏等参数，无需手动配置。

### 软件环境

| 依赖 | 版本要求 | 用途 |
|------|---------|------|
| Python | ≥ 3.10 | 运行环境 |
| Pandoc | ≥ 3.0 | LaTeX → Word 基础转换 |
| XeLaTeX | TeX Live 2022+ 或 MiKTeX | 编译 PDF |
| python-docx | ≥ 0.8 | Word 文件读写 |
| lxml | ≥ 4.9 | XML 解析 |
| SimSun 字体 | 系统已安装 | 中文正文字体 |
| Times New Roman | 系统已安装 | 英文正文字体 |

安装 Python 依赖：
```bash
pip install python-docx lxml
```

## 模块结构

```
├── orchestrator.py              # 主入口管线（9步流水线）
├── shared/                      # 共享工具
│   ├── caption_detect.py        #   caption 检测（字号+文本模式）
│   ├── caption_utils.py         #   caption 格式化（normalize/clean）
│   ├── latex_text_utils.py      #   LaTeX 文本转义/上下标转换
│   ├── template_config.py       #   模板配置动态提取（3种模式）
│   ├── unit_convert.py          #   单位转换（pt/cm/mm/twips）
│   └── word_xml_utils.py        #   Word XML 工具函数
├── text-extract/                # 文本提取（段落/标题/语义分类）
├── document-extract/            # 图片提取（位置+尺寸+caption关联）
├── image-rename/                # 图片顺序匹配（可选高清替换）
├── table-lossless-extract/      # 无损表格（TikZ/tabular → JSON → Word）
├── citation-extract/            # 引用提取 + 交叉引用跳转（bookmark）
├── journal-template-extract/    # 期刊模板提取（CLS → LaTeX spec）
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
)

print(result.summary())
```

### 命令行

```bash
python orchestrator.py input.docx template/ references.bib copernicus output/

# 选项
#   --original-images <dir>   原始高清图片目录
#   --no-pdf                  不编译 PDF
#   --no-word                 不转换为 Word
#   --no-verify               跳过人工确认步骤
#   --continue-on-error       出错时继续执行
```

### 输出文件

| 文件 | 说明 |
|------|------|
| `{journal}_full.tex` | 完整 LaTeX 源码 |
| `{journal}_full.pdf` | 编译后的 PDF |
| `{journal}_converted.docx` | 转换回的 Word 文件 |

## 转换管线

```
Word (.docx)
  │
  ├─[1] 文本提取 → 段落/标题/语义分类
  ├─[2] 图片提取 → 位置+尺寸+caption
  ├─[3] 图片匹配 → 可选高清替换
  ├─[4] 表格提取 → TikZ/tabular 无损还原
  ├─[5] 引用提取 → 红色数字/超链接 → cite key
  ├─[6] 模板提取 → CLS → 格式规范
  ├─[7] 整合生成 → 完整 .tex 文件
  ├─[8] LaTeX 编译 → XeLaTeX → PDF
  └─[9] Word 转换 → Pandoc + 后处理 → .docx
```

## 已验证功能

| 功能 | 状态 |
|------|------|
| Word → LaTeX → PDF → Word 循环 (3轮) | ✓ |
| 公式 (OMML ↔ LaTeX) | ✓ (display + inline) |
| TikZ/booktabs 表格 | ✓ (含合并单元格) |
| vMerge 合并单元格底线补齐 | ✓ |
| 图片位置还原 | ✓ |
| 文献交叉引用跳转 (bookmark) | ✓ |
| 表格黑色方块消除 | ✓ |
| 模板配置动态提取 (3种模式) | ✓ |

## License

MIT
