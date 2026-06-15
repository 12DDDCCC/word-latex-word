---
name: template-extract-lite
version: "1.0"
description: 从LaTeX期刊模板中提取15类生成.tex文件必需的核心参数，输出spec.json+guide.md
---

# template-extract-lite v1.0

从LaTeX期刊模板(.cls/.sty/.tex/.cfg)中提取**生成正确.tex文件必需的核心参数**。

## 设计原则

**不提取的内容**（模板.cls已定义，编译时自动应用）：
- 行距、字号、页边距、页眉页脚、颜色、数学间距、浮动体内部参数

**只提取的内容**（必须在.tex中正确书写）：
- 文档类和选项、必需包、章节命令格式、作者/标题/摘要结构、声明段落、图表caption、参考文献样式

## 15类核心提取规格

| # | 类别 | 说明 |
|---|------|------|
| 1 | document_class | 文档类名、选项、基类 |
| 2 | required_packages | 必需包列表含选项 |
| 3 | title_format | 标题命令格式(参数数/粗体/居中) |
| 4 | author_format | 作者/affil/correspondence命令参数数 |
| 5 | abstract_format | 摘要环境vs命令格式 |
| 6 | keywords_format | 关键词命令(标签粗体/文本) |
| 7 | section_commands | 章节命令(含\let别名和\@startsection) |
| 8 | special_envs | 声明段落环境(含\generateCommand) |
| 9 | figure_format | 子图包/图片扩展名/路径 |
| 10 | table_format | 表格相关包(booktabs/tabularx等) |
| 11 | caption_format | Caption分隔符/粗体标签/位置 |
| 12 | equation_format | amsmath选项/编号格式 |
| 13 | bibliography_format | natbib/biblatex配置+bibpunct |
| 14 | appendix_format | 附录环境vs命令 |
| 15 | template_specific | 模板特有命令+必填声明列表 |

## 兼容的LaTeX定义形式

- `\newcommand`、`\renewcommand`、`\DeclareRobustCommand`
- `\def`、`\long\def`、`\let`别名
- `\generateCommand`（Copernicus风格）
- `\@makecaption#1#2`（双参数形式）
- `\newenvironment`、`\renewenvironment`

## 输出

| 文件 | 说明 |
|------|------|
| `{journal}_template_spec.json` | 核心规格JSON (15类) |
| `{journal}_template_guide.md` | 简明使用指南 |

## 使用

```bash
python template_extract_lite.py --cls-file template.cls --journal "journal_name"

# 指定输出目录
python template_extract_lite.py --cls-file template.cls --journal "acp" --output-dir ./output
```

## 输入

- 必需: `.cls` 文件 (LaTeX文档类)
- 可选: `.sty`、`.tex`、`.cfg` (暂不单独处理，通过cls间接加载)

## 依赖

- Python 3.8+
- 无第三方依赖