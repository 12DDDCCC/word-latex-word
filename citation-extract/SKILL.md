---
name: citation-extract
description: 从Word文档(docx)中无损提取红色引用标记；支持LaTeX文献引用→Word Bookmark+内部HYPERLINK域无损转换(可点击跳转)
version: 1.2.0
triggers:
  - 提取引用
  - 引用标记
  - citation extract
  - 红色括号
  - docx引用
  - 交叉引用
  - cross reference
  - 文献引用
---

# Word引用标记无损提取 + 交叉引用转换 Skill

## 功能概述

### 模块1: 引用标记提取 (`extract_citations.py`)

从docx文件中无损提取红色引用标记，包括：

1. **标记文本**: `(1)`, `(2,3)`, `(5-7)`, `(22–23–24)` 等
2. **编号列表**: 解析为 `[1]`, `[2,3]`, `[5,6,7]`, `[22,23,24]`
3. **颜色分布**: 红色部分/黑色部分分别记录
4. **位置信息**: 段落索引、字符偏移、所在章节
5. **上下文**: 标记前后的文字

### 模块2: 文献交叉引用构建 (`cross_ref_builder.py`)

将LaTeX的文献引用(`\citep`/`\citet`/`\cite`)转为Word原生交叉引用：

- 参考文献条目 → Word Bookmark (`_Bib_key`)
- `(Author, Year)` → Word内部HYPERLINK域 (`HYPERLINK \l "_Bib_key"`) — 可点击跳转且保留显示文本

**核心特性**:
- 文献引用**可点击跳转**（内部HYPERLINK域 + Hyperlink样式）
- 不设置 `updateFields` 打开刷新开关；保留点击跳转，同时避免 Word 安全提示
- 仅处理文献引用，不处理fig/tab/eq的\ref
- 完全封装为函数，供 `tex_to_word.py` 主程序调用

## 使用方法

```bash
python extract_citations.py <docx文件路径> [output.json]
```

## 识别规则

**只提取包含红色(EE0000)的引用标记**，全黑括号不识别。

标记特征：
- 格式: 括号中的数字，如 `(1)`, `(2,3)`, `(27–32)`
- 颜色: 数字部分为红色(EE0000)，括号可能为黑色
- 编号范围: 1-100

## 关键技术点

### 1. Run嵌套遍历

红色run可能在`<w:hyperlink>`等嵌套容器内，直接子元素遍历会遗漏。

解决: `_iter_runs_recursive()` 递归遍历所有容器内的run：
```python
def _iter_runs_recursive(elem):
    for child in elem:
        ln = tag_local(child)
        if ln == 'r':
            yield child
        else:
            yield from _iter_runs_recursive(child)
```

### 2. 以括号为边界定位

不以红色片段为边界（会遗漏括号外的数字），而是以`()`为边界：
1. 找到所有`(`及其匹配的`)`
2. 检查括号内是否为纯数字模式
3. 检查括号内是否包含红色字符
4. 两个条件都满足才识别为引用标记

### 3. 编号解析

支持三种格式：
- 逗号分隔: `1,2,3` → `[1,2,3]`
- 范围引用: `5-7` → `[5,6,7]`
- 链式破折号: `22–23–24` → `[22,23,24]`（中文破折号`–`）

### 4. 过滤规则

- 编号 > 100 的排除（防止`(2025)`年份误判）
- 无红色字符的排除（全黑型不是引用标记）

## 输出JSON结构

```json
{
  "source_file": "xxx.docx",
  "total_citations": 20,
  "citations": [
    {
      "type": "bracket",
      "text": "(1)",
      "inner": "1",
      "numbers": [1],
      "has_red": true,
      "red_part": "1",
      "black_part": "",
      "bold": false,
      "para_index": 15,
      "char_offset": 42,
      "section": "1 Introduction",
      "before": "...前文...",
      "after": "...后文..."
    }
  ]
}
```

## 已知陷阱

1. **Run嵌套**: 红色run在hyperlink内，必须递归遍历
2. **颜色判断用iter()**: `r.iter()`会递归，`get_direct()`不会，颜色提取必须用iter
3. **括号可能在不同run中**: `(`在黑色run，数字在红色run，`)`在红色run
4. **中文破折号**: `–`不是标准`-`，正则需匹配`[-–—]`
5. **年份误判**: `(2025)`括号内是纯数字但不是引用，用编号上限100过滤
6. **全黑引用不是目标**: 文档中可能有全黑的括号数字，需`has_red`过滤

## 文献交叉引用 API

### `cross_ref_builder.py` 导出函数

| 函数 | 功能 | 入口 |
|------|------|------|
| `insert_bib_cross_references(doc, cite_map)` | 在Word中插入Bookmark+内部HYPERLINK域(文献引用) | `postprocess_docx()` |
| `parse_bbl_items(bbl_path)` | 从`.bbl`提取参考文献key、纯文本和DOI/URL | `_bbl_item_parser.py` |
| `build_references_section(doc, bbl_items, cite_map)` | 直接生成带`_Bib_key` bookmark和悬挂缩进的Word References区 | `_ref_section_builder.py` |
| `ensure_hyperlink_style(doc)` | 确保Hyperlink样式(蓝色下划线) | 被insert调用 |
| `clear_update_fields_on_open(doc)` | 删除打开时自动更新域设置，避免 Word 安全提示 | 被 `insert_bib_cross_references()` 调用 |

### 内部HYPERLINK域代码 XML结构

```xml
<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:fldChar fldCharType="begin"/></w:r>
<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:instrText> HYPERLINK \l "_Bib_44" </w:instrText></w:r>
<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:fldChar fldCharType="separate"/></w:r>
<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>(Wang et al., 2022)</w:t></w:r>
<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:fldChar fldCharType="end"/></w:r>
```

内部 `HYPERLINK` 域保留正文显示文本，点击可跳转到参考文献条目的 `_Bib_key` bookmark。

## 更新日志

### v1.2.0 (2026-06-04)
- 文献引用仍使用内部 `HYPERLINK \l "_Bib_key"` 域保证可点击跳转。
- 不再设置 `w:updateFields` 打开时自动更新域开关；生成的 docx 不应再触发 Word “该文档包含的域可能引用了其他文件，是否更新”安全提示。
- 新增 `clear_update_fields_on_open(doc)`，在 `insert_bib_cross_references()` 结束前清理旧输出或上游工具遗留的打开更新域设置。

## 文件清单

- `SKILL.md` - 本文件
- `extract_citations.py` - 引用标记提取脚本
- `cross_ref_builder.py` - LaTeX交叉引用→Word Bookmark+内部HYPERLINK域代码
- `test/test_cross_ref.py` - 交叉引用功能测试
