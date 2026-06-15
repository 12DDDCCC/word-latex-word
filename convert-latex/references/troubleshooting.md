# 排查指南 — Word→LaTeX 转换常见问题

## 编译错误

### `File c.sty not found`

**原因**: 模板 preamble 中伪命令行被直接写入 .tex
**特征**: `\usepackage commands included in the copernicus.cls:` 等行
**修复**: `clean_preamble()` 过滤 `\cmd + 空格 + 小写字母` 的伪命令行
**代码位置**: `convert_direct.py:clean_preamble()`

### `Option clash for package babel/natbib`

**原因**: .cls 内部已加载该包，preamble 又显式加载
**修复**: 正则解析 `\usepackage{...}` 中的包名，移除被 cls 内置的包
**代码位置**: `convert_direct.py:assemble_tex()` 中 `cls_builtin_pkgs` 列表

### `Missing character` 大量出现

**原因**: Unicode 字符未转为 LaTeX 命令
**修复**: text-extract 的 Unicode→LaTeX 映射 + 占位符机制
**代码位置**: `text_extract.py` 中 `UNICODE_MAP`

### 编译超时 (timeout)

**原因**: TikZ 表格渲染 + 27 页文档，3 遍编译耗时长
**修复**: xelatex timeout=300s, bibtex timeout=120s
**代码位置**: `convert_direct.py:compile_tex()`

## 内容问题

### 图片未插入 (0/N)

**原因**: 旧版用 context 文本匹配，但所有图片 context 都是 "(图片段落)"
**修复**: 改为 para_index 映射，`img_insert_map`
**代码位置**: `convert_direct.py:build_image_map()` + `assemble_tex()`

### 表格未插入 (0/N)

**原因**: 旧版用 preceding_heading 匹配，但表格可能无标题
**修复**: `get_table_positions()` 直接从 docx XML 读取表格位置
**代码位置**: `convert_direct.py:get_table_positions()` + `build_table_map()`

### CO₂ 下标缺失

**原因**: Word 中 CO2 是纯文本，无下标格式
**修复**: `postprocess_tex()` 正则替换 CO2→CO$_2$
**代码位置**: `convert_direct.py:postprocess_tex()`
**排除**: OCO2→OCO-2 (卫星名，不下标化)

### URL 超出页面

**原因**: 裸 URL 在 LaTeX 中不会自动断行
**修复**: `postprocess_tex()` 用 `\url{}` 包裹所有裸 URL
**代码位置**: `convert_direct.py:postprocess_tex()`

### `\citep{0}` 未定义

**原因**: 源 Word 中异常引用编号，bib 文件无 key=0 的条目
**修复**: 需手动修正源 Word 中的引用编号
**非自动修复**: 此为源数据问题

## 位置匹配问题

### para_index 不连续

**原因**: text-extract 跳过了表格/图片段落
**影响**: 图片/表格的 para_index 在 text_result 中不存在
**解决**: `img_insert_map`/`tbl_insert_map` 映射到最近的前一个正文段落

### 表格位置获取

**方法**: 遍历 `doc.element.body` 的子元素，按 `}p`(段落) 和 `}tbl`(表格) 交替计数
**输出**: `(para_index, table_index)` 列表，与 `table_result['tables']` 顺序一一对应

## 依赖问题

### `ModuleNotFoundError: No module named 'docx'`

```bash
pip install python-docx
```

### XeLaTeX 未安装

Windows 安装 TeX Live 或 MiKTeX，确保 `xelatex` 在 PATH 中。

### copernicus.cls 版本冲突

原始版 .cls 不支持 xelatex，需用修改版（含 xelatex 兼容补丁）。
修改版优先从 docx 同目录复制，参见 `copy_support_files()`。
