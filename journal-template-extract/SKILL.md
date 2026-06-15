---
name: journal-template-extract
version: "3.2"
description: 从LaTeX期刊模板(.cls/.sty/.cfg)提取48类完整排版规格，并生成46章完整规格文档+模板.tex文件，支持Copernicus/IEEE/Elsevier/Springer/ACS等
---

# journal-template-extract v3.2

从LaTeX期刊模板(.cls/.sty/.tex/.cfg)中提取**48类完整排版规格**，并**强制生成46章完整规格文档+符合模板格式的LaTeX文件**。

## v3.1 → v3.2 更新

- **14个普适性增强类别**: 支持geometry/fontspec/titlesec/caption包/biblatex/定理/算法/行距/enumerate/表格详细/边注/DOI/cleveref/文档类模式
- **多模板兼容**: 支持`\def`/`\DeclareRobustCommand`/`\providecommand`/`\newcommand`四种定义方式
- **XeLaTeX/LuaLaTeX**: fontspec字体配置提取(含CJK字体)
- **biblatex**: 完整biblatex选项和资源提取
- **titlesec**: titlesec包章节格式提取

## 输入

- 必需: `.cls` 文件 (LaTeX文档类)
- 可选: `.sty` 文件 (宏包), `.tex` 模板文件, `.cfg` 配置文件, `.bst` 参考文献样式

## 输出

| 文件 | 说明 |
|------|------|
| `{journal}_paper.tex` | 完整论文模板文件，包含所有模板要求的段落和排版规格注释 |
| `{journal}_layout_spec.json` | 排版规格JSON (48类完整规格) |
| `{journal}_layout_report.md` | 可读排版规格报告 |
| `{journal}_word_styles.json` | Word样式映射 |
| `template_info.json` | 模板元信息 |
| `{journal}_full_spec.md` | 完整规格文档 (46章节) |

## 48类提取规格

### 基础类 (1-18)
1. 文档类与选项 / 2. 页面布局 / 3. 栏数 / 4. 字体规格 / 5. 章节标题规格 / 6. 编号格式 / 7. 图规格 / 8. 表规格 / 9. 公式规格 / 10. 引用格式 / 11. Caption格式 / 12. 声明段落 / 13. 附录编号 / 14. 页眉页脚 / 15. 行号 / 16. 列表样式 / 17. 颜色主题 / 18. 文档结构

### v3.1增强 (19-34)
19. 浮动体设置 / 20. 数学间距 / 21. 脚注详细 / 22. hyperref / 23. 全局排版 / 24. 标题页布局 / 25. 日期声明 / 26. 摘要详细 / 27. 关键词 / 28. 作者详细 / 29. 自定义排版命令 / 30. 子图设置 / 31. 名称重定义 / 32-34. 页面样式详细

### v3.2普适性增强 (35-48)
35. geometry包配置 / 36. fontspec/XeLaTeX字体 / 37. titlesec包章节 / 38. caption包配置 / 39. biblatex配置 / 40. 定理环境 / 41. 算法环境 / 42. 行距设置 / 43. enumerate样式 / 44. 表格详细格式 / 45. 边注设置 / 46. DOI/URL格式 / 47. cleveref交叉引用 / 48. 文档类模式/选项

## 使用

```bash
# 从模板目录生成
python extract_template.py --template-dir /path/to/template --journal "journal_name"

# 从单个 .cls 文件生成
python extract_template.py --cls-file /path/to/template.cls --journal "journal_name"
```

## 依赖

- Python 3.8+
- `layout_spec_extract.py` (排版规格提取器，同目录下)
