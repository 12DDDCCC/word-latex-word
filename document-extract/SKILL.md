---
name: document-extract
description: 从 Word (.docx) 文档中提取嵌入图片，结合 ZIP 解压和 python-docx 关系映射精确定位图片位置
version: 1.1.0
author: 小孟同学
tags: [docx, image-extract, word, latex, python]
---

# 文档图片提取 Skill

从 Word (.docx) 文档中提取嵌入图片，返回图片文件列表及其在文档中的位置信息。

## 适用场景

- Word → LaTeX 转换时需要提取图片并生成 `\includegraphics`
- 批量提取 Word 文档中的所有嵌入图片
- 需要知道图片出现在哪个段落/页面的精确定位

## 技术原理

Word (.docx) 是 ZIP 包，图片存储在 `word/media/` 目录下。文档通过关系文件 (`_rels/document.xml.rels`) 将 `rId` 映射到 `media/imageN.ext`。

```
小论文.docx (ZIP)
├── word/
│   ├── document.xml         ← 正文，含 w:drawing 图片引用
│   ├── media/
│   │   ├── image1.png       ← 实际图片文件
│   │   ├── image2.jpeg
│   │   └── ...
│   └── _rels/
│       └── document.xml.rels  ← rId → media/ 映射
└── [Content_Types].xml
```

## 提取流程

### 步骤1：ZIP 直接提取所有图片

从 ZIP 包的 `word/media/` 目录提取全部图片文件，建立路径→文件名映射。

```python
import zipfile
from pathlib import Path

def extract_images_from_zip(docx_path, output_dir):
    """从 Word ZIP 包中提取所有图片到 output_dir"""
    image_map = {}  # word/media/image1.png -> fig1.png
    fig_count = 0
    img_dir = Path(output_dir) / "fig"
    img_dir.mkdir(exist_ok=True)

    with zipfile.ZipFile(docx_path, 'r') as z:
        for name in z.namelist():
            if name.startswith('word/media/'):
                data = z.read(name)
                ext = Path(name).suffix
                fig_count += 1
                fname = f"fig{fig_count}{ext}"
                with open(img_dir / fname, 'wb') as f:
                    f.write(data)
                image_map[name] = fname

    return image_map  # {'word/media/image1.png': 'fig1.png', ...}
```

### 步骤2：通过 python-docx 精确定位段落中的图片

遍历每个段落的 `w:drawing` → `a:blip` → `r:embed` → `rels[rId]`，找到段落对应的图片。

```python
from docx import Document
from docx.oxml.ns import qn

def get_images_for_paragraph(para, image_map):
    """检查段落中是否包含图片，返回图片文件名列表"""
    images = []
    elem = para._element

    # 查找 w:drawing 中的图片引用
    drawings = elem.findall('.//' + qn('w:drawing'))
    for drawing in drawings:
        # 查找 blip (图片引用)
        blips = drawing.findall('.//' + qn('a:blip'))
        for blip in blips:
            rId = blip.get(qn('r:embed'))
            if rId:
                try:
                    rel = para.part.rels[rId]
                    img_path = rel.target_ref  # e.g. "media/image1.png"
                    # 匹配到 image_map 中的本地文件名
                    key = 'word/media/' + img_path.split('/')[-1]
                    if key in image_map:
                        images.append(image_map[key])
                except Exception:
                    pass

    return images  # ['fig1.png', ...]
```

### 步骤3：组合使用——遍历文档生成图片位置表

```python
def extract_all_images_with_position(docx_path, output_dir):
    """提取所有图片及其在文档中的位置"""
    doc = Document(docx_path)
    image_map = extract_images_from_zip(docx_path, output_dir)

    results = []  # [(段落索引, 图片文件名, 段落文本前缀)]

    for pi, para in enumerate(doc.paragraphs):
        imgs = get_images_for_paragraph(para, image_map)
        for img in imgs:
            # 段落前20字符作为位置提示
            context = para.text[:20] if para.text else "(图片段落)"
            results.append({
                'para_index': pi,
                'image_file': img,
                'caption': context,           # 单行图例
                'caption_full': '',           # 完整图例（多行合并）
                'legend_paragraphs': [],      # 图例段落列表
                'context_above': '',          # 前一段落最后50字
                'context_below': '',          # 后一段落前50字
                'context_above_text': '',     # 图片上方最近的正文段落全文
                'context_below_text': '',     # 图片下方图例之后的正文段落全文
                'width_pt': None,             # 图片宽度(pt)
                'height_pt': None,            # 图片高度(pt)
            })

    return results
```

## 输出示例

```python
results = extract_all_images_with_position('小论文.docx', 'output/')
for r in results:
    print(f"段落{r['para_index']}: {r['image_file']} ({r['context']})")
```

输出：
```
段落0: fig1.jpeg (RegGCAS流程示意图)
段落1: fig2.png (研究区范围，站点分布)
段落2: fig3.jpeg (卫星偏差统计图)
段落3: fig4.png (卫星偏差统计图)
段落4: fig5.jpeg (中国通量的分布)
```

## 完整调用示例

```python
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
import zipfile

# 1. 提取图片
image_map = extract_images_from_zip('小论文.docx', 'output/')

# 2. 遍历段落找图片位置
doc = Document('小论文.docx')
for para in doc.paragraphs:
    images = get_images_for_paragraph(para, image_map)
    if images:
        for img in images:
            print(f"  图片: fig/{img}")
        # 生成 LaTeX \includegraphics
        print(f"  \\includegraphics[width=0.8\\textwidth]{{fig/{img}}}")
```

## 注意事项

| 问题 | 说明 |
|------|------|
| 图片格式 | Word 可能内嵌 PNG、JPEG、EMF、WMF 等格式，EMF/WMF 在 LaTeX 中不直接支持 |
| 图片重复 | 同一张图片在 Word 中多处引用时，ZIP 中只有一个文件 |
| 内联 vs 浮动 | `w:drawing` 是浮动图片，`w:pict` 是内联图片（旧格式） |
| 关系失效 | 某些 rId 可能指向外部链接而非内嵌图片，需检查 `r:link` 属性 |
| 中文路径 | Windows 上 ZIP 打开含中文路径的 docx 时，注意编码问题 |

## 依赖

- Python >= 3.10
- `python-docx` >= 0.8.11
- 标准库 `zipfile`, `pathlib`

```bash
pip install python-docx
```

## v1.1 更新内容

| 新增字段 | 说明 | 用途 |
|----------|------|------|
| `caption_full` | 完整图例文本（多行合并） | 单行 caption 可能截断，caption_full 包含完整图例 |
| `legend_paragraphs` | 图例段落列表 | 图片下方的说明段落，需在 LaTeX 中附加 |
| `context_above` | 前一段落最后50字 | 用于调试和位置确认 |
| `context_below` | 后一段落前50字 | 用于调试和位置确认 |
| `context_above_text` | 图片上方最近的正文段落全文 | 用于调试和位置确认 |
| `context_below_text` | 图片下方图例之后的正文段落全文 | 用于调试和位置确认 |
| `width_pt` | 图片宽度(pt) | 用于 LaTeX 宽度设置参考 |
| `height_pt` | 图片高度(pt) | 用于 LaTeX 高度设置参考 |

### 与 assemble_tex 的映射关系

`assemble_tex()` 使用 `para_index` 进行图片位置映射：

```
图片 para_index → 找最近的"有效"段落 (前一个 heading/body)
  → 跳过 figure_caption/table_caption/empty 段落
  → img_insert_map[text_para_index] = [img_info]
```

**注意**: 图片提取的 `para_index` 是顶层字段，而表格提取的位置信息封装在 `position.paragraph_index` 内。这是两个 skill 的数据结构差异，`convert_direct.py` 中已分别处理。