"""文档图片提取工具模块

从 Word (.docx) 文档中提取嵌入图片，结合 ZIP 解压和 python-docx 关系映射精确定位。
增强功能：
- 提取图片上下文（前后各50字），用于文字匹配定位
- 识别图例(caption)：格式特征（居中、无首行缩进、小字号）+ 文本模式（"图X.X"开头）
- 图例与图片精确关联
"""

import re
import zipfile
import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 共享工具
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from shared.caption_utils import clean_caption as _clean_caption
from shared.caption_detect import (
    is_caption_paragraph as _is_caption_paragraph,
    is_table_caption_text as _is_table_caption_text,
    find_caption_and_context as _find_caption_and_context,
    find_legend_paragraphs as _find_legend_paragraphs,
    get_body_font_size as _get_body_font_size,
    find_context_body_text as _find_context_body_text,
)


def extract_images_from_zip(docx_path, output_dir):
    """从 Word ZIP 包中提取所有图片到 output_dir/fig/"""
    image_map = {}
    fig_count = 0
    img_dir = Path(output_dir) / "fig"
    img_dir.parent.mkdir(parents=True, exist_ok=True)
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

    return image_map


def get_images_for_paragraph(para, image_map):
    """检查段落中是否包含图片，返回图片文件名列表"""
    images = []
    elem = para._element

    drawings = elem.findall('.//' + qn('w:drawing'))
    for drawing in drawings:
        blips = drawing.findall('.//' + qn('a:blip'))
        for blip in blips:
            rId = blip.get(qn('r:embed'))
            if rId:
                try:
                    rel = para.part.rels[rId]
                    img_path = rel.target_ref
                    key = 'word/media/' + img_path.split('/')[-1]
                    if key in image_map:
                        images.append(image_map[key])
                    else:
                        for k, v in image_map.items():
                            if img_path.split('/')[-1] in k:
                                images.append(v)
                                break
                except Exception:
                    pass

    return images


def _is_table_paragraph(para):
    """判断段落是否属于表格单元格"""
    parent = para._element.getparent()
    return parent is not None and parent.tag.split('}')[-1] == 'tc'


def _extract_image_size(drawing):
    """从 w:drawing 元素提取图片尺寸（pt），1pt = 12700 EMU"""
    extent = drawing.find('.//' + qn('wp:extent'))
    if extent is not None:
        cx = extent.get('cx')
        cy = extent.get('cy')
        if cx and cy:
            return round(int(cx) / 12700, 1), round(int(cy) / 12700, 1)
    return None, None


def _get_context_text(text, char_limit=50):
    """截取文本的前N个字符用于匹配定位"""
    if not text:
        return ""
    return text[:char_limit]


def extract_all_images_with_position(docx_path, output_dir):
    """提取所有图片及其在文档中的位置、上下文和图例

    Returns:
        位置信息列表: [{
            'para_index': int,
            'image_file': str,
            'caption': str,
            'legend_paragraphs': list[str],  # 图例段落（图片下方的说明文字）
            'caption_full': str,  # 完整caption（含图例合并）
            'context_above': str,  # 图片前一段落最后50字（兼容旧接口）
            'context_below': str,  # 图片后一段落前50字（兼容旧接口）
            'context_above_text': str,  # 图片上方最近的正文段落全文
            'context_below_text': str,  # 图片下方图例之后的正文段落全文
            'width_pt': float or None,
            'height_pt': float or None,
        }]
    """
    doc = Document(docx_path)
    image_map = extract_images_from_zip(docx_path, output_dir)
    paragraphs = doc.paragraphs

    # 统计正文字号，用于区分图例/表例
    body_font_size = _get_body_font_size(paragraphs)

    # 预计算每个段落的图片列表（避免重复调用）
    _para_images_cache = {}
    for pi, para in enumerate(paragraphs):
        _para_images_cache[pi] = get_images_for_paragraph(para, image_map)

    results = []
    for pi, para in enumerate(paragraphs):
        if _is_table_paragraph(para):
            continue

        imgs = _para_images_cache.get(pi, [])
        if not imgs:
            continue

        # 查找 caption（图片后最近的图例段落）
        caption = ""
        for j in range(pi + 1, min(pi + 6, len(paragraphs))):
            next_para = paragraphs[j]
            next_text = next_para.text.strip()
            if not next_text:
                continue
            if _is_table_caption_text(next_text):
                break
            if _is_caption_paragraph(next_para, body_font_size):
                caption = next_text
                break
            else:
                break

        # 提取图例段落（从 pi+1 开始，跳过 caption 本身如果已计入）
        legend_start = pi + 1
        legend_paragraphs = _find_legend_paragraphs(
            paragraphs, legend_start, body_font_size, expected_kind='figure')

        # 组合完整 caption
        if caption and legend_paragraphs:
            # caption 已包含在 legend_paragraphs[0] 中，用图例组合
            caption_full = ' '.join(legend_paragraphs)
        elif caption:
            caption_full = caption
        elif legend_paragraphs:
            caption_full = ' '.join(legend_paragraphs)
        else:
            caption_full = ""

        # 上下文：兼容旧接口
        context_above = ""
        for k in range(pi - 1, max(pi - 10, -1), -1):
            prev_para = paragraphs[k]
            prev_text = prev_para.text.strip() if prev_para.text else ""
            if not prev_text:
                continue
            if _is_caption_paragraph(prev_para, body_font_size):
                continue
            context_above = prev_text[-50:] if prev_text else ""
            break

        context_below = ""
        for j in range(pi + 1, min(pi + 10, len(paragraphs))):
            next_para = paragraphs[j]
            next_text = next_para.text.strip()
            if not next_text:
                continue
            if _is_caption_paragraph(next_para, body_font_size):
                continue
            if _para_images_cache.get(j, []):
                continue
            if _is_table_paragraph(next_para):
                continue
            context_below = next_text[:50] if next_text else ""
            break

        # 增强上下文：完整正文段落文本
        context_above_text = _find_context_body_text(paragraphs, pi, 'above', body_font_size)
        context_below_text = _find_context_body_text(paragraphs, pi, 'below', body_font_size)

        for img in imgs:
            width_pt, height_pt = None, None
            drawings = para._element.findall('.//' + qn('w:drawing'))
            for drawing in drawings:
                w, h = _extract_image_size(drawing)
                if w is not None:
                    width_pt, height_pt = w, h
                    break

            results.append({
                'para_index': pi,
                'image_file': img,
                'caption': caption,
                'legend_paragraphs': legend_paragraphs,
                'caption_full': caption_full,
                'context_above': context_above,
                'context_below': context_below,
                'context_above_text': context_above_text,
                'context_below_text': context_below_text,
                'width_pt': width_pt,
                'height_pt': height_pt,
            })

    return results


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    if len(sys.argv) < 2:
        print("用法: python document_extract.py <docx_path> [output_dir]")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else '.'

    results = extract_all_images_with_position(docx_path, output_dir)
    for r in results:
        cap = f" | caption: {r['caption']}" if r['caption']else ""
        size = f" | {r['width_pt']}x{r['height_pt']}pt" if r['width_pt'] else ""
        print(f"段落{r['para_index']}: fig/{r['image_file']}{cap}{size}")
        if r['context_above']:
            print(f"  前文(50字): {r['context_above']}")
        if r['context_below']:
            print(f"  后文(50字): {r['context_below']}")
