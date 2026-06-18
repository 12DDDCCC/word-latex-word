#!/usr/bin/env python3
"""Fill an official Word template DOCX with converted manuscript content.

Pandoc's --reference-doc copies styles from a reference DOCX, but the output
document is still a Pandoc-generated package. This module uses the official
template as the DOCX shell and transplants the converted document body into it,
while merging the source document relationships needed by images, hyperlinks,
and other embedded content.
"""

from __future__ import annotations

import argparse
import copy
import os
import posixpath
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


NS = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

for prefix, uri in NS.items():
    if prefix not in {"ct", "pr"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("", NS["pr"])


def _read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def _clean_mc_ignorable(root: ET.Element) -> None:
    """Remove undeclared prefixes from mc:Ignorable to prevent Word corruption errors.

    ET.tostring may rename prefixes (e.g. w14→ns0), but mc:Ignorable still
    references the old names. Strip any prefix not actually declared on the
    root element.
    """
    mc_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    ignorable = root.attrib.get(f"{{{mc_ns}}}Ignorable")
    if not ignorable:
        return
    declared = {
        k.split("}")[0][1:] if "}" in k else k
        for k in root.attrib
        if "}" in k
    }
    # Collect xmlns: prefix declarations on root
    for event, elem in ET.iterparse([], events=["start-ns"]):
        pass  # only needed for parsed trees, not fresh elements
    # Direct approach: scan root attrib for xmlns:XXX
    for attr_name in root.attrib:
        if attr_name.startswith("xmlns:"):
            declared.add(attr_name[6:])
    # Also register what NS dict provides
    declared.update(NS.keys())
    valid = [p for p in ignorable.split() if p in declared]
    if valid:
        root.set(f"{{{mc_ns}}}Ignorable", " ".join(valid))
    else:
        del root.attrib[f"{{{mc_ns}}}Ignorable"]


def _rels_path_exists(zf: zipfile.ZipFile) -> bool:
    return "word/_rels/document.xml.rels" in zf.namelist()


def _target_part_path(target: str) -> str | None:
    if target.startswith("#") or "://" in target:
        return None
    return posixpath.normpath(posixpath.join("word", target))


def _target_from_part_path(part_path: str) -> str:
    return posixpath.relpath(part_path, "word")


def _unique_part_name(existing: set[str], part_path: str) -> str:
    if part_path not in existing:
        return part_path
    folder, name = posixpath.split(part_path)
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    else:
        ext = "." + ext
    idx = 1
    while True:
        candidate = posixpath.join(folder, f"source_{stem}_{idx}{ext}")
        if candidate not in existing:
            return candidate
        idx += 1


def _next_rid(used: set[str]) -> str:
    idx = 1
    while f"rId{idx}" in used:
        idx += 1
    rid = f"rId{idx}"
    used.add(rid)
    return rid


def _merge_content_types(template_root: ET.Element, source_root: ET.Element) -> ET.Element:
    merged = copy.deepcopy(template_root)
    existing_defaults = {
        elem.attrib.get("Extension")
        for elem in merged.findall(f"{{{NS['ct']}}}Default")
    }
    existing_overrides = {
        elem.attrib.get("PartName")
        for elem in merged.findall(f"{{{NS['ct']}}}Override")
    }

    for elem in source_root.findall(f"{{{NS['ct']}}}Default"):
        ext = elem.attrib.get("Extension")
        if ext and ext not in existing_defaults:
            merged.append(copy.deepcopy(elem))
            existing_defaults.add(ext)

    for elem in source_root.findall(f"{{{NS['ct']}}}Override"):
        part_name = elem.attrib.get("PartName")
        if part_name and part_name not in existing_overrides:
            merged.append(copy.deepcopy(elem))
            existing_overrides.add(part_name)

    return merged


def _referenced_relationship_ids(root: ET.Element) -> set[str]:
    rel_attrs = {
        f"{{{NS['r']}}}id",
        f"{{{NS['r']}}}embed",
        f"{{{NS['r']}}}link",
    }
    ids: set[str] = set()
    for elem in root.iter():
        for attr in rel_attrs:
            value = elem.attrib.get(attr)
            if value:
                ids.add(value)
    return ids


def _copy_source_relationships(
    template_zip: zipfile.ZipFile,
    source_zip: zipfile.ZipFile,
    source_doc: ET.Element,
    copied_parts: dict[str, bytes],
) -> tuple[ET.Element, dict[str, str]]:
    rel_tag = f"{{{NS['pr']}}}Relationship"
    template_rels = (
        _read_xml(template_zip, "word/_rels/document.xml.rels")
        if _rels_path_exists(template_zip)
        else ET.Element(f"{{{NS['pr']}}}Relationships")
    )
    source_rels = (
        _read_xml(source_zip, "word/_rels/document.xml.rels")
        if _rels_path_exists(source_zip)
        else ET.Element(f"{{{NS['pr']}}}Relationships")
    )

    merged = copy.deepcopy(template_rels)
    used_ids = {
        rel.attrib.get("Id")
        for rel in merged.findall(rel_tag)
        if rel.attrib.get("Id")
    }
    used_ids.discard(None)

    existing_parts = set(template_zip.namelist()) | set(copied_parts)
    id_map: dict[str, str] = {}
    referenced_ids = _referenced_relationship_ids(source_doc)

    for source_rel in source_rels.findall(rel_tag):
        old_id = source_rel.attrib.get("Id")
        if not old_id:
            continue
        if old_id not in referenced_ids:
            continue
        new_id = _next_rid(used_ids)
        id_map[old_id] = new_id

        new_rel = copy.deepcopy(source_rel)
        new_rel.set("Id", new_id)

        target_mode = new_rel.attrib.get("TargetMode")
        target = new_rel.attrib.get("Target")
        if target and target_mode != "External":
            source_part = _target_part_path(target)
            if source_part and source_part in source_zip.namelist():
                part_bytes = source_zip.read(source_part)
                dest_part = source_part
                if dest_part in existing_parts:
                    try:
                        existing_bytes = template_zip.read(dest_part)
                    except KeyError:
                        existing_bytes = copied_parts.get(dest_part)
                    if existing_bytes != part_bytes:
                        dest_part = _unique_part_name(existing_parts, source_part)
                copied_parts[dest_part] = part_bytes
                existing_parts.add(dest_part)
                new_rel.set("Target", _target_from_part_path(dest_part))

        merged.append(new_rel)

    return merged, id_map


def _replace_relationship_ids(root: ET.Element, id_map: dict[str, str]) -> None:
    rel_attrs = {
        f"{{{NS['r']}}}id",
        f"{{{NS['r']}}}embed",
        f"{{{NS['r']}}}link",
    }
    for elem in root.iter():
        for attr in rel_attrs:
            value = elem.attrib.get(attr)
            if value in id_map:
                elem.set(attr, id_map[value])


def _transplant_body(template_root: ET.Element, source_root: ET.Element, id_map: dict[str, str]) -> ET.Element:
    new_root = copy.deepcopy(template_root)
    template_body = new_root.find(f"{{{NS['w']}}}body")
    source_body = source_root.find(f"{{{NS['w']}}}body")
    if template_body is None or source_body is None:
        raise ValueError("Both DOCX files must contain word/document.xml body")

    template_sect = template_body.find(f"{{{NS['w']}}}sectPr")
    source_children = []
    for child in list(source_body):
        if child.tag == f"{{{NS['w']}}}sectPr":
            continue
        source_children.append(copy.deepcopy(child))

    for child in list(template_body):
        template_body.remove(child)

    for child in source_children:
        _replace_relationship_ids(child, id_map)
        template_body.append(child)

    if template_sect is not None:
        template_body.append(copy.deepcopy(template_sect))

    return new_root


def fill_template_with_docx_content(
    content_docx: str | Path,
    template_docx: str | Path,
    output_docx: str | Path,
) -> str:
    content_docx = Path(content_docx)
    template_docx = Path(template_docx)
    output_docx = Path(output_docx)
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    copied_parts: dict[str, bytes] = {}

    with zipfile.ZipFile(template_docx, "r") as template_zip, zipfile.ZipFile(content_docx, "r") as source_zip:
        template_doc = _read_xml(template_zip, "word/document.xml")
        source_doc = _read_xml(source_zip, "word/document.xml")
        template_ct = _read_xml(template_zip, "[Content_Types].xml")
        source_ct = _read_xml(source_zip, "[Content_Types].xml")

        merged_rels, id_map = _copy_source_relationships(template_zip, source_zip, source_doc, copied_parts)
        merged_doc = _transplant_body(template_doc, source_doc, id_map)
        merged_ct = _merge_content_types(template_ct, source_ct)

        # Clean mc:Ignorable to remove prefixes that ET may have renamed
        _clean_mc_ignorable(merged_doc)

        skip = {
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "[Content_Types].xml",
        }
        with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            for item in template_zip.infolist():
                if item.filename in skip:
                    continue
                if item.filename in copied_parts:
                    continue
                out_zip.writestr(item, template_zip.read(item.filename))

            for part_name, part_bytes in copied_parts.items():
                out_zip.writestr(part_name, part_bytes)

            out_zip.writestr("word/document.xml", ET.tostring(merged_doc, encoding="utf-8", xml_declaration=True))
            out_zip.writestr(
                "word/_rels/document.xml.rels",
                ET.tostring(merged_rels, encoding="utf-8", xml_declaration=True),
            )
            out_zip.writestr("[Content_Types].xml", ET.tostring(merged_ct, encoding="utf-8", xml_declaration=True))

    return str(output_docx)


def _ensure_child(parent: ET.Element, tag: str) -> ET.Element:
    child = parent.find(tag)
    if child is None:
        child = ET.Element(tag)
        parent.insert(0, child)
    return child


def _paragraph_text(para: ET.Element) -> str:
    return "".join(t.text or "" for t in para.findall(f".//{{{NS['w']}}}t"))


def _paragraph_style_id(para: ET.Element) -> str:
    ppr = para.find(f"{{{NS['w']}}}pPr")
    if ppr is None:
        return ""
    pstyle = ppr.find(f"{{{NS['w']}}}pStyle")
    return pstyle.attrib.get(f"{{{NS['w']}}}val", "") if pstyle is not None else ""


def normalize_template_filled_layout(
    docx_path: str | Path,
    *,
    line_spacing: float = 1.0,
    body_size_pt: float | None = None,
) -> str:
    """Apply LaTeX-derived body line spacing after template-shell filling.

    Uses python-docx API (not raw XML/ElementTree) to avoid namespace
    corruption that causes Word to report "file is corrupted".

    The official Word template can define Normal as 1.5-line spacing and a
    document grid. Since inserted body paragraphs may reference styles that the
    template does not define, Word falls back to Normal and visually creates
    large gaps. Direct paragraph spacing keeps the template shell while matching
    the LaTeX layout more closely.
    """
    import re as _re
    from docx import Document
    from docx.shared import Pt, Twips
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    docx_path = Path(docx_path)
    line_twips_val = max(120, int(round(float(line_spacing) * 240)))
    body_styles = {"", "Normal", "BodyText", "FirstParagraph", "Abstract"}
    skip_styles = {"Title", "Author", "AbstractTitle", "Heading1", "Heading2", "Heading3", "Heading4"}

    # 中文正文首行缩进 2em（11pt body → 22pt）
    indent_pt = (body_size_pt or 11) * 2
    indent_twips = str(int(round(indent_pt * 20)))  # pt → twips

    # 图例/表例字号: 比正文小1pt，最低9pt
    caption_size_pt = max((body_size_pt or 11) - 1, 9)
    caption_half_pts = str(int(round(caption_size_pt * 2)))

    # 叙事引用检测：图号后紧跟这些词 → 正文，不是caption
    _narrative_re = _re.compile(r'^(中|是|为|说明|展示|可以看到|所示|的|和|与|对|在|从|到|及|等|将|被|有|不|也)')

    doc = Document(str(docx_path))

    omml_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    for para in doc.paragraphs:
        style_id = (para.style.style_id if para.style else "") or ""
        text = para.text.strip()
        if style_id in skip_styles:
            continue

        # 检测是否为展示公式段落（含 m:oMathPara）
        has_omath_para = para._element.find(f".//{{{omml_ns}}}oMathPara") is not None
        has_omath = para._element.find(f".//{{{omml_ns}}}oMath") is not None
        has_display_formula = has_omath_para or (has_omath and _re.fullmatch(r"[\s\t()0-9.\-]+", text))

        # 检测是否为图例/表例 caption
        is_caption = False
        cap_m = _re.match(r"^(Figure|Table)\s+[\w.\-]+\.?\s+\S", text)
        if cap_m:
            is_caption = True
        cap_m_cn = _re.match(r"^(图|表)\s*[\d\.]+\s*", text)
        if cap_m_cn:
            after_num = text[cap_m_cn.end():]
            if after_num and not _narrative_re.match(after_num):
                is_caption = True

        # 确定段落类型
        is_body = style_id in body_styles
        if not is_body and not is_caption and not has_display_formula:
            if not text:
                continue
            continue

        pPr = para._element.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            para._element.insert(0, pPr)

        # 清除 numPr（防止黑色方块/项目符号）
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            pPr.remove(numPr)

        # 行距设置
        spacing = pPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            pPr.append(spacing)
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        spacing.set(qn("w:line"), str(line_twips_val))
        spacing.set(qn("w:lineRule"), "auto")

        # 关闭文档网格对齐
        snap = pPr.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            pPr.append(snap)
        snap.set(qn("w:val"), "0")

        # 缩进设置
        ind = pPr.find(qn("w:ind"))
        if has_display_formula or is_caption:
            # 公式段落 & caption: 零缩进
            if ind is None:
                ind = OxmlElement("w:ind")
                pPr.append(ind)
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:right"), "0")
            ind.set(qn("w:firstLine"), "0")
        elif is_body and text:
            # 正文段落: 首行缩进 2em
            if ind is None:
                ind = OxmlElement("w:ind")
                pPr.append(ind)
            ind.set(qn("w:firstLine"), indent_twips)

        # Caption 字号
        if is_caption:
            for run in para.runs:
                run.font.size = Pt(caption_size_pt)

        # 展示公式强制左对齐
        if has_display_formula:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 样式级别也更新 — 使用 python-docx API
    for style_name in ("Normal", "Body Text", "FirstParagraph", "Abstract"):
        try:
            style = doc.styles[style_name]
            pf = style.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = line_spacing
        except KeyError:
            pass

    doc.save(str(docx_path))
    return str(docx_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill a Word template shell with converted DOCX content.")
    parser.add_argument("content_docx", help="Converted DOCX containing manuscript body")
    parser.add_argument("template_docx", help="Official Word template DOCX")
    parser.add_argument("-o", "--output", required=True, help="Output filled DOCX")
    args = parser.parse_args()
    print(fill_template_with_docx_content(args.content_docx, args.template_docx, args.output))


if __name__ == "__main__":
    main()
