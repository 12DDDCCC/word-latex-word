"""PDF-guided floating containers for editable cross-column DOCX floats.

This is an opt-in alternative to section-break wrapped full-width floats.
The existing section-break path is stable, but Word balances the preceding
two-column section when it closes, which can create large split blank areas.

This module keeps the original figure/table XML editable and changes the
layout container: cross-column float blocks are moved into a borderless
floating Word table (`w:tblpPr`).  Figure drawings inside those containers are
also converted from `wp:inline` to `wp:anchor` with square wrapping, so Word's
UI reports them as wrapped rather than embedded inline pictures.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import tempfile

from lxml import etree

from _pdf_float_reflow import (
    NS,
    qn,
    _blank_score,
    _collect_float_blocks,
    _docx_to_workdir,
    _find_paragraph_by_text,
    _has_drawing,
    _is_caption_like,
    _is_heading_like,
    _is_score_better,
    _matching_block_indices,
    _norm,
    _parse_document_xml,
    _text_of,
    _write_docx_from_workdir,
    export_docx_to_pdf,
)

WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def wp(tag: str) -> str:
    return f"{{{WP_NS}}}{tag}"


@dataclass
class WrapCandidate:
    captions: list[str]
    source_page: int | None = None
    y_twips: int | None = None
    caption_y_twips: int | None = None
    section_anchor: str = ""


def _guidance_records(items) -> list[tuple[str, dict]]:
    records: list[tuple[str, dict]] = []
    for item in items or []:
        guidance = item.get("pdf_guidance")
        if not isinstance(guidance, dict):
            continue
        caption = str(item.get("caption") or item.get("caption_full") or "").strip()
        number = str(item.get("number") or "").strip()
        candidates = [caption]
        if number and caption:
            candidates.extend([
                f"Figure {number} {caption}",
                f"Fig. {number} {caption}",
                f"Table {number} {caption}",
                f"图{number}{caption}",
                f"表{number}{caption}",
            ])
        for text in candidates:
            key = _norm(text)
            if key:
                records.append((key, guidance))
    return records


def _guidance_for_caption(caption: str, records: list[tuple[str, dict]]) -> dict | None:
    caption_norm = _norm(caption)
    if not caption_norm:
        return None
    for key, guidance in records:
        if not key:
            continue
        size = min(56, len(key), len(caption_norm))
        for window in (56, 42, 30, 20, 12):
            if window > size:
                continue
            if key[:window] in caption_norm or caption_norm[:window] in key:
                return guidance
    return None


def _guidance_page(guidance: dict | None) -> int | None:
    try:
        page = int((guidance or {}).get("page"))
    except (TypeError, ValueError):
        return None
    return page if page > 0 else None


def _guidance_y_twips(guidance: dict | None) -> int | None:
    try:
        y_pt = float((guidance or {}).get("y0_pt"))
    except (TypeError, ValueError):
        return None
    if y_pt < 0:
        return None
    return int(round(y_pt * 20))


def _guidance_caption_y_twips(guidance: dict | None) -> int | None:
    try:
        y_pt = float((guidance or {}).get("caption_y_pt"))
    except (TypeError, ValueError):
        return None
    if y_pt < 0:
        return None
    return int(round(y_pt * 20))


def _candidate_list(docx_path: Path, guidance_items=None, max_group_size: int = 4) -> list[WrapCandidate]:
    records = _guidance_records(guidance_items)
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        _docx_to_workdir(docx_path, workdir)
        _, body = _parse_document_xml(workdir)
        blocks = _collect_float_blocks(body)

    candidates: list[WrapCandidate] = []
    idx = 0
    while idx < len(blocks):
        block = blocks[idx]
        if not block.caption:
            idx += 1
            continue
        guidance = _guidance_for_caption(block.caption, records)
        source_page = _guidance_page(guidance)
        captions = [block.caption]
        last = block
        idx += 1

        while idx < len(blocks) and len(captions) < max_group_size:
            cur = blocks[idx]
            cur_guidance = _guidance_for_caption(cur.caption, records)
            cur_page = _guidance_page(cur_guidance)
            same_pdf_page = source_page is not None and cur_page == source_page
            adjacent_in_docx = cur.start_idx <= last.end_idx + 2
            same_section = cur.section_anchor == block.section_anchor
            if not (same_pdf_page and adjacent_in_docx and same_section):
                break
            captions.append(cur.caption)
            last = cur
            idx += 1

        candidates.append(
            WrapCandidate(
                captions=captions,
                source_page=source_page,
                y_twips=_guidance_y_twips(guidance),
                caption_y_twips=_guidance_caption_y_twips(guidance),
                section_anchor=block.section_anchor,
            )
        )
    return candidates


def _body_text_width_twips(body) -> int:
    sect_pr = body.find("w:sectPr", NS)
    if sect_pr is None:
        for elem in reversed(list(body)):
            sect_pr = elem.find("w:pPr/w:sectPr", NS)
            if sect_pr is not None:
                break
    if sect_pr is None:
        return 9000

    pg_sz = sect_pr.find("w:pgSz", NS)
    pg_mar = sect_pr.find("w:pgMar", NS)
    try:
        page_w = int(pg_sz.get(qn("w:w"))) if pg_sz is not None else 11906
        left = int(pg_mar.get(qn("w:left"))) if pg_mar is not None else 1440
        right = int(pg_mar.get(qn("w:right"))) if pg_mar is not None else 1440
    except (TypeError, ValueError):
        return 9000
    return max(page_w - left - right, 3600)


def _append_nil_borders(parent, tag: str = "w:tblBorders") -> None:
    borders = etree.SubElement(parent, qn(tag))
    edges = ["top", "left", "bottom", "right"]
    if tag == "w:tblBorders":
        edges.extend(["insideH", "insideV"])
    for edge in edges:
        elem = etree.SubElement(borders, qn(f"w:{edge}"))
        elem.set(qn("w:val"), "nil")


def _zero_cell_margins(tc_pr) -> None:
    margins = etree.SubElement(tc_pr, qn("w:tcMar"))
    for edge in ("top", "left", "bottom", "right"):
        elem = etree.SubElement(margins, qn(f"w:{edge}"))
        elem.set(qn("w:w"), "0")
        elem.set(qn("w:type"), "dxa")


def _empty_zero_paragraph():
    p = etree.Element(qn("w:p"))
    p_pr = etree.SubElement(p, qn("w:pPr"))
    spacing = etree.SubElement(p_pr, qn("w:spacing"))
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "1")
    spacing.set(qn("w:lineRule"), "exact")
    return p


def _child_by_local(elem, local: str):
    for child in list(elem):
        if child.tag.rsplit("}", 1)[-1] == local:
            return child
    return None


def _position(tag: str, relative_from: str, align: str | None = None, offset: int = 0):
    pos = etree.Element(wp(tag))
    pos.set("relativeFrom", relative_from)
    if align:
        align_elem = etree.SubElement(pos, wp("align"))
        align_elem.text = align
    else:
        offset_elem = etree.SubElement(pos, wp("posOffset"))
        offset_elem.text = str(offset)
    return pos


def _convert_inline_to_square_anchor(inline) -> bool:
    """Convert one wp:inline drawing to wp:anchor with square wrapping."""
    if inline.tag != wp("inline"):
        return False

    children = {
        local: _child_by_local(inline, local)
        for local in ("extent", "effectExtent", "docPr", "cNvGraphicFramePr")
    }
    graphic = None
    for child in list(inline):
        if child.tag == f"{{{A_NS}}}graphic":
            graphic = child
            break
    if children["extent"] is None or children["docPr"] is None or graphic is None:
        return False

    inline.clear()
    inline.tag = wp("anchor")
    for key, value in {
        "simplePos": "0",
        "relativeHeight": "0",
        "behindDoc": "0",
        "locked": "0",
        "layoutInCell": "1",
        "allowOverlap": "0",
        "distT": "0",
        "distB": "0",
        "distL": "0",
        "distR": "0",
    }.items():
        inline.set(key, value)

    simple_pos = etree.SubElement(inline, wp("simplePos"))
    simple_pos.set("x", "0")
    simple_pos.set("y", "0")
    inline.append(_position("positionH", "column", align="center"))
    inline.append(_position("positionV", "paragraph", offset=0))
    inline.append(children["extent"])
    if children["effectExtent"] is not None:
        inline.append(children["effectExtent"])
    wrap = etree.SubElement(inline, wp("wrapSquare"))
    wrap.set("wrapText", "bothSides")
    inline.append(children["docPr"])
    if children["cNvGraphicFramePr"] is not None:
        inline.append(children["cNvGraphicFramePr"])
    inline.append(graphic)
    return True


def _convert_drawings_to_square_anchors(nodes: list) -> int:
    converted = 0
    for node in nodes:
        for inline in list(node.iter(wp("inline"))):
            if _convert_inline_to_square_anchor(inline):
                converted += 1
    return converted


def _estimate_float_content_height_twips(nodes: list) -> int | None:
    max_graphic_twips = 0
    text_twips = 0
    for node in nodes:
        for extent in node.iter(wp("extent")):
            try:
                max_graphic_twips = max(max_graphic_twips, int(extent.get("cy") or 0) // 635)
            except (TypeError, ValueError):
                pass
        for p in node.iter(qn("w:p")):
            text = "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()
            if text:
                text_twips += max(240, ((len(text) // 70) + 1) * 240)
    if max_graphic_twips <= 0 and text_twips <= 0:
        return None
    return max_graphic_twips + text_twips + 120


def _pdf_text_blocks_for_page(pdf_path: Path, page_number: int) -> list[tuple[float, float, float, float, str]]:
    if page_number <= 0 or not pdf_path.exists():
        return []
    try:
        import fitz
    except Exception:
        return []

    try:
        with fitz.open(str(pdf_path)) as doc:
            page_index = page_number - 1
            if not (0 <= page_index < doc.page_count):
                return []
            page = doc[page_index]
            height = float(page.rect.height or 1.0)
            blocks = []
            for block in page.get_text("blocks"):
                x0, y0, x1, y1, text, *_ = block
                text = " ".join((text or "").split())
                if not text:
                    continue
                if y0 > height * 0.88 and len(text) <= 8:
                    continue
                blocks.append((float(x0), float(y0), float(x1), float(y1), text))
            return sorted(blocks, key=lambda item: (item[1], item[0]))
    except Exception:
        return []


def _matches_candidate_caption(text: str, candidate: WrapCandidate) -> bool:
    text_norm = _norm(text)
    if len(text_norm) < 8:
        return False
    for caption in candidate.captions:
        cap_norm = _norm(caption)
        if not cap_norm:
            continue
        size = min(len(text_norm), len(cap_norm), 44)
        for window in (44, 32, 24, 16, 10, 8):
            if window <= size and (text_norm[:window] in cap_norm or cap_norm[:window] in text_norm):
                return True
    return False


def _pdf_body_anchor_text(pdf_path: Path | None, candidate: WrapCandidate) -> str:
    """Return body text on the source PDF page that should follow this float."""
    if pdf_path is None or candidate.source_page is None:
        return ""
    blocks = _pdf_text_blocks_for_page(pdf_path, candidate.source_page)
    usable = []
    for x0, y0, x1, y1, text in blocks:
        if len(_norm(text)) < 10:
            continue
        if _is_caption_like(text) or _matches_candidate_caption(text, candidate):
            continue
        usable.append((x0, y0, x1, y1, text))
    if not usable:
        return ""

    after_y = None
    if candidate.caption_y_twips is not None:
        after_y = candidate.caption_y_twips / 20.0 + 8.0
    elif candidate.y_twips is not None:
        after_y = candidate.y_twips / 20.0 + 36.0
    if after_y is not None:
        after = [block for block in usable if block[1] >= after_y]
        if after:
            return after[0][4]
    return usable[0][4]


def _is_body_anchor_elem(elem) -> bool:
    if elem is None or elem.tag != qn("w:p"):
        return False
    if _has_drawing(elem) or _is_heading_like(elem):
        return False
    text = _text_of(elem)
    return bool(text and not _is_caption_like(text) and len(_norm(text)) >= 8)


def _next_body_anchor_after(body, elem):
    children = list(body)
    try:
        start = children.index(elem) + 1
    except ValueError:
        return None
    for candidate in children[start:]:
        if _is_body_anchor_elem(candidate):
            return candidate
    return None


def _anchor_before_section(body, anchor, section_anchor_text: str) -> bool:
    if not section_anchor_text or anchor is None:
        return False
    section_anchor = _find_paragraph_by_text(body, section_anchor_text)
    if section_anchor is None:
        return False
    children = list(body)
    try:
        return children.index(anchor) < children.index(section_anchor)
    except ValueError:
        return False


def _section_body_anchor(body, section_anchor_text: str):
    if not section_anchor_text:
        return None
    section_anchor = _find_paragraph_by_text(body, section_anchor_text)
    if section_anchor is None:
        return None
    return _next_body_anchor_after(body, section_anchor) or section_anchor


def _target_anchor_for_candidate(body, candidate: WrapCandidate, source_pdf_path: Path | None, fallback):
    anchor_text = _pdf_body_anchor_text(source_pdf_path, candidate)
    anchor = _find_paragraph_by_text(body, anchor_text) if anchor_text else None
    if anchor is None:
        return fallback
    if _anchor_before_section(body, anchor, candidate.section_anchor):
        return _section_body_anchor(body, candidate.section_anchor) or fallback
    if _is_heading_like(anchor):
        return _next_body_anchor_after(body, anchor) or anchor
    return anchor


def _floating_table(
    nodes: list,
    width_twips: int,
    y_twips: int | None,
    desc: str,
    min_height_twips: int | None = None,
):
    tbl = etree.Element(qn("w:tbl"))
    tbl_pr = etree.SubElement(tbl, qn("w:tblPr"))

    tblp_pr = etree.SubElement(tbl_pr, qn("w:tblpPr"))
    tblp_pr.set(qn("w:leftFromText"), "0")
    tblp_pr.set(qn("w:rightFromText"), "0")
    tblp_pr.set(qn("w:topFromText"), "0")
    tblp_pr.set(qn("w:bottomFromText"), "0")
    tblp_pr.set(qn("w:horzAnchor"), "margin")
    tblp_pr.set(qn("w:tblpXSpec"), "center")
    tblp_pr.set(qn("w:vertAnchor"), "margin")
    if y_twips is None:
        tblp_pr.set(qn("w:tblpYSpec"), "top")
    else:
        tblp_pr.set(qn("w:tblpY"), str(max(y_twips, 0)))

    overlap = etree.SubElement(tbl_pr, qn("w:tblOverlap"))
    overlap.set(qn("w:val"), "never")

    tbl_w = etree.SubElement(tbl_pr, qn("w:tblW"))
    tbl_w.set(qn("w:w"), str(width_twips))
    tbl_w.set(qn("w:type"), "dxa")

    jc = etree.SubElement(tbl_pr, qn("w:jc"))
    jc.set(qn("w:val"), "center")
    _append_nil_borders(tbl_pr)

    desc_el = etree.SubElement(tbl_pr, qn("w:tblDescription"))
    desc_el.set(qn("w:val"), desc)

    grid = etree.SubElement(tbl, qn("w:tblGrid"))
    col = etree.SubElement(grid, qn("w:gridCol"))
    col.set(qn("w:w"), str(width_twips))

    tr = etree.SubElement(tbl, qn("w:tr"))
    if min_height_twips:
        tr_pr = etree.SubElement(tr, qn("w:trPr"))
        tr_height = etree.SubElement(tr_pr, qn("w:trHeight"))
        tr_height.set(qn("w:val"), str(max(int(min_height_twips), 1)))
        tr_height.set(qn("w:hRule"), "atLeast")
    tc = etree.SubElement(tr, qn("w:tc"))
    tc_pr = etree.SubElement(tc, qn("w:tcPr"))
    tc_w = etree.SubElement(tc_pr, qn("w:tcW"))
    tc_w.set(qn("w:w"), str(width_twips))
    tc_w.set(qn("w:type"), "dxa")
    _zero_cell_margins(tc_pr)
    _append_nil_borders(tc_pr, tag="w:tcBorders")

    for node in nodes:
        tc.append(node)
    # Word is much happier when every cell terminates with a paragraph.
    tc.append(_empty_zero_paragraph())
    return tbl


def _apply_candidate(
    src_docx: Path,
    dst_docx: Path,
    candidate: WrapCandidate,
    *,
    use_pdf_y: bool = True,
    source_pdf_path: Path | None = None,
) -> int | None:
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        _docx_to_workdir(src_docx, workdir)
        tree, body = _parse_document_xml(workdir)
        blocks = _collect_float_blocks(body)
        indices = _matching_block_indices(blocks, candidate.captions)
        if not indices:
            return None

        children = list(body)
        first = blocks[min(indices)]
        anchor = children[first.start_idx]
        inner_nodes = []
        removed = []
        for block_idx in indices:
            block = blocks[block_idx]
            inner_nodes.extend(children[block.start_idx + 1:block.end_idx])
            removed.extend(children[block.start_idx:block.end_idx + 1])
        if not inner_nodes:
            return None

        width_twips = _body_text_width_twips(body)
        desc = "skill-pdf-float-wrap"
        if candidate.source_page:
            desc += f"-p{candidate.source_page}"
        min_height_twips = _estimate_float_content_height_twips(inner_nodes)
        converted_drawings = _convert_drawings_to_square_anchors(inner_nodes)
        if converted_drawings <= 0:
            return None
        y_twips = None
        float_tbl = _floating_table(inner_nodes, width_twips, y_twips, desc, min_height_twips)
        anchor.addprevious(float_tbl)
        for node in removed:
            if node.getparent() is body:
                body.remove(node)

        tree.write(str(workdir / "word" / "document.xml"), encoding="utf-8", xml_declaration=True, standalone=True)
        _write_docx_from_workdir(workdir, dst_docx)
    return converted_drawings


def _score_acceptable(new_score: tuple[float, int, float], old_score: tuple[float, int, float]) -> bool:
    return _is_score_better(new_score, old_score)


def _accepts_trial(
    trial_score: tuple[float, int, float],
    current_score: tuple[float, int, float],
    *,
    force_all: bool,
    converted_drawings: int,
) -> bool:
    if force_all and converted_drawings <= 0:
        return False
    if force_all and converted_drawings > 0:
        return True
    return _score_acceptable(trial_score, current_score)


def wrap_cross_column_floats(
    docx_path: str | Path,
    guidance_items=None,
    *,
    source_pdf_path: str | Path | None = None,
    max_iterations: int = 10,
    threshold_pt: float = 160.0,
    keep_debug: bool = True,
    verify_render: bool = True,
    force_all: bool = False,
) -> dict:
    """Convert section-break full-width floats to editable floating containers."""
    docx_path = Path(docx_path).resolve()
    if not docx_path.exists():
        return {"enabled": False, "reason": "docx not found"}

    debug_dir = docx_path.with_suffix("")
    debug_dir = debug_dir.parent / f"{debug_dir.name}_float_wrap"
    debug_dir.mkdir(parents=True, exist_ok=True)

    current_docx = debug_dir / "current.docx"
    shutil.copy2(docx_path, current_docx)
    accepted: list[dict] = []
    source_pdf = Path(source_pdf_path).resolve() if source_pdf_path else None

    if verify_render:
        current_pdf = debug_dir / "current.pdf"
        if not export_docx_to_pdf(current_docx, current_pdf):
            return {"enabled": False, "reason": "docx to pdf export failed"}
        current_score = _blank_score(current_pdf, threshold_pt)
    else:
        current_score = (0.0, 0, 0.0)

    if not verify_render:
        candidates = _candidate_list(current_docx, guidance_items)
        for cand_idx, candidate in enumerate(candidates):
            trial_docx = debug_dir / f"once_{cand_idx + 1}.docx"
            converted_drawings = _apply_candidate(
                current_docx,
                trial_docx,
                candidate,
                use_pdf_y=True,
                source_pdf_path=source_pdf,
            )
            if converted_drawings is None:
                continue
            shutil.copy2(trial_docx, current_docx)
            accepted.append({
                "iteration": 1,
                "float_count": len(candidate.captions),
                "captions": candidate.captions,
                "source_page": candidate.source_page,
                "converted_drawings": converted_drawings,
                "score_before": current_score,
                "score_after": current_score,
            })
        if accepted:
            shutil.copy2(current_docx, docx_path)
        if not keep_debug:
            shutil.rmtree(debug_dir, ignore_errors=True)
        return {
            "enabled": True,
            "accepted": accepted,
            "final_score": current_score,
            "debug_dir": str(debug_dir),
            "one_pass": True,
        }

    for iteration in range(max_iterations):
        candidates = _candidate_list(current_docx, guidance_items)
        if not candidates:
            break
        accepted_this_round = False
        for cand_idx, candidate in enumerate(candidates):
            trial_docx = debug_dir / f"trial_{iteration + 1}_{cand_idx + 1}.docx"
            converted_drawings = _apply_candidate(
                current_docx,
                trial_docx,
                candidate,
                use_pdf_y=not force_all,
                source_pdf_path=source_pdf,
            )
            if converted_drawings is None:
                continue
            if not verify_render:
                shutil.copy2(trial_docx, current_docx)
                accepted.append({
                    "iteration": iteration + 1,
                    "float_count": len(candidate.captions),
                    "captions": candidate.captions,
                    "source_page": candidate.source_page,
                    "converted_drawings": converted_drawings,
                    "score_before": current_score,
                    "score_after": current_score,
                })
                accepted_this_round = True
                break

            trial_pdf = debug_dir / f"trial_{iteration + 1}_{cand_idx + 1}.pdf"
            if not export_docx_to_pdf(trial_docx, trial_pdf):
                continue
            trial_score = _blank_score(trial_pdf, threshold_pt)
            if not _accepts_trial(
                trial_score,
                current_score,
                force_all=force_all,
                converted_drawings=converted_drawings,
            ):
                continue
            shutil.copy2(trial_docx, current_docx)
            shutil.copy2(trial_pdf, current_pdf)
            accepted.append({
                "iteration": iteration + 1,
                "float_count": len(candidate.captions),
                "captions": candidate.captions,
                "source_page": candidate.source_page,
                "converted_drawings": converted_drawings,
                "score_before": current_score,
                "score_after": trial_score,
            })
            current_score = trial_score
            accepted_this_round = True
            break
        if not accepted_this_round:
            break

    if accepted:
        shutil.copy2(current_docx, docx_path)
    if not keep_debug:
        shutil.rmtree(debug_dir, ignore_errors=True)
    return {
        "enabled": True,
        "accepted": accepted,
        "final_score": current_score,
        "debug_dir": str(debug_dir),
    }
