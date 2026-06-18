"""PDF-derived placement hints for editable Word floats.

The hints are intentionally advisory: they never replace editable Word
figures/tables with PDF screenshots.  They only tell the Word insertion layer
whether LaTeX placed a cross-column float near the page top or inline.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
try:
    from shared.latex_text_utils import clean_latex_text
except Exception:  # pragma: no cover - fallback for isolated imports
    def clean_latex_text(text):
        return re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?(?:\{([^{}]*)\})?", r"\1", text or "")


def apply_pdf_float_guidance(tex_path, images, tikz_tables):
    """Attach PDF placement hints to extracted figure/table metadata.

    Returns a small stats dict.  Missing PDFs or PyMuPDF simply disable hints.
    """
    pdf_path = _find_pdf_for_tex(tex_path)
    if pdf_path is None:
        return {"enabled": False, "reason": "pdf-not-found", "matched": 0}

    try:
        import fitz
    except Exception as exc:
        return {"enabled": False, "reason": f"pymupdf-unavailable: {exc}", "matched": 0}

    try:
        pdf_doc = fitz.open(str(pdf_path))
    except Exception as exc:
        return {"enabled": False, "reason": f"pdf-open-failed: {exc}", "matched": 0}

    try:
        try:
            tex_content = Path(tex_path).read_text(encoding="utf-8", errors="ignore")
        except Exception:
            tex_content = ""
        used_hits = set()
        guided_items = list(_iter_guided_items(images, tikz_tables))
        matched = 0
        for record in guided_items:
            guidance = _find_item_guidance(pdf_doc, record, used_hits)
            if guidance:
                _attach_delay_after_guidance(pdf_doc, tex_content, record["item"], guidance)
                item = record["item"]
                item["pdf_guidance"] = guidance
                matched += 1
        return {
            "enabled": True,
            "pdf": str(pdf_path),
            "matched": matched,
            "total": len(guided_items),
            "top": sum(1 for record in guided_items
                       if record["item"].get("pdf_guidance", {}).get("position") == "top"),
            "inline": sum(1 for record in guided_items
                          if record["item"].get("pdf_guidance", {}).get("position") == "inline"),
            "delayed": sum(1 for record in guided_items
                           if record["item"].get("pdf_guidance", {}).get("delay_after")),
        }
    finally:
        pdf_doc.close()


def _find_pdf_for_tex(tex_path):
    tex = Path(tex_path)
    candidates = [
        tex.with_suffix(".pdf"),
        tex.parent / f"{tex.stem}.pdf",
    ]
    candidates.extend(sorted(tex.parent.glob("*.pdf")))
    return next((path for path in candidates if path.exists()), None)


def _iter_guided_items(images, tikz_tables):
    for index, item in enumerate(images or []):
        if item.get("is_full_width"):
            yield {"kind": "figure", "index": index, "item": item}
    for index, item in enumerate(tikz_tables or []):
        if _is_full_width_table_item(item):
            yield {"kind": "table", "index": index, "item": item}


def _is_full_width_table_item(item):
    if item.get("is_full_width"):
        return True
    tikz_body = str(item.get("tikz_body") or "")
    return "% meta:full_width=1" in tikz_body


def _find_item_guidance(pdf_doc, record, used_hits):
    for source, query in _search_queries(record):
        for page_index in range(pdf_doc.page_count):
            page = pdf_doc[page_index]
            try:
                hits = page.search_for(query)
            except Exception:
                hits = []
            for hit_index, hit in enumerate(hits):
                if source == "label" and not _hit_line_starts_with(page, hit, query):
                    continue
                key = (page_index, round(hit.y0, 1), hit_index)
                if key in used_hits:
                    continue
                used_hits.add(key)
                return _guidance_from_hit(page, page_index, hit, query, source)
    return None


def _search_queries(record):
    item = record["item"]
    caption = _caption_query(item)
    for query in _query_variants(caption):
        yield "caption", query
    for label in _label_queries(record["kind"], item, record["index"]):
        yield "label", label


def _label_queries(kind, item, index):
    number = str(item.get("number") or index + 1).strip()
    if not number:
        return []
    if kind == "table":
        prefixes = ("Table", "Tab.", "表", "表 ")
    else:
        prefixes = ("Figure", "Fig.", "图", "图 ")
    return [f"{prefix} {number}" if prefix.isascii() else f"{prefix}{number}" for prefix in prefixes]


def _caption_query(item):
    caption = clean_latex_text(item.get("caption") or "")
    caption = re.sub(r"\s+", " ", caption).strip()
    return caption


def _query_variants(text):
    text = re.sub(r"\s+", " ", text or "").strip()
    if not text:
        return []
    variants = []
    for length in (100, 80, 60, 45, 32, 24):
        if len(text) >= length:
            variants.append(text[:length])
    variants.append(text)
    seen = set()
    return [value for value in variants if value and not (value in seen or seen.add(value))]


def _guidance_from_hit(page, page_index, caption_rect, query, source):
    page_h = float(page.rect.height or 1.0)
    y0 = _float_top_from_pdf_blocks(page, caption_rect)
    y0_ratio = max(0.0, min(1.0, y0 / page_h))
    caption_y_ratio = max(0.0, min(1.0, caption_rect.y0 / page_h))
    position = _classify_position(y0_ratio, caption_y_ratio)
    return {
        "page": page_index + 1,
        "y0_pt": round(float(y0), 2),
        "caption_y_pt": round(float(caption_rect.y0), 2),
        "y0_ratio": round(y0_ratio, 4),
        "caption_y_ratio": round(caption_y_ratio, 4),
        "position": position,
        "query": query,
        "source": source,
    }


def _classify_position(y0_ratio, caption_y_ratio):
    if y0_ratio <= 0.24:
        return "top"
    if y0_ratio >= 0.32:
        return "inline"
    return "inline" if caption_y_ratio >= 0.55 else "top"


def _hit_line_starts_with(page, hit, query):
    line = _line_text_for_hit(page, hit)
    return _normalize_search_text(line).startswith(_normalize_search_text(query))


def _line_text_for_hit(page, hit):
    try:
        blocks = page.get_text("dict").get("blocks", [])
    except Exception:
        return ""
    for block in blocks:
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            x0 = min(float(span["bbox"][0]) for span in spans)
            y0 = min(float(span["bbox"][1]) for span in spans)
            x1 = max(float(span["bbox"][2]) for span in spans)
            y1 = max(float(span["bbox"][3]) for span in spans)
            if x0 <= hit.x1 and x1 >= hit.x0 and y0 <= hit.y1 and y1 >= hit.y0:
                return "".join(span.get("text", "") for span in spans)
    return ""


def _float_top_from_pdf_blocks(page, caption_rect):
    nearest_above = []
    try:
        blocks = page.get_text("dict").get("blocks", [])
    except Exception:
        return float(caption_rect.y0)

    page_width = float(page.rect.width or 1.0)
    max_gap = float(page.rect.height or 1.0) * 0.75
    for block in blocks:
        bbox = block.get("bbox")
        if not bbox or len(bbox) != 4:
            continue
        x0, y0, x1, y1 = [float(value) for value in bbox]
        if y1 > caption_rect.y0 + 8 or y1 < caption_rect.y0 - max_gap:
            continue
        width_ratio = (x1 - x0) / page_width
        overlaps_caption = x1 >= caption_rect.x0 and x0 <= caption_rect.x1
        if block.get("type") == 1 or width_ratio >= 0.45 or overlaps_caption:
            nearest_above.append((abs(caption_rect.y0 - y1), y0))
    if not nearest_above:
        return float(caption_rect.y0)
    return min(nearest_above, key=lambda item: item[0])[1]


def _attach_delay_after_guidance(pdf_doc, tex_content, item, guidance):
    if guidance.get("position") != "top":
        return
    float_page_index = int(guidance.get("page", 1)) - 1
    if float_page_index <= 0 or not tex_content:
        return
    delay_after = _find_latest_text_before_float_page(
        pdf_doc, tex_content, int(item.get("end") or 0), float_page_index)
    if delay_after:
        guidance["delay_after"] = delay_after


def _find_latest_text_before_float_page(pdf_doc, tex_content, source_end, float_page_index):
    latest = None
    for candidate in _following_text_candidates(tex_content, source_end):
        hit = _find_text_before_page(pdf_doc, candidate["text"], float_page_index)
        if hit:
            latest = {**candidate, **hit}
    return latest


def _following_text_candidates(tex_content, source_end, max_candidates=35):
    tail = tex_content[source_end:source_end + 18000]
    collected = 0
    for chunk in re.split(r"\n\s*\n+", tail):
        if re.search(r"\\begin\{(?:figure|table|strip)\*?\}", chunk):
            break
        text = _latex_chunk_to_text(chunk)
        if len(_normalize_search_text(text)) < 18:
            continue
        collected += 1
        yield {
            "text": text[:160],
            "query": _body_query_variants(text)[0],
            "source_order": collected,
        }
        if collected >= max_candidates:
            break


def _latex_chunk_to_text(chunk):
    chunk = re.sub(r"(?<!\\)%.*", " ", chunk or "")
    if re.search(r"\\begin\{(?:equation|align|gather|tikzpicture|tabular)", chunk):
        return ""
    chunk = re.sub(r"\\(?:label|cite\w*|ref|eqref|url)\{[^{}]*\}", " ", chunk)
    chunk = re.sub(r"\$[^$]*\$", " ", chunk)
    text = clean_latex_text(chunk)
    text = re.sub(r"[{}\\]", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _find_text_before_page(pdf_doc, text, float_page_index):
    for query in _body_query_variants(text):
        for page_index in range(float_page_index):
            page = pdf_doc[page_index]
            try:
                hits = page.search_for(query)
            except Exception:
                hits = []
            if hits:
                return {
                    "page": page_index + 1,
                    "query": query,
                    "y0_ratio": round(float(hits[0].y0) / float(page.rect.height or 1), 4),
                }
            if _normalized_text_on_page(page, query):
                return {
                    "page": page_index + 1,
                    "query": query,
                    "y0_ratio": None,
                    "match": "normalized-page-text",
                }
    return None


def _normalized_text_on_page(page, text):
    try:
        page_norm = _normalize_search_text(page.get_text("text"))
    except Exception:
        return False
    query_norm = _normalize_search_text(text)
    if len(query_norm) < 18:
        return False
    for size in (60, 45, 32, 24, 18):
        if len(query_norm) >= size and query_norm[:size] in page_norm:
            return True
    window = 24
    max_start = min(max(len(query_norm) - window, 0), 96)
    for start in range(0, max_start + 1, 8):
        if query_norm[start:start + window] in page_norm:
            return True
    return False


def _body_query_variants(text):
    text = re.sub(r"\s+", " ", text or "").strip()
    variants = []
    for length in (80, 60, 45, 32, 24):
        if len(text) >= length:
            variants.append(text[:length])
    if not variants and text:
        variants.append(text)
    return variants


def _normalize_search_text(text):
    return re.sub(r"[\W_]+", "", text or "", flags=re.UNICODE).lower()
