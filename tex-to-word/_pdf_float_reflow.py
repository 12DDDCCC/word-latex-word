"""PDF-guided cross-column float reflow for generated DOCX files.

This module is intentionally conservative:
- It moves editable figure/table XML blocks, never PDF screenshots.
- It only accepts a move after Word re-renders the trial DOCX to PDF and
  the blank-space score improves.
- It is opt-in from the CLI because it requires Microsoft Word automation.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import os
import re
import shutil
import subprocess
import tempfile

import fitz
from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(name: str) -> str:
    prefix, local = name.split(":")
    if prefix != "w":
        raise ValueError(name)
    return f"{{{W_NS}}}{local}"


def _norm(text: str) -> str:
    return re.sub(r"[\W_]+", "", text or "", flags=re.UNICODE).lower()


def _text_of(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(qn("w:t"))).strip()


def _is_caption_like(text: str) -> bool:
    text = text or ""
    lead = r"^\s*(?:[\u2022\u25aa\u25cf\-]\s*)?"
    english = lead + r"(?:Figure|Fig\.|Table)\s*\d+(?:\.\d+)*(?!\.\d)\s*(?:[.:;:\uff1a\uff1b\u3001\uff0c,]|\s|$)"
    chinese = lead + r"[\u56fe\u8868]\s*\d+(?:\.\d+)*(?!\.\d)\s*(?:[.\u3002:\uff1a;\uff1b\u3001\uff0c,]|\s|$)"
    return re.match(english, text, re.I) is not None or re.match(chinese, text) is not None


def _heading_major(text: str) -> str:
    lead = r"^\s*(?:[\u2022\u25aa\u25cf\-]\s*)?"
    match = re.match(lead + r"(\d+)(?:\.\d+)*\s+\S+", text or "")
    return match.group(1) if match else ""


def _figure_major(text: str) -> str:
    lead = r"^\s*(?:[\u2022\u25aa\u25cf\-]\s*)?"
    match = re.match(lead + r"(?:Figure|Fig\.|[\u56fe])\s*(\d+)(?:\.\d+)*", text or "", re.I)
    return match.group(1) if match else ""


def _captions_match_section(captions: list[str], section_anchor: str) -> bool:
    section_major = _heading_major(section_anchor)
    if not section_major:
        return True
    figure_majors = {_figure_major(caption) for caption in captions}
    figure_majors.discard("")
    return not figure_majors or all(major == section_major for major in figure_majors)


def _section_column_count(elem) -> int | None:
    sect = elem.find("w:pPr/w:sectPr", NS)
    if sect is None:
        return None
    cols = sect.find("w:cols", NS)
    if cols is None:
        return 1
    try:
        return int(cols.get(qn("w:num")) or 1)
    except (TypeError, ValueError):
        return 1


def _is_empty_section_para(elem) -> bool:
    return (
        elem.tag == qn("w:p")
        and elem.find("w:pPr/w:sectPr", NS) is not None
        and not _text_of(elem)
    )


def _has_drawing(elem) -> bool:
    return elem.find(".//w:drawing", NS) is not None


def _paragraph_style_id(elem) -> str:
    p_style = elem.find("w:pPr/w:pStyle", NS)
    return (p_style.get(qn("w:val")) or "") if p_style is not None else ""


def _is_heading_like(elem) -> bool:
    if elem.tag != qn("w:p"):
        return False
    text = _text_of(elem)
    if not text or _is_caption_like(text):
        return False
    style_id = _paragraph_style_id(elem).lower()
    if style_id.startswith("heading") or "section" in style_id:
        return True
    lead = r"^\s*(?:[\u2022\u25aa\u25cf\-]\s*)?"
    numbered = lead + r"\d+(?:\.\d+)*\s+\S+"
    chinese = lead + r"[一二三四五六七八九十]+[、.．]\s*\S+"
    return re.match(numbered, text) is not None or re.match(chinese, text) is not None


def _is_section_body_anchor(elem) -> bool:
    if elem.tag != qn("w:p"):
        return False
    if _is_empty_section_para(elem) or _has_drawing(elem):
        return False
    text = _text_of(elem)
    if not text or "[[" in text or _is_caption_like(text) or _is_heading_like(elem):
        return False
    return len(_norm(text)) >= 12


def _remove_page_break_before(p_elem) -> None:
    ppr = p_elem.find("w:pPr", NS)
    if ppr is None:
        return
    for pbb in list(ppr.findall("w:pageBreakBefore", NS)):
        ppr.remove(pbb)


@dataclass
class FloatBlock:
    start_idx: int
    end_idx: int
    caption: str
    section_anchor: str = ""
    section_floor: str = ""


@dataclass
class PageInfo:
    index: int
    blank_pt: float
    first_text: str


@dataclass
class ReflowCandidate:
    blank_page: int
    blank_pt: float
    anchor_text: str
    captions: list[str]
    min_page: int = 0
    section_anchor: str = ""


def _docx_to_workdir(docx_path: Path, workdir: Path) -> None:
    with ZipFile(docx_path) as zf:
        zf.extractall(workdir)


def _write_docx_from_workdir(workdir: Path, docx_path: Path) -> None:
    docx_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(docx_path, "w", ZIP_DEFLATED) as out:
        for path in workdir.rglob("*"):
            if path.is_file():
                out.write(path, path.relative_to(workdir).as_posix())


def _parse_document_xml(workdir: Path):
    parser = etree.XMLParser(remove_blank_text=False)
    tree = etree.parse(str(workdir / "word" / "document.xml"), parser)
    body = tree.getroot().find("w:body", NS)
    return tree, body


def _float_caption(nodes: list) -> str:
    for elem in nodes:
        text = _text_of(elem)
        if _is_caption_like(text):
            return text
    for elem in nodes:
        text = _text_of(elem)
        if text:
            return text
    return ""


def _collect_float_blocks(body) -> list[FloatBlock]:
    children = list(body)
    blocks: list[FloatBlock] = []
    last_heading = ""
    section_floor = ""
    headings_by_major: dict[str, str] = {}
    floors_by_major: dict[str, str] = {}
    idx = 0
    while idx < len(children):
        elem = children[idx]
        if _is_heading_like(elem):
            last_heading = _text_of(elem)
            section_floor = ""
            major = _heading_major(last_heading)
            if major:
                headings_by_major[major] = last_heading
                floors_by_major[major] = ""
        elif last_heading and not section_floor and _is_section_body_anchor(elem):
            section_floor = _text_of(elem)
            major = _heading_major(last_heading)
            if major:
                floors_by_major[major] = section_floor
        if not _is_empty_section_para(elem) or (_section_column_count(elem) or 1) < 2:
            idx += 1
            continue

        end = idx + 1
        saw_float = False
        section_anchor = last_heading
        section_floor_anchor = section_floor
        while end < len(children):
            cur = children[end]
            if end > idx + 1 and _is_empty_section_para(cur):
                if saw_float:
                    nodes = children[idx : end + 1]
                    caption = _float_caption(nodes)
                    resolved_anchor = section_anchor
                    resolved_floor = section_floor_anchor
                    figure_major = _figure_major(caption)
                    if figure_major and _heading_major(resolved_anchor) != figure_major:
                        resolved_anchor = headings_by_major.get(figure_major, resolved_anchor)
                        resolved_floor = floors_by_major.get(figure_major, resolved_floor)
                    blocks.append(
                        FloatBlock(
                            idx,
                            end,
                            caption,
                            resolved_anchor,
                            resolved_floor,
                        )
                    )
                break
            if cur.tag == qn("w:tbl") or _has_drawing(cur) or _is_caption_like(_text_of(cur)):
                saw_float = True
            end += 1
        idx = max(end, idx + 1)
    return blocks


def _pdf_pages(pdf_path: Path) -> list[PageInfo]:
    pages: list[PageInfo] = []
    with fitz.open(str(pdf_path)) as doc:
        for page_index, page in enumerate(doc):
            height = page.rect.height
            blocks = []
            for block in page.get_text("blocks"):
                x0, y0, x1, y1, text, *_ = block
                text = " ".join((text or "").split())
                if not text:
                    continue
                if y0 > height * 0.88 and len(text) <= 8:
                    continue
                blocks.append((x0, y0, x1, y1, text))
            bottom = max((b[3] for b in blocks), default=0.0)
            first_text = ""
            if blocks:
                top_blocks = sorted(blocks, key=lambda b: (b[1], b[0]))
                first_text = top_blocks[0][4]
            pages.append(PageInfo(page_index, height - bottom if bottom else height, first_text))
    return pages


def _text_page(pdf_path: Path, text: str, *, min_len: int = 10) -> int | None:
    needle = _norm(text)[:56]
    if len(needle) < min_len:
        return None
    with fitz.open(str(pdf_path)) as doc:
        for idx, page in enumerate(doc):
            page_text = _norm(page.get_text("text"))
            for size in (56, 44, 34, 26, 18, 12, 10, 8, 6, 5, 4):
                if size >= min_len and len(needle) >= size and needle[:size] in page_text:
                    return idx
    return None


def _caption_page(pdf_path: Path, caption: str) -> int | None:
    return _text_page(pdf_path, caption, min_len=10)


def _section_page(pdf_path: Path, section_anchor: str) -> int | None:
    return _text_page(pdf_path, section_anchor, min_len=4)


def _block_min_page(pdf_path: Path, block: FloatBlock) -> int | None:
    if block.section_anchor:
        page = _section_page(pdf_path, block.section_anchor)
        if page is not None:
            return page
        if block.section_floor:
            page = _text_page(pdf_path, block.section_floor, min_len=12)
            if page is not None:
                return page
        return None
    return 0


def _find_anchor_text_for_page(pdf_path: Path, page_index: int) -> str:
    pages = _pdf_pages(pdf_path)
    if not (0 <= page_index < len(pages)):
        return ""
    return pages[page_index].first_text


def _find_paragraph_by_text(body, text: str):
    needle = _norm(text)
    min_size = 4 if re.match(r"^\d+(?:\.\d+)*", text.strip()) else 8
    if len(needle) < min_size:
        return None
    children = list(body)
    for size in (80, 60, 42, 30, 20, 12, 8, 6, 5, 4):
        if size < min_size or len(needle) < size:
            continue
        piece = needle[:size]
        for elem in children:
            if elem.tag != qn("w:p"):
                continue
            if piece in _norm(_text_of(elem)):
                return elem
    return None


def _is_page_top_float(page: PageInfo) -> bool:
    return _is_caption_like(page.first_text)


def _float_starts_page(page: PageInfo, captions: list[str]) -> bool:
    first = _norm(page.first_text)
    if len(first) < 8:
        return False
    for caption in captions:
        cap = _norm(caption)
        if cap and (cap.startswith(first[:16]) or first.startswith(cap[:16])):
            return True
    return False


def _group_consecutive_blocks(
    blocks: list[FloatBlock],
    start: int,
    page: int,
    pdf_path: Path,
    *,
    max_size: int = 3,
) -> list[FloatBlock]:
    group = [blocks[start]]
    last_page = page
    idx = start + 1
    while idx < len(blocks):
        prev = group[-1]
        cur = blocks[idx]
        # Consecutive wrapped floats look like: previous end section, next start section.
        if cur.start_idx > prev.end_idx + 2:
            break
        if cur.section_anchor != group[0].section_anchor:
            break
        base_major = _figure_major(group[0].caption)
        cur_major = _figure_major(cur.caption)
        if base_major and cur_major and base_major != cur_major:
            break
        cur_page = _caption_page(pdf_path, cur.caption)
        if cur_page is None or cur_page < page or cur_page > last_page + 1:
            break
        group.append(cur)
        last_page = max(last_page, cur_page)
        if len(group) >= max_size:
            break
        idx += 1
    return group


def _candidate_list(docx_path: Path, pdf_path: Path, threshold_pt: float) -> list[ReflowCandidate]:
    pages = _pdf_pages(pdf_path)
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        _docx_to_workdir(docx_path, workdir)
        _, body = _parse_document_xml(workdir)
        blocks = _collect_float_blocks(body)

    block_pages = [_caption_page(pdf_path, block.caption) for block in blocks]
    block_min_pages = []
    for block, block_page in zip(blocks, block_pages):
        min_page = _block_min_page(pdf_path, block)
        if min_page is None:
            min_page = block_page if block_page is not None else len(pages)
        block_min_pages.append(min_page)
    candidates: list[ReflowCandidate] = []
    seen: set[tuple[int, tuple[str, ...]]] = set()

    def add_candidate(
        target_page: int,
        metric_blank: float,
        variant: list[FloatBlock],
        *,
        prefer_section_anchor: bool = False,
    ) -> None:
        if not (0 <= target_page < len(pages)):
            return
        captions = [g.caption for g in variant if g.caption]
        if not captions:
            return
        if _is_page_top_float(pages[target_page]):
            return
        indices = [blocks.index(g) for g in variant]
        min_page = max(block_min_pages[i] for i in indices) if indices else 0
        if target_page < min_page:
            return
        section_anchor = next((g.section_anchor for g in reversed(variant) if g.section_anchor), "")
        if not _captions_match_section(captions, section_anchor):
            return
        anchor_text = section_anchor if prefer_section_anchor and section_anchor else ""
        if not anchor_text:
            anchor_text = _find_anchor_text_for_page(pdf_path, target_page)
        if not anchor_text:
            return
        key = (target_page, tuple(_norm(c)[:56] for c in captions), (_norm(anchor_text)[:40],))
        if key in seen:
            return
        seen.add(key)
        candidates.append(
            ReflowCandidate(
                target_page,
                metric_blank,
                anchor_text,
                captions,
                min_page=min_page,
                section_anchor=section_anchor,
            )
        )

    for block_idx, block in enumerate(blocks):
        block_page = block_pages[block_idx]
        if block_page is None or block_page <= 0:
            continue
        prev_page = block_page - 1

        group = _group_consecutive_blocks(blocks, block_idx, block_page, pdf_path)
        variants = [group[:size] for size in range(len(group), 0, -1)]
        prev_blank = pages[prev_page].blank_pt
        own_blank = pages[block_page].blank_pt
        starts_page = _float_starts_page(pages[block_page], [block.caption])
        should_try = (
            prev_blank >= threshold_pt
            or own_blank >= threshold_pt
            or not starts_page
        )
        if not should_try:
            continue
        metric_blank = max(prev_blank, own_blank)
        if prev_page >= block_min_pages[block_idx]:
            for variant in variants:
                variant_min_page = max(block_min_pages[blocks.index(g)] for g in variant)
                if prev_page == variant_min_page:
                    add_candidate(prev_page, metric_blank, variant, prefer_section_anchor=True)
                add_candidate(prev_page, metric_blank, variant)
        if (own_blank >= threshold_pt or not starts_page) and block_page >= block_min_pages[block_idx]:
            for variant in variants:
                add_candidate(block_page, own_blank, variant, prefer_section_anchor=True)
                add_candidate(block_page, own_blank, variant, prefer_section_anchor=False)

    candidates.sort(key=lambda c: c.blank_pt, reverse=True)
    return candidates


def _matching_block_indices(blocks: list[FloatBlock], captions: list[str]) -> list[int]:
    targets = [_norm(caption)[:56] for caption in captions]
    matched: list[int] = []
    for target in targets:
        found = None
        for idx, block in enumerate(blocks):
            block_norm = _norm(block.caption)
            if target and (block_norm.startswith(target[:24]) or target[:24] in block_norm):
                found = idx
                break
        if found is None:
            return []
        matched.append(found)
    return matched


def _apply_candidate(src_docx: Path, dst_docx: Path, candidate: ReflowCandidate) -> bool:
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        _docx_to_workdir(src_docx, workdir)
        tree, body = _parse_document_xml(workdir)
        blocks = _collect_float_blocks(body)
        indices = _matching_block_indices(blocks, candidate.captions)
        if not indices:
            return False
        anchor = _find_paragraph_by_text(body, candidate.anchor_text)
        if anchor is None:
            return False

        children = list(body)
        anchor_idx = children.index(anchor)
        if candidate.section_anchor:
            section_anchor = _find_paragraph_by_text(body, candidate.section_anchor)
            if section_anchor is not None and anchor_idx < children.index(section_anchor):
                return False

        nodes = []
        for idx in indices:
            block = blocks[idx]
            nodes.extend(children[block.start_idx : block.end_idx + 1])
        if anchor in nodes:
            return False

        for node in nodes:
            if node.getparent() is body:
                body.remove(node)
        insert_at = list(body).index(anchor) + (1 if _is_heading_like(anchor) else 0)
        for offset, node in enumerate(nodes):
            body.insert(insert_at + offset, node)
        if nodes:
            _remove_page_break_before(nodes[0])

        tree.write(str(workdir / "word" / "document.xml"), encoding="utf-8", xml_declaration=True, standalone=True)
        _write_docx_from_workdir(workdir, dst_docx)
    return True


def _blank_score(pdf_path: Path, threshold_pt: float = 160.0) -> tuple[float, int, float]:
    pages = _pdf_pages(pdf_path)
    scored = pages[:-1] if len(pages) > 1 else pages
    excess = [max(0.0, page.blank_pt - threshold_pt) for page in scored]
    return (sum(excess), sum(1 for item in excess if item > 0), max(excess, default=0.0))


def _is_score_better(new_score: tuple[float, int, float], old_score: tuple[float, int, float]) -> bool:
    if new_score[1] < old_score[1]:
        return new_score[0] <= old_score[0] + 40 and new_score[2] <= old_score[2] + 40
    if new_score[1] == old_score[1] and new_score[0] < old_score[0] - 10:
        return True
    if new_score[1] == old_score[1] and new_score[2] < old_score[2] - 15:
        return True
    return False


def _caption_starts_page(pdf_path: Path, caption: str) -> bool | None:
    page_index = _caption_page(pdf_path, caption)
    if page_index is None:
        return None
    pages = _pdf_pages(pdf_path)
    if not (0 <= page_index < len(pages)):
        return None
    return _float_starts_page(pages[page_index], [caption])


def _alignment_improved(current_pdf: Path, trial_pdf: Path, candidate: ReflowCandidate) -> bool:
    for caption in candidate.captions:
        before = _caption_starts_page(current_pdf, caption)
        after = _caption_starts_page(trial_pdf, caption)
        if before is False and after is True:
            return True
    return False


def _is_alignment_safe(
    new_score: tuple[float, int, float],
    old_score: tuple[float, int, float],
) -> bool:
    return new_score[1] <= old_score[1] and new_score[0] <= old_score[0] + 5


def export_docx_to_pdf(docx_path: Path, pdf_path: Path) -> bool:
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    script = Path(__file__).with_name("_docx2pdf.ps1")
    if not script.exists():
        return False
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    powershell_candidates = []
    for name in ("pwsh", "powershell"):
        found = shutil.which(name)
        if found:
            powershell_candidates.append(Path(found))
    system_root = os.environ.get("SystemRoot")
    if system_root:
        powershell_candidates.append(
            Path(system_root) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
        )
    powershell_candidates.append(Path("powershell.exe"))
    powershell = next((candidate for candidate in powershell_candidates if candidate.exists()), powershell_candidates[-1])
    cmd = [
        str(powershell),
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-Docx",
        str(docx_path),
        "-Pdf",
        str(pdf_path),
    ]
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180,
    )
    return result.returncode == 0 and pdf_path.exists()


def reflow_cross_column_floats(
    docx_path: str | Path,
    *,
    max_iterations: int = 10,
    threshold_pt: float = 160.0,
    keep_debug: bool = True,
) -> dict:
    """Try moving next-page top float groups onto the preceding blank page."""
    docx_path = Path(docx_path).resolve()
    if not docx_path.exists():
        return {"enabled": False, "reason": "docx not found"}

    debug_dir = docx_path.with_suffix("")
    debug_dir = debug_dir.parent / f"{debug_dir.name}_float_reflow"
    debug_dir.mkdir(parents=True, exist_ok=True)

    current_docx = debug_dir / "current.docx"
    shutil.copy2(docx_path, current_docx)
    accepted: list[dict] = []

    current_pdf = debug_dir / "current.pdf"
    if not export_docx_to_pdf(current_docx, current_pdf):
        return {"enabled": False, "reason": "docx to pdf export failed"}
    current_score = _blank_score(current_pdf, threshold_pt)

    for iteration in range(max_iterations):
        candidates = _candidate_list(current_docx, current_pdf, threshold_pt)
        if not candidates:
            break
        accepted_this_round = False
        for cand_idx, candidate in enumerate(candidates):
            trial_docx = debug_dir / f"trial_{iteration + 1}_{cand_idx + 1}.docx"
            trial_pdf = debug_dir / f"trial_{iteration + 1}_{cand_idx + 1}.pdf"
            if not _apply_candidate(current_docx, trial_docx, candidate):
                continue
            if not export_docx_to_pdf(trial_docx, trial_pdf):
                continue
            trial_score = _blank_score(trial_pdf, threshold_pt)
            alignment_improved = _alignment_improved(current_pdf, trial_pdf, candidate)
            if not (
                _is_score_better(trial_score, current_score)
                or (alignment_improved and _is_alignment_safe(trial_score, current_score))
            ):
                continue
            shutil.copy2(trial_docx, current_docx)
            shutil.copy2(trial_pdf, current_pdf)
            accepted.append({
                "iteration": iteration + 1,
                "blank_page": candidate.blank_page + 1,
                "blank_pt": round(candidate.blank_pt, 1),
                "float_count": len(candidate.captions),
                "min_page": candidate.min_page + 1,
                "section": candidate.section_anchor,
                "alignment_improved": alignment_improved,
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
