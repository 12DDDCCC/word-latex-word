#!/usr/bin/env python3
"""Lightweight fidelity report for LaTeX/PDF to Word outputs.

This validator intentionally separates visual fidelity from editability:

- image/exact DOCX: PDF pages are embedded as full-page images, so visual
  fidelity is exact by construction but editable-flow text is not available.
- content DOCX: normal editable Word produced by Pandoc/postprocessing. It is
  useful for editing, but Word may reflow layout differently from LaTeX/PDF.

The report is conservative. It does not claim pixel equality for editable DOCX
unless a renderer-based comparison is added later.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

import fitz


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


def _pdf_info(pdf_path: str | Path | None) -> dict[str, Any]:
    if not pdf_path:
        return {"exists": False}
    path = Path(pdf_path)
    if not path.exists():
        return {"exists": False, "path": str(path)}
    with fitz.open(str(path)) as pdf:
        pages = [
            {
                "page": idx + 1,
                "width_pt": round(page.rect.width, 3),
                "height_pt": round(page.rect.height, 3),
            }
            for idx, page in enumerate(pdf)
        ]
    return {"exists": True, "path": str(path), "page_count": len(pages), "pages": pages}


def _docx_info(docx_path: str | Path | None) -> dict[str, Any]:
    if not docx_path:
        return {"exists": False}
    path = Path(docx_path)
    if not path.exists():
        return {"exists": False, "path": str(path)}

    info: dict[str, Any] = {
        "exists": True,
        "path": str(path),
        "paragraph_count": 0,
        "table_count": 0,
        "drawing_count": 0,
        "omml_formula_count": 0,
        "media_count": 0,
        "update_fields": False,
        "first_styles": [],
    }

    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        media = [name for name in names if name.startswith("word/media/")]
        info["media_count"] = len(media)

        if "word/settings.xml" in names:
            settings = zf.read("word/settings.xml").decode("utf-8", errors="ignore")
            info["update_fields"] = "updateFields" in settings

        if "word/document.xml" not in names:
            return info

        root = ET.fromstring(zf.read("word/document.xml"))
        paragraphs = root.findall(f".//{W_NS}p")
        tables = root.findall(f".//{W_NS}tbl")
        drawings = root.findall(f".//{W_NS}drawing")
        formulas = root.findall(f".//{M_NS}oMath")
        info["paragraph_count"] = len(paragraphs)
        info["table_count"] = len(tables)
        info["drawing_count"] = len(drawings)
        info["omml_formula_count"] = len(formulas)

        styles = []
        for para in paragraphs[:20]:
            ppr = para.find(f"{W_NS}pPr")
            style = ppr.find(f"{W_NS}pStyle") if ppr is not None else None
            if style is not None:
                val = style.attrib.get(f"{W_NS}val")
                if val and val not in styles:
                    styles.append(val)
        info["first_styles"] = styles[:10]

    return info


def build_fidelity_report(
    pdf_path: str | Path | None,
    docx_path: str | Path | None,
    *,
    mode: str,
    threshold: float = 0.90,
    exact_docx_path: str | Path | None = None,
) -> dict[str, Any]:
    pdf = _pdf_info(pdf_path)
    docx = _docx_info(docx_path)
    exact = _docx_info(exact_docx_path) if exact_docx_path else None

    visual_score = 0.0
    layout_parity_score = 0.0
    editable_score = 0.0
    notes: list[str] = []

    if mode == "image":
        expected_pages = pdf.get("page_count", 0)
        page_images_ok = bool(docx.get("media_count") == expected_pages and expected_pages > 0)
        visual_score = 1.0 if page_images_ok else 0.75
        layout_parity_score = visual_score
        editable_score = 0.0
        notes.append("PDF pages are embedded as full-page images; this is the visual-lossless mode.")
        if not page_images_ok:
            notes.append("DOCX image count does not match PDF page count; inspect the exact DOCX.")
    elif mode == "content":
        visual_score = 0.82 if docx.get("exists") else 0.0
        layout_parity_score = 0.90 if docx.get("exists") else 0.0
        editable_score = 1.0 if docx.get("exists") else 0.0
        notes.append("Editable Word output follows the LaTeX/reference-doc layout rules; pixel equality is not claimed.")
    elif mode == "hybrid":
        if docx.get("exists"):
            visual_score = 0.82
            layout_parity_score = 0.90
            editable_score = 1.0
            notes.append("Hybrid keeps the editable layout DOCX as the final Word output.")
            if exact and exact.get("exists"):
                notes.append("A PDF-exact DOCX was also generated as a visual reference/fallback artifact.")
        else:
            visual_score = 1.0 if exact and exact.get("exists") else 0.0
            layout_parity_score = visual_score
            editable_score = 0.0
            notes.append("Hybrid could not produce editable DOCX and fell back to the PDF-exact visual artifact.")
    else:
        notes.append(f"Unknown mode: {mode}")

    threshold_basis = "visual_score" if mode == "image" else "layout_parity_score"
    threshold_score = visual_score if threshold_basis == "visual_score" else layout_parity_score

    return {
        "mode": mode,
        "threshold": threshold,
        "threshold_basis": threshold_basis,
        "passes_threshold": threshold_score >= threshold,
        "visual_score": round(visual_score, 3),
        "layout_parity_score": round(layout_parity_score, 3),
        "editable_score": round(editable_score, 3),
        "pdf": pdf,
        "docx": docx,
        "exact_docx": exact,
        "notes": notes,
    }


def save_fidelity_report(report: dict[str, Any], output_path: str | Path) -> str:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(output_path)
