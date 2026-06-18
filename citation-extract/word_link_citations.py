#!/usr/bin/env python3
"""Extract Word citation hyperlinks and resolve them to BibTeX keys."""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


@dataclass
class BibEntry:
    key: str
    author: str = ""
    year: str = ""
    title: str = ""


def bib_key_to_bookmark(key: str) -> str:
    return "_Bib_" + re.sub(r"[^a-zA-Z0-9_]", "_", str(key))


def _field_value(body: str, name: str) -> str:
    m = re.search(r"\b" + re.escape(name) + r"\s*=\s*([{\"])", body, re.I)
    if not m:
        return ""
    opener = m.group(1)
    start = m.end()
    if opener == '"':
        end = body.find('"', start)
        return body[start:end].strip() if end >= 0 else ""
    depth = 1
    i = start
    while i < len(body) and depth:
        if body[i] == "{":
            depth += 1
        elif body[i] == "}":
            depth -= 1
        i += 1
    return body[start:i - 1].strip() if depth == 0 else ""


def parse_bib_entries(bib_path: str | Path | None) -> dict[str, BibEntry]:
    if not bib_path:
        return {}
    path = Path(bib_path)
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8", errors="ignore")
    entries = {}
    for m in re.finditer(r"@\w+\s*\{\s*([^,\s]+)\s*,", text):
        key = m.group(1).strip()
        start = m.end()
        depth = 1
        i = start
        while i < len(text) and depth:
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
            i += 1
        body = text[start:i - 1]
        entries[key] = BibEntry(
            key=key,
            author=_field_value(body, "author"),
            year=_field_value(body, "year"),
            title=_field_value(body, "title"),
        )
    return entries


def _first_author_last_name(author: str) -> str:
    first = re.split(r"\s+and\s+", author, flags=re.I)[0].strip()
    if "," in first:
        return first.split(",", 1)[0].strip()
    parts = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]+", first)
    return parts[-1] if parts else ""


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def make_citation_resolver(bib_path: str | Path | None = None) -> dict:
    entries = parse_bib_entries(bib_path)
    bookmarks = {bib_key_to_bookmark(key): key for key in entries}
    bookmarks.update({key: key for key in entries})
    author_year = []
    for entry in entries.values():
        last = _first_author_last_name(entry.author)
        if last and entry.year:
            author_year.append((_norm(last), entry.year, entry.key))
    return {"entries": entries, "bookmarks": bookmarks, "author_year": author_year}


def resolve_citation_key(target: str, display: str, resolver: dict | None) -> str:
    resolver = resolver or {}
    target = (target or "").strip()
    display = display or ""
    bookmarks = resolver.get("bookmarks", {})
    if target in bookmarks:
        return bookmarks[target]
    if target.startswith("_Bib_"):
        candidate = target[len("_Bib_"):]
        if candidate in bookmarks:
            return bookmarks[candidate]
    display_norm = _norm(display)
    for last, year, key in resolver.get("author_year", []):
        if year in display and last and last in display_norm:
            return key
    return ""


def _para_text(para) -> str:
    return "".join(t.text or "" for t in para.findall(f".//{W}t"))


def _field_records(root, resolver: dict | None) -> list[dict]:
    records = []
    for pi, para in enumerate(root.findall(f".//{W}p")):
        text = _para_text(para)
        active = None
        separated = False
        for run in para.findall(f"./{W}r"):
            fld = run.find(f"{W}fldChar")
            instr = run.find(f"{W}instrText")
            rtext = "".join(t.text or "" for t in run.findall(f".//{W}t"))
            if fld is not None:
                typ = fld.attrib.get(f"{W}fldCharType")
                if typ == "begin":
                    active = {"instruction": "", "display": ""}
                    separated = False
                elif typ == "separate":
                    separated = True
                elif typ == "end" and active is not None:
                    target = _internal_target(active["instruction"])
                    key = resolve_citation_key(target, active["display"], resolver)
                    if key:
                        records.append(_record(pi, text, target, active["display"], key, "field"))
                    active = None
                    separated = False
            elif instr is not None and active is not None:
                active["instruction"] += instr.text or ""
            elif separated and active is not None:
                active["display"] += rtext
    return records


def _hyperlink_records(root, resolver: dict | None) -> list[dict]:
    records = []
    for pi, para in enumerate(root.findall(f".//{W}p")):
        text = _para_text(para)
        for link in para.findall(f"./{W}hyperlink"):
            target = link.attrib.get(f"{W}anchor", "")
            display = _para_text(link)
            key = resolve_citation_key(target, display, resolver)
            if key:
                records.append(_record(pi, text, target, display, key, "hyperlink"))
    return records


def _internal_target(instruction: str) -> str:
    m = re.search(r'HYPERLINK\s+\\l\s+"([^"]+)"', instruction or "", re.I)
    return m.group(1) if m else ""


def _record(para_index: int, para_text: str, target: str, display: str, key: str, source: str) -> dict:
    return {
        "type": "word_link",
        "source": source,
        "text": display,
        "inner": display.strip("()"),
        "key": key,
        "numbers": [int(key)] if str(key).isdigit() else [],
        "target": target,
        "para_index": para_index,
        "section": "",
        "before": para_text[:80],
        "after": para_text[-80:],
    }


def extract_word_link_citations(docx_path: str | Path, bib_path: str | Path | None = None) -> dict:
    resolver = make_citation_resolver(bib_path)
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    citations = _field_records(root, resolver) + _hyperlink_records(root, resolver)
    seen = set()
    unique = []
    for item in citations:
        ident = (item["para_index"], item["target"], item["text"], item["key"])
        if ident not in seen:
            seen.add(ident)
            unique.append(item)
    return {
        "source_file": Path(docx_path).name,
        "total_link_citations": len(unique),
        "citations": unique,
    }
