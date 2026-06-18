"""Parse .bbl bibliography items for Word reference-section generation."""

from __future__ import annotations

import os
import re
from dataclasses import dataclass


@dataclass
class BibItem:
    """One bibliography item extracted from a .bbl file."""

    key: str
    raw_text: str = ""
    plain_text: str = ""
    doi_url: str | None = None
    external_url: str | None = None


_ACCENT_MAP = {
    r"\'\i": "\u00ed", r"\'{\i}": "\u00ed",
    r"\'a": "\u00e1", r"\'e": "\u00e9", r"\'i": "\u00ed",
    r"\'o": "\u00f3", r"\'u": "\u00fa", r"\'A": "\u00c1",
    r"\'E": "\u00c9", r"\'I": "\u00cd", r"\'O": "\u00d3",
    r"\'U": "\u00da", r"\'{a}": "\u00e1", r"\'{e}": "\u00e9",
    r"\'{i}": "\u00ed", r"\'{o}": "\u00f3", r"\'{u}": "\u00fa",
    r"\`a": "\u00e0", r"\`e": "\u00e8", r"\`i": "\u00ec",
    r"\`o": "\u00f2", r"\`u": "\u00f9", r"\`{a}": "\u00e0",
    r"\`{e}": "\u00e8", r"\`{i}": "\u00ec", r"\`{o}": "\u00f2",
    r"\`{u}": "\u00f9", r'\"a': "\u00e4", r'\"e': "\u00eb",
    r'\"i': "\u00ef", r'\"o': "\u00f6", r'\"u': "\u00fc",
    r'\"{a}': "\u00e4", r'\"{e}': "\u00eb", r'\"{i}': "\u00ef",
    r'\"{o}': "\u00f6", r'\"{u}': "\u00fc", r"\~n": "\u00f1",
    r"\~{n}": "\u00f1", r"\c{c}": "\u00e7", r"\c{C}": "\u00c7",
}


def parse_bbl_items(bbl_path: str) -> list[BibItem]:
    """Extract ordered bibliography items from a .bbl file."""
    if not bbl_path or not os.path.isfile(bbl_path):
        print(f"  [bbl_item_parser] .bbl not found: {bbl_path}")
        return []

    with open(bbl_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    pattern = re.compile(r"\\bibitem(?:\[[^\]]*\])?\{([^}]+)\}")
    matches = list(pattern.finditer(content))
    items: list[BibItem] = []

    for idx, match in enumerate(matches):
        key = match.group(1).strip()
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(content)
        raw = content[start:end]
        raw = re.sub(r"\\end\{thebibliography\}.*", "", raw, flags=re.DOTALL).strip()
        if not raw:
            continue

        doi_url = _extract_doi(raw)
        external_url = _extract_url(raw)
        plain_text = _to_plain_text(raw)
        items.append(BibItem(key, raw, plain_text, doi_url, external_url))

    return items


def _extract_doi(text: str) -> str | None:
    match = re.search(r"\\doi\{([^}]+)\}", text)
    if match:
        doi = match.group(1).strip()
        return doi if doi.startswith("http") else f"https://doi.org/{doi}"
    match = re.search(r"https?://(?:dx\.)?doi\.org/[^\s{}\\]+", text)
    return match.group(0) if match else None


def _extract_url(text: str) -> str | None:
    for match in re.finditer(r"\\url\{([^}]+)\}", text):
        url = match.group(1).strip()
        if "doi.org/" not in url:
            return url
    for match in re.finditer(r"https?://[^\s{}\\]+", text):
        url = match.group(0)
        if "doi.org/" not in url:
            return url
    return None


def _to_plain_text(text: str) -> str:
    result = text
    result = re.sub(r"\\begin\{thebibliography\}\{[^}]*\}", " ", result)
    result = re.sub(r"\\end\{thebibliography\}", " ", result)
    result = re.sub(r"\\href\{[^}]*\}\{([^}]*)\}", r"\1", result)
    result = re.sub(r"\\doi\{[^}]*\}", " ", result)
    result = re.sub(r"\\url\{[^}]*\}", " ", result)
    result = re.sub(r"\\newblock\s*", " ", result)
    result = re.sub(r"\\natexlab\{([^}]*)\}", r"\1", result)

    for pattern in sorted(_ACCENT_MAP, key=len, reverse=True):
        result = result.replace(pattern, _ACCENT_MAP[pattern])

    for command in ("emph", "textbf", "textit", "textrm", "text", "mathrm"):
        result = re.sub(rf"\\{command}\{{([^{{}}]*)\}}", r"\1", result)

    replacements = {
        "~": " ",
        "---": "\u2014",
        "--": "\u2013",
        "``": '"',
        "''": '"',
        r"\&": "&",
        r"\$": "$",
        r"\%": "%",
        r"\#": "#",
        r"\_": "_",
        r"\{": "{",
        r"\}": "}",
    }
    for old, new in replacements.items():
        result = result.replace(old, new)

    # bibinfo 字段标记（elsarticle/natbib bst 输出）：
    # \bibinfo{author}{Baker, D.} → Baker, D.（只保留内容，去掉字段名）
    # 必须在通用 \command 清理之前处理，否则字段名会与内容粘连
    result = re.sub(r"\\bibinfo\{[^}]*\}\{([^{}]*)\}", r"\1", result)

    result = re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?", " ", result)
    for _ in range(8):
        prev = result
        result = re.sub(r"\{([^{}]*)\}", r"\1", result)
        if result == prev:
            break
    return re.sub(r"\s+", " ", result).strip()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Parse .bbl items")
    parser.add_argument("bbl_file")
    args = parser.parse_args()
    for item in parse_bbl_items(args.bbl_file):
        print(f"{item.key}: {item.plain_text[:120]}")
