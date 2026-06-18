"""Template-driven multipage LaTeX table generation."""

import base64

from shared.latex_text_utils import escape_latex


def requires_multipage_table(tbl_data, layout_spec):
    """Return true when the source table is taller than the template text area."""
    textheight_mm = (layout_spec or {}).get("page_geometry", {}).get("textheight_mm")
    environment = (layout_spec or {}).get("table", {}).get("multipage_environment")
    if environment != "supertabular" or not textheight_mm:
        return False
    height_mm = estimate_rendered_table_height_mm(tbl_data, layout_spec)
    return height_mm > float(textheight_mm) * 0.92


def estimate_rendered_table_height_mm(tbl_data, layout_spec=None):
    """Estimate the final TikZ table height after text wrapping."""
    rows = tbl_data.get("rows", [])
    grid_cols = tbl_data.get("grid_cols", [])
    if not rows:
        return 0.0
    col_widths_cm = _scaled_column_widths_cm(grid_cols, layout_spec)
    data_start = _data_start(rows, grid_cols)
    heights_cm = _row_heights_cm(rows[data_start:])
    for row_index, row in enumerate(rows):
        col_pos = 0
        for cell in row.get("cells", []):
            span = int(cell.get("gridSpan", 1) or 1)
            text = (cell.get("text") or "").strip()
            if text and col_widths_cm:
                width = sum(col_widths_cm[col_pos:col_pos + span])
                char_units = sum(1.0 if ord(ch) > 127 else 0.55 for ch in text)
                est_lines = max(1, int(char_units * 0.42 / width) + 1) if width > 0 else 1
                data_index = row_index - data_start
                if est_lines > 1 and 0 <= data_index < len(heights_cm):
                    heights_cm[data_index] = max(heights_cm[data_index], est_lines * 0.55)
            col_pos += span
    return sum(heights_cm) * 10.0


def nonpageable_table_height_limit_mm(layout_spec):
    """Return the template-relative body-height limit for a one-page table."""
    page_spec = (layout_spec or {}).get("page_geometry", {})
    try:
        textheight_mm = float(page_spec.get("textheight_mm") or 0)
    except (TypeError, ValueError):
        return 0.0
    return textheight_mm * 0.82 if textheight_mm else 0.0


def rendered_table_output_height_mm(tbl_data, layout_spec=None):
    """Return the table body height after the renderer's one-page scaling."""
    estimated = estimate_rendered_table_height_mm(tbl_data, layout_spec)
    limit = nonpageable_table_height_limit_mm(layout_spec)
    return min(estimated, limit) if limit else estimated


def _data_start(rows, grid_cols):
    if not rows or not grid_cols:
        return 0
    first_cells = rows[0].get("cells", [])
    if first_cells and int(first_cells[0].get("gridSpan", 1) or 1) == len(grid_cols):
        return 1
    return 0


def _row_heights_cm(rows):
    heights = []
    for row in rows:
        raw = row.get("row_height", 400)
        try:
            height = int(raw)
        except (TypeError, ValueError):
            height = 400
        heights.append(max(height, 400) / 567.0)
    return heights


def _scaled_column_widths_cm(grid_cols, layout_spec):
    widths = [float(col.get("width_twips", 0) or 0) / 567.0 for col in grid_cols]
    total = sum(widths)
    max_width = _max_table_width_cm(layout_spec)
    if total > 0 and max_width:
        widths = [max(width * max_width / total, 0.2) for width in widths]
    return widths


def _max_table_width_cm(layout_spec):
    page_spec = (layout_spec or {}).get("page_geometry", {})
    table_spec = (layout_spec or {}).get("table", {})
    width_mm = table_spec.get("max_width_mm") or table_spec.get("textwidth_mm") or page_spec.get("textwidth_mm")
    try:
        return float(width_mm) / 10.0 if width_mm else None
    except (TypeError, ValueError):
        return None


def _cell_is_bold(cell):
    if cell.get("bold"):
        return True
    return any(
        run.get("format", {}).get("bold")
        for para in cell.get("paragraphs", [])
        for run in para.get("runs", [])
    )


def _header_count(rows):
    last_bold = -1
    for index, row in enumerate(rows):
        if any(_cell_is_bold(cell) for cell in row.get("cells", [])):
            last_bold = index
        elif last_bold >= 0:
            break
    return last_bold + 1


def _row_latex(row, num_cols):
    cells = []
    cursor = 0
    for cell in row.get("cells", []):
        start = int(cell.get("col_start", cursor) or cursor)
        while cursor < start:
            cells.append("")
            cursor += 1
        span = int(cell.get("gridSpan", 1) or 1)
        text = "" if cell.get("vMerge") == "continue" else escape_latex(cell.get("text", ""))
        if text and _cell_is_bold(cell):
            text = rf"\textbf{{{text}}}"
        cells.append(rf"\multicolumn{{{span}}}{{c}}{{{text}}}" if span > 1 else text)
        cursor += span
    cells.extend([""] * max(0, num_cols - cursor))
    return " & ".join(cells) + r" \\"


def build_supertabular(tbl_data, tikz_code, caption, label_name, source_number, layout_spec):
    """Build the template's multipage table environment and preserve Word metadata."""
    grid_cols = tbl_data.get("grid_cols", [])
    rows = tbl_data.get("rows", [])
    total_width = sum(float(col.get("width_twips", 0) or 0) for col in grid_cols) or 1
    col_specs = []
    for index, col in enumerate(grid_cols):
        alignment = r"\raggedright" if index == 0 else r"\centering"
        fraction = float(col.get("width_twips", 0) or 0) / total_width
        col_specs.append(
            rf">{{{alignment}\arraybackslash}}p{{{fraction:.6f}\columnwidth}}"
        )
    header_count = _header_count(rows)
    header_rows = rows[:header_count]
    body_rows = rows[header_count:]
    header = "\n".join(_row_latex(row, len(grid_cols)) for row in header_rows)
    body = "\n".join(_row_latex(row, len(grid_cols)) for row in body_rows)
    metadata = base64.b64encode(tikz_code.encode("utf-8")).decode("ascii")
    number_cmd = rf"\renewcommand{{\thetable}}{{{source_number}}}" if source_number else ""
    label = rf"\label{{{label_name}}}" if label_name else ""

    return "\n".join([
        f"% WORD_SUPERTABLE_BEGIN number={source_number}",
        f"% WORD_SUPERTABLE_TIKZ={metadata}",
        "{",
        number_cmd,
        r"\sloppy",
        rf"\tablecaption{{{caption}}}{label}",
        rf"\tablehead{{\tophline {header} \middlehline}}",
        r"\tabletail{\middlehline}",
        rf"\begin{{supertabular}}{{@{{}}{'@{}'.join(col_specs)}@{{}}}}",
        body,
        r"\bottomhline",
        r"\end{supertabular}",
        "}",
        "% WORD_SUPERTABLE_END",
    ])
