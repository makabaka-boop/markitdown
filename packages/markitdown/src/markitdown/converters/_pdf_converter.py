import sys
import io
import re
from typing import BinaryIO, Any

from .._base_converter import DocumentConverter, DocumentConverterResult
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Pattern for MasterFormat-style partial numbering (e.g., ".1", ".2", ".10")
PARTIAL_NUMBERING_PATTERN = re.compile(r"^\.\d+$")


def _merge_partial_numbering_lines(text: str) -> str:
    """
    Post-process extracted text to merge MasterFormat-style partial numbering
    with the following text line.

    MasterFormat documents use partial numbering like:
        .1  The intent of this Request for Proposal...
        .2  Available information relative to...

    Some PDF extractors split these into separate lines:
        .1
        The intent of this Request for Proposal...

    This function merges them back together.
    """
    lines = text.split("\n")
    result_lines: list[str] = []
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Check if this line is ONLY a partial numbering
        if PARTIAL_NUMBERING_PATTERN.match(stripped):
            # Look for the next non-empty line to merge with
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1

            if j < len(lines):
                # Merge the partial numbering with the next line
                next_line = lines[j].strip()
                result_lines.append(f"{stripped} {next_line}")
                i = j + 1  # Skip past the merged line
            else:
                # No next line to merge with, keep as is
                result_lines.append(line)
                i += 1
        else:
            result_lines.append(line)
            i += 1

    return "\n".join(result_lines)


# Load dependencies
_dependency_exc_info = None
try:
    import pdfminer
    import pdfminer.high_level
    import pdfminer.pdftypes
    import pdfplumber
    import pdfplumber.utils
except ImportError:
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/pdf",
    "application/x-pdf",
]

ACCEPTED_FILE_EXTENSIONS = [".pdf"]


def _to_markdown_table(table: list[list[str]], include_separator: bool = True) -> str:
    """Convert a 2D list (rows/columns) into a nicely aligned Markdown table.

    Args:
        table: 2D list of cell values
        include_separator: If True, include header separator row (standard markdown).
                          If False, output simple pipe-separated rows.
    """
    if not table:
        return ""

    # Normalize None → ""
    table = [[cell if cell is not None else "" for cell in row] for row in table]

    # Filter out empty rows
    table = [row for row in table if any(cell.strip() for cell in row)]

    if not table:
        return ""

    # Column widths
    col_widths = [max(len(str(cell)) for cell in col) for col in zip(*table)]

    def fmt_row(row: list[str]) -> str:
        return (
            "|"
            + "|".join(str(cell).ljust(width) for cell, width in zip(row, col_widths))
            + "|"
        )

    if include_separator:
        header, *rows = table
        md = [fmt_row(header)]
        md.append("|" + "|".join("-" * w for w in col_widths) + "|")
        for row in rows:
            md.append(fmt_row(row))
    else:
        md = [fmt_row(row) for row in table]

    return "\n".join(md)


def _extract_form_content_from_words(page: Any) -> str | None:
    """
    Extract form-style content from a PDF page by analyzing word positions.
    This handles borderless forms/tables where words are aligned in columns.

    Returns markdown with proper table formatting:
    - Tables have pipe-separated columns with header separator rows
    - Non-table content is rendered as plain text

    Returns None if the page doesn't appear to be a form-style document,
    indicating that pdfminer should be used instead for better text spacing.
    """
    words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
    if not words:
        return None

    # Group words by their Y position (rows)
    y_tolerance = 5
    rows_by_y: dict[float, list[dict]] = {}
    for word in words:
        y_key = round(word["top"] / y_tolerance) * y_tolerance
        if y_key not in rows_by_y:
            rows_by_y[y_key] = []
        rows_by_y[y_key].append(word)

    # Sort rows by Y position
    sorted_y_keys = sorted(rows_by_y.keys())
    page_width = page.width if hasattr(page, "width") else 612

    # First pass: analyze each row
    row_info: list[dict] = []
    for y_key in sorted_y_keys:
        row_words = sorted(rows_by_y[y_key], key=lambda w: w["x0"])
        if not row_words:
            continue

        first_x0 = row_words[0]["x0"]
        last_x1 = row_words[-1]["x1"]
        line_width = last_x1 - first_x0
        combined_text = " ".join(w["text"] for w in row_words)

        # Count distinct x-position groups (columns)
        x_positions = [w["x0"] for w in row_words]
        x_groups: list[float] = []
        for x in sorted(x_positions):
            if not x_groups or x - x_groups[-1] > 50:
                x_groups.append(x)

        # Determine row type
        is_paragraph = line_width > page_width * 0.55 and len(combined_text) > 60

        # Check for MasterFormat-style partial numbering (e.g., ".1", ".2")
        # These should be treated as list items, not table rows
        has_partial_numbering = False
        if row_words:
            first_word = row_words[0]["text"].strip()
            if PARTIAL_NUMBERING_PATTERN.match(first_word):
                has_partial_numbering = True

        row_info.append(
            {
                "y_key": y_key,
                "words": row_words,
                "text": combined_text,
                "x_groups": x_groups,
                "is_paragraph": is_paragraph,
                "num_columns": len(x_groups),
                "has_partial_numbering": has_partial_numbering,
            }
        )

    # Collect ALL x-positions from rows with 3+ columns (table-like rows)
    # This gives us the global column structure
    all_table_x_positions: list[float] = []
    for info in row_info:
        if info["num_columns"] >= 3 and not info["is_paragraph"]:
            all_table_x_positions.extend(info["x_groups"])

    if not all_table_x_positions:
        return None

    # Compute adaptive column clustering tolerance based on gap analysis
    all_table_x_positions.sort()

    # Calculate gaps between consecutive x-positions
    gaps = []
    for i in range(len(all_table_x_positions) - 1):
        gap = all_table_x_positions[i + 1] - all_table_x_positions[i]
        if gap > 5:  # Only significant gaps
            gaps.append(gap)

    # Determine optimal tolerance using statistical analysis
    if gaps and len(gaps) >= 3:
        # Use 70th percentile of gaps as threshold (balances precision/recall)
        sorted_gaps = sorted(gaps)
        percentile_70_idx = int(len(sorted_gaps) * 0.70)
        adaptive_tolerance = sorted_gaps[percentile_70_idx]

        # Clamp tolerance to reasonable range [25, 50]
        adaptive_tolerance = max(25, min(50, adaptive_tolerance))
    else:
        # Fallback to conservative value
        adaptive_tolerance = 35

    # Compute global column boundaries using adaptive tolerance
    global_columns: list[float] = []
    for x in all_table_x_positions:
        if not global_columns or x - global_columns[-1] > adaptive_tolerance:
            global_columns.append(x)

    # Adaptive max column check based on page characteristics
    # Calculate average column width
    if len(global_columns) > 1:
        content_width = global_columns[-1] - global_columns[0]
        avg_col_width = content_width / len(global_columns)

        # Forms with very narrow columns (< 30px) are likely dense text
        if avg_col_width < 30:
            return None

        # Compute adaptive max based on columns per inch
        # Typical forms have 3-8 columns per inch
        columns_per_inch = len(global_columns) / (content_width / 72)

        # If density is too high (> 10 cols/inch), likely not a form
        if columns_per_inch > 10:
            return None

        # Adaptive max: allow more columns for wider pages
        # Standard letter is 612pt wide, so scale accordingly
        adaptive_max_columns = int(20 * (page_width / 612))
        adaptive_max_columns = max(15, adaptive_max_columns)  # At least 15

        if len(global_columns) > adaptive_max_columns:
            return None
    else:
        # Single column, not a form
        return None

    # Now classify each row as table row or not
    # A row is a table row if it has words that align with 2+ of the global columns
    for info in row_info:
        if info["is_paragraph"]:
            info["is_table_row"] = False
            continue

        # Rows with partial numbering (e.g., ".1", ".2") are list items, not table rows
        if info["has_partial_numbering"]:
            info["is_table_row"] = False
            continue

        # Count how many global columns this row's words align with
        aligned_columns: set[int] = set()
        for word in info["words"]:
            word_x = word["x0"]
            for col_idx, col_x in enumerate(global_columns):
                if abs(word_x - col_x) < 40:
                    aligned_columns.add(col_idx)
                    break

        # If row uses 2+ of the established columns, it's a table row
        info["is_table_row"] = len(aligned_columns) >= 2

    # Find table regions (consecutive table rows)
    table_regions: list[tuple[int, int]] = []  # (start_idx, end_idx)
    i = 0
    while i < len(row_info):
        if row_info[i]["is_table_row"]:
            start_idx = i
            while i < len(row_info) and row_info[i]["is_table_row"]:
                i += 1
            end_idx = i
            table_regions.append((start_idx, end_idx))
        else:
            i += 1

    # Check if enough rows are table rows (at least 20%)
    total_table_rows = sum(end - start for start, end in table_regions)
    if len(row_info) > 0 and total_table_rows / len(row_info) < 0.2:
        return None

    # Build output - collect table data first, then format with proper column widths
    result_lines: list[str] = []
    num_cols = len(global_columns)

    # Helper function to extract cells from a row
    def extract_cells(info: dict) -> list[str]:
        cells: list[str] = ["" for _ in range(num_cols)]
        for word in info["words"]:
            word_x = word["x0"]
            # Find the correct column using boundary ranges
            assigned_col = num_cols - 1  # Default to last column
            for col_idx in range(num_cols - 1):
                col_end = global_columns[col_idx + 1]
                if word_x < col_end - 20:
                    assigned_col = col_idx
                    break
            if cells[assigned_col]:
                cells[assigned_col] += " " + word["text"]
            else:
                cells[assigned_col] = word["text"]
        return cells

    # Process rows, collecting table data for proper formatting
    idx = 0
    while idx < len(row_info):
        info = row_info[idx]

        # Check if this row starts a table region
        table_region = None
        for start, end in table_regions:
            if idx == start:
                table_region = (start, end)
                break

        if table_region:
            start, end = table_region
            # Collect all rows in this table
            table_data: list[list[str]] = []
            for table_idx in range(start, end):
                cells = extract_cells(row_info[table_idx])
                table_data.append(cells)

            # Calculate column widths for this table
            if table_data:
                col_widths = [
                    max(len(row[col]) for row in table_data) for col in range(num_cols)
                ]
                # Ensure minimum width of 3 for separator dashes
                col_widths = [max(w, 3) for w in col_widths]

                # Format header row
                header = table_data[0]
                header_str = (
                    "| "
                    + " | ".join(
                        cell.ljust(col_widths[i]) for i, cell in enumerate(header)
                    )
                    + " |"
                )
                result_lines.append(header_str)

                # Format separator row
                separator = (
                    "| "
                    + " | ".join("-" * col_widths[i] for i in range(num_cols))
                    + " |"
                )
                result_lines.append(separator)

                # Format data rows
                for row in table_data[1:]:
                    row_str = (
                        "| "
                        + " | ".join(
                            cell.ljust(col_widths[i]) for i, cell in enumerate(row)
                        )
                        + " |"
                    )
                    result_lines.append(row_str)

            idx = end  # Skip to end of table region
        else:
            # Check if we're inside a table region (not at start)
            in_table = False
            for start, end in table_regions:
                if start < idx < end:
                    in_table = True
                    break

            if not in_table:
                # Non-table content
                result_lines.append(info["text"])
            idx += 1

    return "\n".join(result_lines)


def _extract_tables_from_words(page: Any) -> list[list[list[str]]]:
    """
    Extract tables from a PDF page by analyzing word positions.
    This handles borderless tables where words are aligned in columns.

    This function is designed for structured tabular data (like invoices),
    not for multi-column text layouts in scientific documents.
    """
    words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
    if not words:
        return []

    # Group words by their Y position (rows)
    y_tolerance = 5
    rows_by_y: dict[float, list[dict]] = {}
    for word in words:
        y_key = round(word["top"] / y_tolerance) * y_tolerance
        if y_key not in rows_by_y:
            rows_by_y[y_key] = []
        rows_by_y[y_key].append(word)

    # Sort rows by Y position
    sorted_y_keys = sorted(rows_by_y.keys())

    # Find potential column boundaries by analyzing x positions across all rows
    all_x_positions = []
    for words_in_row in rows_by_y.values():
        for word in words_in_row:
            all_x_positions.append(word["x0"])

    if not all_x_positions:
        return []

    # Cluster x positions to find column starts
    all_x_positions.sort()
    x_tolerance_col = 20
    column_starts: list[float] = []
    for x in all_x_positions:
        if not column_starts or x - column_starts[-1] > x_tolerance_col:
            column_starts.append(x)

    # Need at least 3 columns but not too many (likely text layout, not table)
    if len(column_starts) < 3 or len(column_starts) > 10:
        return []

    # Find rows that span multiple columns (potential table rows)
    table_rows = []
    for y_key in sorted_y_keys:
        words_in_row = sorted(rows_by_y[y_key], key=lambda w: w["x0"])

        # Assign words to columns
        row_data = [""] * len(column_starts)
        for word in words_in_row:
            # Find the closest column
            best_col = 0
            min_dist = float("inf")
            for i, col_x in enumerate(column_starts):
                dist = abs(word["x0"] - col_x)
                if dist < min_dist:
                    min_dist = dist
                    best_col = i

            if row_data[best_col]:
                row_data[best_col] += " " + word["text"]
            else:
                row_data[best_col] = word["text"]

        # Only include rows that have content in multiple columns
        non_empty = sum(1 for cell in row_data if cell.strip())
        if non_empty >= 2:
            table_rows.append(row_data)

    # Validate table quality - tables should have:
    # 1. Enough rows (at least 3 including header)
    # 2. Short cell content (tables have concise data, not paragraphs)
    # 3. Consistent structure across rows
    if len(table_rows) < 3:
        return []

    # Check if cells contain short, structured data (not long text)
    long_cell_count = 0
    total_cell_count = 0
    for row in table_rows:
        for cell in row:
            if cell.strip():
                total_cell_count += 1
                # If cell has more than 30 chars, it's likely prose text
                if len(cell.strip()) > 30:
                    long_cell_count += 1

    # If more than 30% of cells are long, this is probably not a table
    if total_cell_count > 0 and long_cell_count / total_cell_count > 0.3:
        return []

    return [table_rows]


# ---------------------------------------------------------------------------
# Annotation extraction (opt-in via pdf_extract_annotations=True)
# ---------------------------------------------------------------------------

ANNOTATIONS_HEADING = "## Annotations"

# Markup subtypes that carry review information. Popups and links are ignored.
_ANNOT_SUBTYPE_NOTE = "Text"
_ANNOT_SUBTYPE_HIGHLIGHT = "Highlight"
_ANNOT_TYPE_LABELS = {
    _ANNOT_SUBTYPE_NOTE: "Note",
    _ANNOT_SUBTYPE_HIGHLIGHT: "Highlight",
}

# Padding (in PDF points) around a highlight quad when matching characters.
# Zero padding clips glyph edges; more than ~1pt starts pulling in text from
# neighbouring lines, which must never leak into a quoted selection.
_ANNOT_QUAD_PADDING = 1.0
# Two highlight quads whose vertical centres are within this distance are
# considered fragments of the same visual line (joined with a space);
# otherwise they represent different lines (joined with a newline).
_ANNOT_SAME_LINE_TOLERANCE = 4.0


def _decode_pdf_text_string(value: Any) -> str:
    """Decode a PDF text string (``Contents``/``T``) and normalize newlines."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        if value.startswith((b"\xfe\xff", b"\xff\xfe")):
            # UTF-16 with a byte-order mark, as written by most PDF producers.
            value = value.decode("utf-16", errors="replace")
        else:
            try:
                value = value.decode("utf-8")
            except UnicodeDecodeError:
                # PDFDocEncoding is close enough to Latin-1 for retrieval.
                value = value.decode("latin-1")
    text = str(value)
    # PDF line endings may be CR, LF, or CRLF.
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _pdf_name(value: Any) -> str:
    """Return the plain name for a pdfminer name literal (``Subtype`` etc.)."""
    name = getattr(value, "name", value)
    return str(name) if name is not None else ""


def _clean_multiline_text(text: str) -> str:
    """Trim surrounding whitespace per line while preserving line breaks."""
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def _rotate_pdf_point(
    x: float, y: float, rotation: int, width: float, height: float
) -> tuple[float, float]:
    """Rotate a point from raw PDF space the same way pdfplumber rotates Rects."""
    turns = rotation // 90
    for i in range(turns):
        comp = width if i == turns % 2 else height
        x, y = y, comp - x
    return x, y


def _pdf_rect_to_bbox(points: list[tuple[float, float]], page: Any) -> tuple | None:
    """Convert raw PDF-space polygon points to a top-based pdfplumber bbox."""
    try:
        rotated = [
            _rotate_pdf_point(x, y, page.rotation, page.width, page.height)
            for x, y in points
        ]
    except Exception:
        return None
    if not rotated:
        return None
    x0 = min(x for x, _ in rotated)
    x1 = max(x for x, _ in rotated)
    y0 = min(y for _, y in rotated)
    y1 = max(y for _, y in rotated)
    # Rotated points are still bottom-up; invert against the page height.
    return (x0, page.height - y1, x1, page.height - y0)


def _highlight_bboxes(raw_annot: dict, page: Any) -> list[tuple]:
    """Return the top-based bbox of each quad in a highlight annotation.

    Falls back to the annotation rectangle when ``QuadPoints`` is missing or
    malformed.
    """
    quadpoints = raw_annot.get("QuadPoints")
    bboxes: list[tuple] = []
    if isinstance(quadpoints, list):
        for i in range(0, len(quadpoints) - 7, 8):
            group = quadpoints[i : i + 8]
            points = [(group[j], group[j + 1]) for j in (0, 2, 4, 6)]
            if not all(isinstance(coord, (int, float)) for coord in group):
                continue
            bbox = _pdf_rect_to_bbox(points, page)
            if bbox is not None:
                bboxes.append(bbox)

    if not bboxes:
        rect = raw_annot.get("Rect")
        if isinstance(rect, list) and len(rect) == 4:
            points = [
                (rect[0], rect[1]),
                (rect[2], rect[1]),
                (rect[0], rect[3]),
                (rect[2], rect[3]),
            ]
            if all(isinstance(coord, (int, float)) for coord in rect):
                bbox = _pdf_rect_to_bbox(points, page)
                if bbox is not None:
                    bboxes.append(bbox)
    return bboxes


def _extract_highlight_quote(raw_annot: dict, page: Any) -> str:
    """Quote the text actually covered by a highlight annotation.

    Each quad is extracted independently so that text lying *between*
    discontinuous selections is never included. Characters shared by
    overlapping quads are only emitted once. Quads on different lines keep
    their line break; fragments on the same line are joined with a space.
    """
    bboxes = _highlight_bboxes(raw_annot, page)
    if not bboxes:
        return ""

    try:
        chars = page.chars
    except Exception:
        return ""

    fragments: list[tuple[float, float, str]] = []
    seen_chars: set[int] = set()
    for x0, top, x1, bottom in bboxes:
        selected = []
        for char in chars:
            char_id = id(char)
            if char_id in seen_chars:
                continue
            try:
                inside = (
                    char["x0"] >= x0 - _ANNOT_QUAD_PADDING
                    and char["x1"] <= x1 + _ANNOT_QUAD_PADDING
                    and char["top"] >= top - _ANNOT_QUAD_PADDING
                    and char["bottom"] <= bottom + _ANNOT_QUAD_PADDING
                )
            except (KeyError, TypeError):
                continue
            if inside:
                selected.append(char)
                seen_chars.add(char_id)
        if not selected:
            continue
        try:
            text = pdfplumber.utils.extract_text(
                selected, x_tolerance=1.5, y_tolerance=3
            )
        except Exception:
            continue
        text = text.strip()
        if text:
            mean_top = sum(char["top"] for char in selected) / len(selected)
            mean_x = sum(char["x0"] for char in selected) / len(selected)
            fragments.append((mean_top, mean_x, text))

    if not fragments:
        return ""

    fragments.sort(key=lambda fragment: (fragment[0], fragment[1]))
    lines: list[str] = []
    previous_top: float | None = None
    for mean_top, _mean_x, text in fragments:
        if (
            previous_top is not None
            and abs(mean_top - previous_top) <= _ANNOT_SAME_LINE_TOLERANCE
        ):
            lines[-1] = f"{lines[-1]} {text}"
        else:
            lines.append(text)
        previous_top = mean_top
    return "\n".join(lines)


def _iter_raw_annotations(page_obj: Any) -> list[tuple[int, Any, dict]]:
    """Return ``(array_index, objid_or_None, raw_dict)`` for a page.

    Entries that cannot be resolved are dropped here; a single bad reference
    must not prevent the remaining annotations from being read.
    """
    try:
        raw_annots = pdfminer.pdftypes.resolve1(page_obj.annots)
    except Exception:
        return []
    if not isinstance(raw_annots, list):
        return []

    annotations: list[tuple[int, Any, dict]] = []
    for order, reference in enumerate(raw_annots):
        try:
            if isinstance(reference, pdfminer.pdftypes.PDFObjRef):
                objid: Any = reference.objid
                raw = pdfminer.pdftypes.resolve1(reference)
            else:
                objid = None
                raw = reference
        except Exception:
            continue
        if isinstance(raw, dict):
            annotations.append((order, objid, raw))
    return annotations


def _build_annotation_record(
    page: Any, page_number: int, order: int, objid: Any, raw: dict
) -> dict | None:
    """Build one annotation record, or ``None`` if it carries no content."""
    try:
        subtype = _pdf_name(raw.get("Subtype"))
        if subtype not in _ANNOT_TYPE_LABELS:
            return None

        author = _clean_multiline_text(_decode_pdf_text_string(raw.get("T")))
        content = _clean_multiline_text(_decode_pdf_text_string(raw.get("Contents")))

        quote = ""
        if subtype == _ANNOT_SUBTYPE_HIGHLIGHT:
            quote = _extract_highlight_quote(raw, page)

        # Notes need comment text; highlights are kept whenever they quote
        # selected text, even with no comment of their own.
        if subtype == _ANNOT_SUBTYPE_NOTE and not content:
            return None
        if subtype == _ANNOT_SUBTYPE_HIGHLIGHT and not content and not quote:
            return None

        parent: Any = None
        irt = raw.get("IRT")
        if isinstance(irt, pdfminer.pdftypes.PDFObjRef):
            parent = irt.objid
        elif isinstance(irt, dict):
            parent = f"inline-{id(irt)}"

        if objid is None:
            key: Any = f"inline-{id(raw)}"
        else:
            key = objid

        return {
            "key": key,
            "raw_key": id(raw),
            "page": page_number,
            "order": order,
            "subtype": subtype,
            "author": author,
            "content": content,
            "quote": quote,
            "parent": parent,
        }
    except Exception:
        # Never let a single malformed annotation break the others.
        return None


def _render_annotation_thread(
    record: dict,
    children_by_parent: dict,
    emitted: set,
    depth: int,
) -> list[str]:
    """Render an annotation and its reply subtree, each entry exactly once."""
    if record["key"] in emitted:
        return []
    emitted.add(record["key"])

    indent = "  " * depth
    body_indent = indent + "  "
    label = "Reply" if depth > 0 else _ANNOT_TYPE_LABELS[record["subtype"]]
    if record["author"]:
        label = f"{label} — {record['author']}"

    lines = [f"{indent}- **{label}**"]
    if record["quote"]:
        for quote_line in record["quote"].split("\n"):
            lines.append(f"{body_indent}> {quote_line}")
    if record["content"]:
        for content_line in record["content"].split("\n"):
            lines.append(f"{body_indent}{content_line}")

    for child in children_by_parent.get(record["key"], []):
        lines.append("")
        lines.extend(
            _render_annotation_thread(child, children_by_parent, emitted, depth + 1)
        )
    return lines


def _extract_pdf_annotations(pdf_bytes: BinaryIO) -> str | None:
    """Build the annotations section for a PDF, or ``None`` if there is none.

    Independent annotations are ordered by page and by their position in the
    page's annotation array. Replies are nested under their parent (via the
    ``/IRT`` reference) and emitted once; replies whose parent is missing are
    rendered as independent annotations.
    """
    records: list[dict] = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            page_number = page_idx + 1
            try:
                for order, objid, raw in _iter_raw_annotations(page.page_obj):
                    record = _build_annotation_record(
                        page, page_number, order, objid, raw
                    )
                    if record is not None:
                        records.append(record)
            except Exception:
                # A broken page must not take down annotation extraction.
                continue
            finally:
                page.close()

    if not records:
        return None

    by_key = {record["key"]: record for record in records}
    # Also allow inline (non-indirect) parents to be matched by object id.
    for record in records:
        by_key[f"inline-{record['raw_key']}"] = record

    children_by_parent: dict = {}
    roots: list[dict] = []
    for record in records:
        parent = record["parent"]
        if parent is not None and parent in by_key and parent != record["key"]:
            children_by_parent.setdefault(parent, []).append(record)
        else:
            # Independent annotation, or a reply whose parent is missing.
            roots.append(record)

    emitted: set = set()
    lines = [ANNOTATIONS_HEADING, ""]
    for page_number in sorted({record["page"] for record in roots}):
        page_roots = [
            record
            for record in roots
            if record["page"] == page_number and record["key"] not in emitted
        ]
        if not page_roots:
            continue
        lines.append(f"### Page {page_number}")
        lines.append("")
        for record in page_roots:
            lines.extend(
                _render_annotation_thread(record, children_by_parent, emitted, 0)
            )
            lines.append("")

    # Defensive fallback for anything not reached (e.g. a reply cycle).
    leftovers = [record for record in records if record["key"] not in emitted]
    for record in sorted(leftovers, key=lambda r: (r["page"], r["order"])):
        lines.extend(_render_annotation_thread(record, children_by_parent, emitted, 0))
        lines.append("")

    section = "\n".join(lines).strip()
    return section or None


class PdfConverter(DocumentConverter):
    """
    Converts PDFs to Markdown.
    Supports extracting tables into aligned Markdown format (via pdfplumber).
    Falls back to pdfminer if pdfplumber is missing or fails.

    Pass ``pdf_extract_annotations=True`` to additionally extract review
    annotations (sticky notes, highlights, and threaded replies). They are
    appended in a dedicated section after the document body.
    """

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,
    ) -> DocumentConverterResult:
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".pdf",
                    feature="pdf",
                )
            ) from _dependency_exc_info[1].with_traceback(
                _dependency_exc_info[2]
            )  # type: ignore[union-attr]

        assert isinstance(file_stream, io.IOBase)

        # Read file stream into BytesIO for compatibility with pdfplumber
        pdf_bytes = io.BytesIO(file_stream.read())

        try:
            # Single pass: check every page for form-style content.
            # Pages with tables/forms get rich extraction; plain-text
            # pages are collected separately. page.close() is called
            # after each page to free pdfplumber's cached objects and
            # keep memory usage constant regardless of page count.
            markdown_chunks: list[str] = []
            form_page_count = 0
            plain_page_indices: list[int] = []

            with pdfplumber.open(pdf_bytes) as pdf:
                for page_idx, page in enumerate(pdf.pages):
                    page_content = _extract_form_content_from_words(page)

                    if page_content is not None:
                        form_page_count += 1
                        if page_content.strip():
                            markdown_chunks.append(page_content)
                    else:
                        plain_page_indices.append(page_idx)
                        text = page.extract_text()
                        if text and text.strip():
                            markdown_chunks.append(text.strip())

                    page.close()  # Free cached page data immediately

            # If no pages had form-style content, use pdfminer for
            # the whole document (better text spacing for prose).
            if form_page_count == 0:
                pdf_bytes.seek(0)
                markdown = pdfminer.high_level.extract_text(pdf_bytes)
            else:
                markdown = "\n\n".join(markdown_chunks).strip()

        except Exception:
            # Fallback if pdfplumber fails
            pdf_bytes.seek(0)
            markdown = pdfminer.high_level.extract_text(pdf_bytes)

        # Fallback if still empty
        if not markdown:
            pdf_bytes.seek(0)
            markdown = pdfminer.high_level.extract_text(pdf_bytes)

        # Post-process to merge MasterFormat-style partial numbering with following text
        markdown = _merge_partial_numbering_lines(markdown)

        # Optional: append review annotations (notes, highlights, replies).
        # Runs in a separate pass and after body post-processing so that the
        # default output is byte-for-byte unchanged and annotation extraction
        # can never affect body/table extraction or its fallbacks.
        if kwargs.get("pdf_extract_annotations"):
            annotation_section: str | None = None
            try:
                pdf_bytes.seek(0)
                annotation_section = _extract_pdf_annotations(pdf_bytes)
            except Exception:
                annotation_section = None
            if annotation_section:
                markdown = (
                    f"{markdown}\n\n{annotation_section}"
                    if markdown.strip()
                    else annotation_section
                )

        return DocumentConverterResult(markdown=markdown)
