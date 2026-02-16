"""Convert .docx paragraphs to pseudo-Markdown fragment representation.

This module converts the runs within each paragraph to a pseudo-Markdown format
that preserves bold (**), italic (_), and underline (__) formatting while being
suitable for LLM consumption.

Unicode characters (smart quotes, em-dashes, section symbols, etc.) are preserved
as-is. Only Markdown syntax characters (*, _) are escaped with backslashes.

Tables are extracted as structured data (CellInfo/TableInfo) with cell text in
pseudo-Markdown. Non-simple tables are represented as SkippedTableInfo with a
reason.
"""

from __future__ import annotations

import re

from lxml import etree

from docx_mcp.models import CellInfo, SkippedTableInfo, TableInfo
from docx_mcp.namespaces import qn, xpath
from docx_mcp.table_utils import get_cell_paragraphs, is_simple_table, table_dimensions


def _has_bool_property(rpr: etree._Element | None, local_name: str) -> bool:
    """Check if a run property element has a boolean property set to true.

    In OOXML, a boolean property like <w:b/> means true.
    <w:b w:val="false"/> or <w:b w:val="0"/> means false.
    Absence means inherit/false.
    """
    if rpr is None:
        return False
    elements = xpath(rpr, f"w:{local_name}")
    if not elements:
        return False
    el = elements[0]
    val = el.get(qn("w", "val"))
    if val is None:
        return True  # <w:b/> with no val means true
    return val.lower() not in ("false", "0", "off")


def _is_bold(rpr: etree._Element | None) -> bool:
    return _has_bool_property(rpr, "b")


def _is_italic(rpr: etree._Element | None) -> bool:
    return _has_bool_property(rpr, "i")


def _is_underline(rpr: etree._Element | None) -> bool:
    """Check if run has single underline (the most common type)."""
    if rpr is None:
        return False
    elements = xpath(rpr, "w:u")
    if not elements:
        return False
    el = elements[0]
    val = el.get(qn("w", "val"))
    # "none" means explicitly no underline; absence of the element means no underline
    # "single" or any other value means underlined
    if val is None:
        return True
    return val.lower() not in ("none", "false", "0")


def _escape_markdown(text: str) -> str:
    """Escape Markdown syntax characters in text.

    We escape * and _ since they are our formatting markers.
    We do NOT escape other characters like #, |, [, ] for the MVP
    since they don't conflict with our pseudo-Markdown syntax.
    """
    # Escape backslashes first, then * and _
    text = text.replace("\\", "\\\\")
    text = text.replace("*", "\\*")
    text = text.replace("_", "\\_")
    return text


def _unescape_markdown(text: str) -> str:
    """Reverse the escaping done by _escape_markdown."""
    text = text.replace("\\_", "_")
    text = text.replace("\\*", "*")
    text = text.replace("\\\\", "\\")
    return text


# Regex that strips our pseudo-Markdown formatting wrappers.
# Handles bold (**…**), underline (__…__), and italic (_…_).
# Must be applied iteratively (outermost first) because wrappers nest.
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.DOTALL)
_UNDERLINE_RE = re.compile(r"__(.+?)__", re.DOTALL)
_ITALIC_RE = re.compile(r"(?<![\\])_(.+?)(?<![\\])_", re.DOTALL)


def pseudo_markdown_to_raw(text: str) -> str:
    """Convert pseudo-Markdown text back to raw (plain) text.

    Strips formatting wrappers (``**bold**``, ``__underline__``,
    ``_italic_``) and reverses escape sequences (``\\*`` → ``*``,
    ``\\_`` → ``_``, ``\\\\`` → ``\\``).

    This is the inverse of the formatting applied by
    :func:`paragraph_to_pseudo_markdown`, modulo whitespace normalisation.

    Args:
        text: Pseudo-Markdown text.

    Returns:
        Plain text with formatting markers and escapes removed.
    """
    # Strip wrappers outermost-first: bold → underline → italic.
    # Iterate until stable because nested wrappers need multiple passes.
    for _pass in range(5):
        prev = text
        text = _BOLD_RE.sub(r"\1", text)
        text = _UNDERLINE_RE.sub(r"\1", text)
        text = _ITALIC_RE.sub(r"\1", text)
        if text == prev:
            break

    return _unescape_markdown(text)


def _wrap_formatting(text: str, *, bold: bool, italic: bool, underline: bool) -> str:
    """Wrap text in pseudo-Markdown formatting markers.

    Nesting order (outermost to innermost): bold > underline > italic
    This produces: **__..._text_...__**
    """
    if not text:
        return text

    result = text
    if italic:
        result = f"_{result}_"
    if underline:
        result = f"__{result}__"
    if bold:
        result = f"**{result}**"
    return result


def _extract_run_text(run: etree._Element, *, include_del_text: bool = False) -> str:
    """Extract text content from a <w:r> element.

    Concatenates all <w:t> children. Handles <w:br/> as newline
    and <w:tab/> as tab (though tabs are rare in legal docs).

    When *include_del_text* is True, ``<w:delText>`` elements are also
    included.  This is needed for tracked-change-aware extraction where
    the paragraph contains ``<w:del>`` wrappers with ``<w:delText>``
    instead of ``<w:t>``.
    """
    parts: list[str] = []
    for child in run:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else None
        if (tag == "t" and child.text) or (tag == "delText" and include_del_text and child.text):
            parts.append(child.text)
        elif tag == "br":
            parts.append("\n")
        elif tag == "tab":
            parts.append("\t")
    return "".join(parts)


def _format_run(run: etree._Element, *, include_del_text: bool = False) -> str:
    """Extract text from a run and apply pseudo-Markdown formatting.

    Returns the escaped and formatting-wrapped text for a single
    ``<w:r>`` element.  Returns an empty string if the run has no text.
    """
    text = _extract_run_text(run, include_del_text=include_del_text)
    if not text:
        return ""

    rpr_list = xpath(run, "w:rPr")
    rpr = rpr_list[0] if rpr_list else None

    bold = _is_bold(rpr)
    italic = _is_italic(rpr)
    underline = _is_underline(rpr)

    escaped = _escape_markdown(text)
    return _wrap_formatting(escaped, bold=bold, italic=italic, underline=underline)


def paragraph_to_pseudo_markdown(
    paragraph: etree._Element,
    *,
    markup: bool = False,
) -> str:
    """Convert a single <w:p> element to pseudo-Markdown text.

    Walks the runs, extracts text and formatting, and produces a
    pseudo-Markdown string with **, _, __ markers.

    Unicode characters (smart quotes, em-dashes, etc.) are preserved as-is.
    Markdown syntax characters (* and _) in the source text are escaped.
    Multiple consecutive spaces are collapsed to single spaces.

    When *markup* is True, tracked changes are included:

    - Inserted runs (inside ``<w:ins>``) are wrapped with ``++…++``.
    - Deleted runs (inside ``<w:del>``) are wrapped with ``~~…~~``.

    This allows round-trip validation of redlined documents.  When
    *markup* is False (the default), only direct ``<w:r>`` children
    of the paragraph are processed, giving the "original" view.
    """
    segments: list[str] = []

    if not markup:
        # Original behaviour: only direct <w:r> children.
        for run in xpath(paragraph, "w:r"):
            formatted = _format_run(run)
            if formatted:
                segments.append(formatted)
    else:
        # Tracked-change-aware: walk all children in document order.
        for child in paragraph:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

            if tag == "r":
                # Direct run (unchanged text)
                formatted = _format_run(child)
                if formatted:
                    segments.append(formatted)

            elif tag == "ins":
                # Inserted region — collect all runs inside <w:ins>
                ins_parts: list[str] = []
                for run in xpath(child, "w:r"):
                    formatted = _format_run(run)
                    if formatted:
                        ins_parts.append(formatted)
                if ins_parts:
                    segments.append(f"++{''.join(ins_parts)}++")

            elif tag == "del":
                # Deleted region — collect all runs inside <w:del>
                # Deleted runs use <w:delText> instead of <w:t>
                del_parts: list[str] = []
                for run in xpath(child, "w:r"):
                    formatted = _format_run(run, include_del_text=True)
                    if formatted:
                        del_parts.append(formatted)
                if del_parts:
                    segments.append(f"~~{''.join(del_parts)}~~")

    result = "".join(segments)

    # Collapse multiple consecutive spaces to single space
    # (but preserve non-breaking spaces and other unicode whitespace as-is)
    result = re.sub(r"  +", " ", result)

    return result


def document_to_fragments(
    paragraphs: list[etree._Element],
    *,
    markup: bool = False,
) -> list[tuple[int, str]]:
    """Convert a list of paragraph elements to (fragment_id, pseudo_markdown) pairs.

    Args:
        paragraphs: List of <w:p> elements (typically from DocxDocument.paragraphs).
        markup: When True, include tracked-change markers (``++…++`` for
            insertions, ``~~…~~`` for deletions).

    Returns:
        List of (fragment_id, pseudo_markdown) tuples, 1-indexed.
        Empty paragraphs are included with empty string text.
    """
    fragments: list[tuple[int, str]] = []
    for i, para in enumerate(paragraphs, start=1):
        md = paragraph_to_pseudo_markdown(para, markup=markup)
        fragments.append((i, md))
    return fragments


def fragments_to_tagged_text(fragments: list[tuple[int, str]]) -> str:
    """Format fragments as tagged text for LLM consumption.

    Produces:
        <f=1>First paragraph text.</f=1>
        <f=2>Second paragraph text.</f=2>
        ...
    """
    lines: list[str] = []
    for fid, text in fragments:
        lines.append(f"<f={fid}>{text}</f={fid}>")
    return "\n".join(lines)


# --- Table Extraction ---


def _extract_table_info(
    tbl: etree._Element,
    table_id: int,
    *,
    markup: bool = False,
) -> TableInfo | SkippedTableInfo:
    """Extract structured table info from a <w:tbl> element.

    Args:
        tbl: A ``<w:tbl>`` element.
        table_id: 1-based table index in document order.
        markup: When True, include tracked-change markers in cell text.

    Returns:
        TableInfo if the table is simple, or SkippedTableInfo if not.
    """
    is_simple, reason = is_simple_table(tbl)

    if not is_simple:
        return SkippedTableInfo(table_id=table_id, reason=reason)

    rows_count, cols_count = table_dimensions(tbl)

    cells: list[list[CellInfo]] = []
    rows = list(xpath(tbl, "./w:tr"))

    for row_idx, row in enumerate(rows, start=1):
        cell_row: list[CellInfo] = []
        tcs = list(xpath(row, "./w:tc"))

        for col_idx, tc in enumerate(tcs, start=1):
            cell_id = f"{table_id}.{row_idx}.{col_idx}"
            paras = get_cell_paragraphs(tc)

            cell_text_parts: list[str] = []
            for para in paras:
                para_md = paragraph_to_pseudo_markdown(para, markup=markup)
                cell_text_parts.append(para_md)

            cell_text = "\n".join(cell_text_parts)

            cell_info = CellInfo(
                cell_id=cell_id,
                row=row_idx,
                col=col_idx,
                text=cell_text,
            )
            cell_row.append(cell_info)

        cells.append(cell_row)

    return TableInfo(
        table_id=table_id,
        rows=rows_count,
        cols=cols_count,
        cells=cells,
    )


FragmentItem = tuple[int, str] | TableInfo | SkippedTableInfo


def body_to_fragments(
    body_elements: list[etree._Element],
    *,
    markup: bool = False,
) -> list[FragmentItem]:
    """Convert mixed <w:p>/<w:tbl> elements to fragment items.

    Tables and paragraphs share the same 1-based ID space in document order.

    Args:
        body_elements: List of ``<w:p>`` and ``<w:tbl>`` elements
            (typically from DocxDocument.body_elements).
        markup: When True, include tracked-change markers in text.

    Returns:
        List of FragmentItem (either (id, text) for paragraphs or
        TableInfo/SkippedTableInfo for tables).
    """
    items: list[FragmentItem] = []

    for i, el in enumerate(body_elements, start=1):
        tag_local = el.tag.split("}")[-1] if "}" in el.tag else el.tag

        if tag_local == "p":
            md = paragraph_to_pseudo_markdown(el, markup=markup)
            items.append((i, md))

        elif tag_local == "tbl":
            table_info = _extract_table_info(el, table_id=i, markup=markup)
            items.append(table_info)

    return items


def fragments_to_tagged_text_interleaved(items: list[FragmentItem]) -> str:
    """Format mixed paragraph/table fragments as tagged text.

    Produces:
        <f=1>First paragraph text.</f=1>
        <table=2 rows=2 cols=3>
        <cell=2.1.1>Header A</cell=2.1.1>
        <cell=2.1.2>First para
        Second para</cell=2.1.2>
        ...
        </table=2>
        <table=4 skipped reason="contains merged cells"/>
        <f=5>Another paragraph.</f=5>
    """
    lines: list[str] = []

    for item in items:
        if isinstance(item, tuple):
            fid, text = item
            lines.append(f"<f={fid}>{text}</f={fid}>")

        elif isinstance(item, SkippedTableInfo):
            lines.append(f'<table={item.table_id} skipped reason="{item.reason}"/>')

        elif isinstance(item, TableInfo):
            lines.append(f"<table={item.table_id} rows={item.rows} cols={item.cols}>")
            for row in item.cells:
                for cell in row:
                    lines.append(f"<cell={cell.cell_id}>{cell.text}</cell={cell.cell_id}>")
            lines.append(f"</table={item.table_id}>")

    return "\n".join(lines)


def fragments_to_json_interleaved(items: list[FragmentItem]) -> list[dict]:
    """Format mixed paragraph/table fragments as JSON-serializable dicts.

    Returns:
        List of dicts with discriminated union:
        - Paragraph: {"type": "paragraph", "fragment_id": 1, "text": "..."}
        - Table: {"type": "table", "table_id": 2, "rows": 2, "cols": 3,
                  "cells": [[{"cell_id": "2.1.1", "row": 1, "col": 1, "text": "..."}]]}
        - Skipped: {"type": "table", "table_id": 4, "skipped": true,
                    "reason": "contains merged cells"}
    """
    result: list[dict] = []

    for item in items:
        if isinstance(item, tuple):
            fid, text = item
            result.append({"type": "paragraph", "fragment_id": fid, "text": text})

        elif isinstance(item, SkippedTableInfo):
            result.append(
                {
                    "type": "table",
                    "table_id": item.table_id,
                    "skipped": True,
                    "reason": item.reason,
                }
            )

        elif isinstance(item, TableInfo):
            cells_json = [
                [
                    {
                        "cell_id": cell.cell_id,
                        "row": cell.row,
                        "col": cell.col,
                        "text": cell.text,
                    }
                    for cell in row
                ]
                for row in item.cells
            ]
            result.append(
                {
                    "type": "table",
                    "table_id": item.table_id,
                    "rows": item.rows,
                    "cols": item.cols,
                    "cells": cells_json,
                }
            )

    return result
