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
from collections.abc import Callable

from lxml import etree

from docx_mcp.document import DocxDocument
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
    if not text or not text.strip():
        return ""

    rpr_list = xpath(run, "w:rPr")
    rpr = rpr_list[0] if rpr_list else None

    bold = _is_bold(rpr)
    italic = _is_italic(rpr)
    underline = _is_underline(rpr)

    escaped = _escape_markdown(text)
    return _wrap_formatting(escaped, bold=bold, italic=italic, underline=underline)


def _collect_formatted_from_container(
    container: etree._Element,
    *,
    include_del_text: bool = False,
    hyperlink_resolver: Callable[[str], str | None] | None = None,
) -> list[str]:
    """Collect formatted text segments from runs and hyperlinks inside *container*.

    Returns a list of formatted strings (already escaped and wrapped).
    """
    parts: list[str] = []
    for child in container:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        if tag == "r":
            formatted = _format_run(child, include_del_text=include_del_text)
            if formatted:
                parts.append(formatted)
        elif tag == "hyperlink":
            rel_id = child.get(qn("r", "id"))
            url = hyperlink_resolver(rel_id) if hyperlink_resolver and rel_id else None
            link_parts: list[str] = []
            for run in xpath(child, "w:r"):
                formatted = _format_run(run, include_del_text=include_del_text)
                if formatted:
                    link_parts.append(formatted)
            if link_parts:
                link_text = "".join(link_parts)
                if url:
                    parts.append(f"[{link_text}]({url})")
                else:
                    parts.append(f"[{link_text}]")
    return parts


def paragraph_to_pseudo_markdown(
    paragraph: etree._Element,
    *,
    markup: bool = False,
    hyperlink_resolver: Callable[[str], str | None] | None = None,
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

    Args:
        paragraph: The ``<w:p>`` element to convert.
        markup: When True, include tracked-change markers.
        hyperlink_resolver: Optional callable that takes a relationship ID
            and returns the target URL. If provided, hyperlinks are rendered
            as ``[text](url)``.
    """
    segments: list[str] = []

    if not markup:
        segments = _collect_formatted_from_container(
            paragraph, hyperlink_resolver=hyperlink_resolver
        )
    else:
        # Tracked-change-aware: walk all children in document order.
        for child in paragraph:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

            if tag == "r":
                # Direct run (unchanged text)
                formatted = _format_run(child)
                if formatted:
                    segments.append(formatted)

            elif tag == "hyperlink":
                # Direct hyperlink (unchanged text)
                rel_id = child.get(qn("r", "id"))
                url = hyperlink_resolver(rel_id) if hyperlink_resolver and rel_id else None
                link_parts: list[str] = []
                for run in xpath(child, "w:r"):
                    formatted = _format_run(run)
                    if formatted:
                        link_parts.append(formatted)
                if link_parts:
                    link_text = "".join(link_parts)
                    if url:
                        segments.append(f"[{link_text}]({url})")
                    else:
                        segments.append(f"[{link_text}]")

            elif tag == "ins":
                # Inserted region — collect all runs/hyperlinks inside <w:ins>
                ins_parts = _collect_formatted_from_container(
                    child,
                    hyperlink_resolver=hyperlink_resolver,
                )
                if ins_parts:
                    segments.append(f"++{''.join(ins_parts)}++")

            elif tag == "del":
                # Deleted region — collect all runs/hyperlinks inside <w:del>
                del_parts = _collect_formatted_from_container(
                    child,
                    include_del_text=True,
                    hyperlink_resolver=hyperlink_resolver,
                )
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
    hyperlink_resolver: Callable[[str], str | None] | None = None,
) -> TableInfo | SkippedTableInfo:
    """Extract structured table info from a <w:tbl> element.

    Supports simple tables and tables with clean horizontal/vertical merges.
    Tables with nested tables or malformed merges are skipped.

    Args:
        tbl: A ``<w:tbl>`` element.
        table_id: 1-based table index in document order.
        markup: When True, include tracked-change markers in cell text.
        hyperlink_resolver: Optional callable to resolve hyperlink URLs.

    Returns:
        TableInfo if the table is extractable, or SkippedTableInfo if not.
    """
    from docx_mcp.table_utils import build_table_grid, table_dimensions

    # Check for nested tables first (we can't handle these)
    for row in xpath(tbl, "./w:tr"):
        for cell in xpath(row, "./w:tc"):
            nested = xpath(cell, ".//w:tbl")
            if nested:
                return SkippedTableInfo(
                    table_id=table_id,
                    reason=f"table {table_id}, cell contains nested table",
                )

    # Try to build a logical grid (handles merges)
    grid, error = build_table_grid(tbl, table_id=table_id)
    if error:
        return SkippedTableInfo(table_id=table_id, reason=error)

    rows_count, cols_count = table_dimensions(tbl)

    cells: list[list[CellInfo]] = []
    for row_idx, grid_row in enumerate(grid, start=1):
        cell_row: list[CellInfo] = []
        for col_idx, grid_cell in enumerate(grid_row, start=1):
            if grid_cell is None:
                # Should not happen in a valid grid, but handle gracefully
                cell_info = CellInfo(
                    cell_id=f"{table_id}.{row_idx}.{col_idx}",
                    row=row_idx,
                    col=col_idx,
                    text="",
                    span=0,
                    vspan=0,
                )
                cell_row.append(cell_info)
                continue

            if grid_cell.is_spanned_over:
                # Spanned-over cell: empty text, span=0, vspan=0
                cell_info = CellInfo(
                    cell_id=f"{table_id}.{row_idx}.{col_idx}",
                    row=row_idx,
                    col=col_idx,
                    text="",
                    span=grid_cell.span,
                    vspan=grid_cell.vspan,
                )
                cell_row.append(cell_info)
                continue

            # Real cell: extract text
            paras = get_cell_paragraphs(grid_cell.element)
            cell_text_parts: list[str] = []
            for para in paras:
                para_md = paragraph_to_pseudo_markdown(
                    para, markup=markup, hyperlink_resolver=hyperlink_resolver
                )
                cell_text_parts.append(para_md)

            cell_text = "\n".join(cell_text_parts)

            cell_info = CellInfo(
                cell_id=f"{table_id}.{row_idx}.{col_idx}",
                row=row_idx,
                col=col_idx,
                text=cell_text,
                span=grid_cell.span,
                vspan=grid_cell.vspan,
            )
            cell_row.append(cell_info)

        cells.append(cell_row)

    return TableInfo(
        table_id=table_id,
        rows=rows_count,
        cols=cols_count,
        cells=cells,
    )


FragmentItem = tuple[str, str] | TableInfo | SkippedTableInfo


class FragmentResult:
    """Result of converting body elements to fragments.

    Attributes:
        items: List of FragmentItem (paragraphs, tables, skipped tables).
        skipped_elements: List of dicts describing skipped content.
    """

    def __init__(self) -> None:
        self.items: list[FragmentItem] = []
        self.skipped_elements: list[dict] = []


def _is_para_empty(para: etree._Element) -> bool:
    """Check if a <w:p> has no visible text (helper for collapse_empty)."""
    for t in xpath(para, ".//w:t"):
        if t.text and t.text.strip():
            return False
    for dt in xpath(para, ".//w:delText"):
        if dt.text and dt.text.strip():
            return False
    return True


def body_to_fragments(
    body_elements: list[etree._Element],
    *,
    markup: bool = False,
    hyperlink_resolver: Callable[[str], str | None] | None = None,
    collapse_empty: bool = False,
) -> FragmentResult:
    """Convert mixed <w:p>/<w:tbl> elements to fragment items.

    Tables and paragraphs share the same 1-based ID space in document order.

    Args:
        body_elements: List of ``<w:p>`` and ``<w:tbl>`` elements
            (typically from DocxDocument.body_elements).
        markup: When True, include tracked-change markers in text.
        hyperlink_resolver: Optional callable to resolve hyperlink URLs.
        collapse_empty: When True, skip empty ``<w:p>`` elements.

    Returns:
        :class:`FragmentResult` with ``items`` (FragmentItem list) and
        ``skipped_elements`` (metadata about skipped content).
    """
    result = FragmentResult()

    idx = 0
    for el in body_elements:
        tag_local = el.tag.split("}")[-1] if "}" in el.tag else el.tag

        if tag_local == "p":
            if collapse_empty and _is_para_empty(el):
                continue
            idx += 1
            md = paragraph_to_pseudo_markdown(
                el, markup=markup, hyperlink_resolver=hyperlink_resolver
            )
            result.items.append((str(idx), md))
            # Check for images in the paragraph (<w:drawing>/<w:pict> are inside <w:r>)
            if xpath(el, ".//w:drawing") or xpath(el, ".//w:pict"):
                result.skipped_elements.append(
                    {
                        "type": "image",
                        "location": f"body paragraph {idx}",
                        "description": "inline image",
                    }
                )

        elif tag_local == "tbl":
            idx += 1
            table_info = _extract_table_info(
                el, table_id=idx, markup=markup, hyperlink_resolver=hyperlink_resolver
            )
            result.items.append(table_info)
            if isinstance(table_info, SkippedTableInfo):
                result.skipped_elements.append(
                    {
                        "type": "merged_table"
                        if "merge" in table_info.reason.lower()
                        else "skipped_table",
                        "table_id": table_info.table_id,
                        "reason": table_info.reason,
                    }
                )
            else:
                # Check for nested tables
                nested = xpath(el, ".//w:tbl")
                if len(nested) > 1:  # includes self
                    result.skipped_elements.append(
                        {
                            "type": "nested_table",
                            "table_id": table_info.table_id,
                            "reason": "contains nested table",
                        }
                    )

    return result


def full_to_fragments(
    doc: DocxDocument,
    markup: bool = False,
    hyperlink_resolver: Callable[[str], str | None] | None = None,
    collapse_empty: bool = False,
) -> FragmentResult:
    """Convert all extractable content (body, headers, footers) to fragments.

    Body paragraphs and tables are included with their normal numeric IDs.
    Header and footer paragraphs use prefixed IDs (``header_1.1``, ``footer_2.1``).
    Tables inside headers/footers are reported as skipped elements.

    Args:
        doc: :class:`DocxDocument` instance.
        markup: When True, include tracked-change markers in text.
        hyperlink_resolver: Optional callable to resolve hyperlink URLs.
        collapse_empty: When True, skip empty ``<w:p>`` elements.

    Returns:
        :class:`FragmentResult` with all fragments and skipped elements.
    """
    from docx_mcp.document import DocxDocument

    assert isinstance(doc, DocxDocument)

    # Start with body content
    result = body_to_fragments(
        doc.body_elements,
        markup=markup,
        hyperlink_resolver=hyperlink_resolver,
        collapse_empty=collapse_empty,
    )

    # Header paragraphs
    for part_idx, (_rel_id, tree) in enumerate(doc.header_trees.items(), start=1):
        para_idx = 0
        for child in tree:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local == "p":
                if collapse_empty and _is_para_empty(child):
                    continue
                para_idx += 1
                fid = f"header_{part_idx}.{para_idx}"
                md = paragraph_to_pseudo_markdown(
                    child, markup=markup, hyperlink_resolver=hyperlink_resolver
                )
                result.items.append((fid, md))
                # Check for images
                if xpath(child, ".//w:drawing") or xpath(child, ".//w:pict"):
                    result.skipped_elements.append(
                        {
                            "type": "image",
                            "location": fid,
                            "description": "inline image in header",
                        }
                    )
            elif tag_local == "tbl":
                result.skipped_elements.append(
                    {
                        "type": "table",
                        "location": f"header_{part_idx}",
                        "reason": "tables in headers are not editable in this version",
                    }
                )

    # Footer paragraphs
    for part_idx, (_rel_id, tree) in enumerate(doc.footer_trees.items(), start=1):
        para_idx = 0
        for child in tree:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local == "p":
                if collapse_empty and _is_para_empty(child):
                    continue
                para_idx += 1
                fid = f"footer_{part_idx}.{para_idx}"
                md = paragraph_to_pseudo_markdown(
                    child, markup=markup, hyperlink_resolver=hyperlink_resolver
                )
                result.items.append((fid, md))
                if xpath(child, ".//w:drawing") or xpath(child, ".//w:pict"):
                    result.skipped_elements.append(
                        {
                            "type": "image",
                            "location": fid,
                            "description": "inline image in footer",
                        }
                    )
            elif tag_local == "tbl":
                result.skipped_elements.append(
                    {
                        "type": "table",
                        "location": f"footer_{part_idx}",
                        "reason": "tables in footers are not editable in this version",
                    }
                )

    return result


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
                    attrs = ""
                    if cell.span != 1:
                        attrs += f' span="{cell.span}"'
                    if cell.vspan != 1:
                        attrs += f' vspan="{cell.vspan}"'
                    lines.append(
                        f"<cell={cell.cell_id}{attrs}>{cell.text}</cell={cell.cell_id}>"
                    )
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
            cells_json = []
            for row in item.cells:
                row_json = []
                for cell in row:
                    cell_dict: dict = {
                        "cell_id": cell.cell_id,
                        "row": cell.row,
                        "col": cell.col,
                        "text": cell.text,
                    }
                    if cell.span != 1:
                        cell_dict["span"] = cell.span
                    if cell.vspan != 1:
                        cell_dict["vspan"] = cell.vspan
                    row_json.append(cell_dict)
                cells_json.append(row_json)

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
