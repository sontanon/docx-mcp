"""Redliner orchestrator — main entry point for applying tracked changes.

Coordinates the full pipeline:

1. Load the document.
2. Build the fragment map.
3. Initialise the ID manager.
4. Sort and validate changes.
5. Dispatch each change to its handler (modify / delete / append_after).
6. Attach comments with justification text.
7. Return the modified document.

Usage::

    from docx_mcp.models import ParagraphChange, ParagraphChangeType, RedlineConfig
    from docx_mcp.redliner import apply_redlines

    changes = [
        ParagraphChange(
            kind="paragraph",
            fragment_id=3,
            change_type=ParagraphChangeType.MODIFY,
            new_text="The Company must provide notice.",
            justification="Strengthened obligation language.",
        ),
        ParagraphChange(
            kind="paragraph",
            fragment_id=5,
            change_type=ParagraphChangeType.DELETE,
            justification="Removed redundant clause.",
        ),
        ParagraphChange(
            kind="paragraph",
            fragment_id=7,
            change_type=ParagraphChangeType.APPEND_AFTER,
            new_text="The foregoing shall survive termination.",
            justification="Added survival provision.",
        ),
    ]

    doc = apply_redlines("input.docx", changes)
    doc.save("output_redlined.docx")
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree

from docx_mcp.comments import add_comment
from docx_mcp.document import DocxDocument
from docx_mcp.handlers.append import handle_append_after
from docx_mcp.handlers.delete import handle_delete
from docx_mcp.handlers.modify import handle_modify
from docx_mcp.id_manager import IdManager
from docx_mcp.models import (
    Change,
    ParagraphChange,
    ParagraphChangeType,
    RedlineConfig,
    TableChange,
)
from docx_mcp.namespaces import xpath
from docx_mcp.table_redliner import apply_table_changes
from docx_mcp.table_utils import is_simple_table

# Tag names of elements that carry visible text inside a run.
_TEXT_TAGS = frozenset({"t", "delText"})


def apply_redlines(
    source: Path | str | bytes,
    changes: list[Change],
    config: RedlineConfig | None = None,
) -> DocxDocument:
    """Apply tracked changes to a .docx document.

    Accepts a unified list of changes that can include both paragraph changes
    (modify, delete, append_after) and table cell changes (modify_cell, clear_cell).

    Args:
        source: Path to a ``.docx`` file, or raw bytes of one.
        changes: List of changes to apply (mix of ParagraphChange and TableChange).
        config: Redline configuration (author, date). Defaults to
            ``RedlineConfig()`` which uses "AI Review" and current time.

    Returns:
        A :class:`DocxDocument` with all changes applied as tracked
        changes with comments. Call ``.save(path)`` or ``.to_bytes()``
        to produce the output file.

    Raises:
        ValueError: If a change references a non-existent ID,
            targets the wrong element type, or is missing required fields.
    """
    if config is None:
        config = RedlineConfig()

    # --- 1. Load ---
    doc = DocxDocument(data=source) if isinstance(source, bytes) else DocxDocument(path=source)

    # --- 1b. Reject documents with pre-existing tracked changes ---
    dirty_parts = doc.has_tracked_changes()
    if dirty_parts:
        locations = ", ".join(dirty_parts)
        msg = (
            f"Document contains pre-existing tracked changes in {locations}. "
            "Please accept or reject all changes before redlining."
        )
        raise ValueError(msg)

    # --- 2. Element map (interleaved paragraphs and tables) ---
    element_map = doc.interleaved_element_map()
    max_id = max(element_map.keys()) if element_map else 0

    # --- 3. Split changes by type ---
    paragraph_changes: list[ParagraphChange] = []
    table_changes: list[TableChange] = []

    for change in changes:
        if isinstance(change, ParagraphChange):
            paragraph_changes.append(change)
        else:  # TableChange
            table_changes.append(change)

    # --- 4. Validate changes ---
    _validate_paragraph_changes(paragraph_changes, max_id, element_map)
    _validate_table_changes(table_changes, max_id, element_map)

    # --- 5. ID manager ---
    id_manager = IdManager(start_after=doc.max_annotation_id())

    # --- 6. Sort and apply paragraph changes ---
    sorted_para_changes = _sort_paragraph_changes(paragraph_changes)

    for change in sorted_para_changes:
        paragraph = element_map[change.fragment_id]

        if change.change_type == ParagraphChangeType.MODIFY:
            assert change.new_text is not None
            annotation_ids = handle_modify(
                paragraph,
                change.new_text,
                id_manager=id_manager,
                config=config,
            )
            # Comment spans tracked-change elements in the paragraph
            if annotation_ids:
                range_els = _get_tracked_change_elements(paragraph)
                add_comment(
                    doc,
                    paragraph,
                    change.justification,
                    id_manager=id_manager,
                    config=config,
                    range_elements=range_els if range_els else None,
                )

        elif change.change_type == ParagraphChangeType.DELETE:
            handle_delete(
                paragraph,
                id_manager=id_manager,
                config=config,
                preserve_paragraph_mark=False,
            )
            # Comment spans the whole paragraph
            add_comment(
                doc,
                paragraph,
                change.justification,
                id_manager=id_manager,
                config=config,
            )

            # Also delete trailing blank paragraphs if requested
            if change.delete_next_blanks > 0:
                _delete_trailing_blanks(
                    paragraph,
                    count=change.delete_next_blanks,
                    fragment_id=change.fragment_id,
                    id_manager=id_manager,
                    config=config,
                )

        elif change.change_type == ParagraphChangeType.APPEND_AFTER:
            assert change.new_text is not None
            new_p, _ = handle_append_after(
                paragraph,
                change.new_text,
                id_manager=id_manager,
                config=config,
                blank_lines_before=change.blank_lines_before,
                blank_lines_after=change.blank_lines_after,
                hyperlink_creator=doc.create_hyperlink_relationship,
            )
            # Comment on the new paragraph
            add_comment(
                doc,
                new_p,
                change.justification,
                id_manager=id_manager,
                config=config,
            )

    # --- 7. Apply table changes ---
    if table_changes:
        apply_table_changes(
            doc,
            table_changes,
            element_map,
            id_manager=id_manager,
            config=config,
        )

    return doc


def _validate_paragraph_changes(
    changes: list[ParagraphChange],
    max_id: int,
    element_map: dict[int, etree._Element],
) -> None:
    """Validate all paragraph changes before applying any.

    Args:
        changes: List of paragraph changes.
        max_id: Maximum valid element ID.
        element_map: Mapping of element_id → element.

    Raises:
        ValueError: If a change is invalid or targets the wrong element type.
    """
    for change in changes:
        if change.fragment_id < 1 or change.fragment_id > max_id:
            msg = (
                f"Paragraph change references fragment_id={change.fragment_id}, "
                f"but document has elements 1..{max_id}"
            )
            raise ValueError(msg)

        # Verify it targets a paragraph, not a table
        el = element_map[change.fragment_id]
        tag_local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag_local != "p":
            msg = (
                f"Paragraph change references fragment_id={change.fragment_id}, "
                f"but element {change.fragment_id} is a <w:{tag_local}>, not <w:p>. "
                f"Use TableChange with cell_id for table changes."
            )
            raise ValueError(msg)


def _validate_table_changes(
    table_changes: list[TableChange],
    max_id: int,
    element_map: dict[int, etree._Element],
) -> None:
    """Validate all table changes before applying any.

    Args:
        table_changes: List of table cell changes.
        max_id: Maximum valid element ID.
        element_map: Mapping of element_id → element.

    Raises:
        ValueError: If a table change is invalid or targets the wrong element type.
    """
    for change in table_changes:
        if change.table_id < 1 or change.table_id > max_id:
            msg = (
                f"Table change references table_id={change.table_id} (cell_id={change.cell_id}), "
                f"but document has elements 1..{max_id}"
            )
            raise ValueError(msg)

        # Verify it targets a table, not a paragraph
        el = element_map.get(change.table_id)
        if el is None:
            msg = (
                f"Table change references table_id={change.table_id} (cell_id={change.cell_id}), "
                f"but no element exists at that position"
            )
            raise ValueError(msg)

        tag_local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        if tag_local != "tbl":
            msg = (
                f"Table change references table_id={change.table_id} (cell_id={change.cell_id}), "
                f"but element {change.table_id} is a <w:{tag_local}>, not <w:tbl>. "
                f"Use ParagraphChange with fragment_id for paragraph changes."
            )
            raise ValueError(msg)

        # Check if table is simple
        is_simple, reason = is_simple_table(el)
        if not is_simple:
            msg = (
                f"Table change for cell_id={change.cell_id} targets "
                f"a non-simple table (table_id={change.table_id}): {reason}"
            )
            raise ValueError(msg)

        # Validate row/col are positive
        if change.row < 1 or change.col < 1:
            msg = (
                f"Table change for cell_id={change.cell_id} has invalid "
                f"row/col (must be positive integers)"
            )
            raise ValueError(msg)


def _sort_paragraph_changes(changes: list[ParagraphChange]) -> list[ParagraphChange]:
    """Sort paragraph changes for correct processing order.

    Ordering rules:
    - Modify changes first (they need original text for diffing).
    - Then delete changes.
    - Then append_after changes (in reverse fragment order, so later
      appends don't shift earlier targets).
    - Within append_after changes to the same fragment, reverse the order
      so the first user-listed append ends up first in the document.
    """
    type_order = {
        ParagraphChangeType.MODIFY: 0,
        ParagraphChangeType.DELETE: 1,
        ParagraphChangeType.APPEND_AFTER: 2,
    }

    # Group append_after changes by fragment_id
    from itertools import groupby

    def _group_key(c: ParagraphChange) -> tuple[int, int]:
        order = type_order[c.change_type]
        fid = -c.fragment_id if c.change_type == ParagraphChangeType.APPEND_AFTER else c.fragment_id
        return (order, fid)

    sorted_by_key = sorted(changes, key=_group_key)

    result: list[ParagraphChange] = []
    for (_order, _fid), group in groupby(sorted_by_key, key=_group_key):
        group_list = list(group)
        if _order == type_order[ParagraphChangeType.APPEND_AFTER]:
            # Reverse within the group so first-listed append is processed last
            # and ends up closest to the reference paragraph
            group_list = list(reversed(group_list))
        result.extend(group_list)

    return result


def _get_tracked_change_elements(paragraph: etree._Element) -> list[etree._Element]:
    """Get all ``<w:del>`` and ``<w:ins>`` direct children of a paragraph."""
    elements: list[etree._Element] = []
    for child in paragraph:
        tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        if tag_local in ("del", "ins"):
            elements.append(child)
    return elements


# ---------------------------------------------------------------------------
# Blank-paragraph helpers
# ---------------------------------------------------------------------------


def _is_blank_paragraph(paragraph: etree._Element) -> bool:
    """Return ``True`` if *paragraph* contains no visible text.

    A paragraph is blank if it has no ``<w:r>`` elements, or all of its
    runs contain only whitespace text (or no text at all).  Tabs, breaks,
    and empty ``<w:t>`` elements are treated as whitespace.
    """
    for run in xpath(paragraph, "w:r"):
        for child in run:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
            if tag in _TEXT_TAGS and child.text and child.text.strip():
                return False
    return True


def _paragraph_preview(paragraph: etree._Element, *, max_len: int = 60) -> str:
    """Extract a short text preview of *paragraph* for error messages."""
    parts: list[str] = []
    for run in xpath(paragraph, "w:r"):
        for child in run:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
            if tag in _TEXT_TAGS and child.text:
                parts.append(child.text)
    text = "".join(parts).strip()
    if len(text) > max_len:
        return text[:max_len] + "..."
    return text


def _delete_trailing_blanks(
    paragraph: etree._Element,
    *,
    count: int,
    fragment_id: int,
    id_manager: IdManager,
    config: RedlineConfig,
) -> None:
    """Mark *count* blank siblings after *paragraph* as tracked deletions.

    Raises:
        ValueError: If a sibling does not exist or is not blank.
    """
    current = paragraph
    for i in range(count):
        next_el = current.getnext()
        if next_el is None:
            msg = (
                f"delete_next_blanks={count} on fragment {fragment_id}, "
                f"but only {i} paragraph(s) follow the deleted paragraph"
            )
            raise ValueError(msg)

        # Verify it's a <w:p>
        next_tag = etree.QName(next_el.tag).localname if isinstance(next_el.tag, str) else ""
        if next_tag != "p":
            msg = (
                f"delete_next_blanks={count} on fragment {fragment_id}, "
                f"but the element at position {i + 1} after the deleted paragraph "
                f"is <w:{next_tag}>, not <w:p>"
            )
            raise ValueError(msg)

        if not _is_blank_paragraph(next_el):
            preview = _paragraph_preview(next_el)
            msg = (
                f"delete_next_blanks={count} on fragment {fragment_id}, "
                f"but the paragraph at position {i + 1} after the deleted paragraph "
                f"is not blank: '{preview}'"
            )
            raise ValueError(msg)

        handle_delete(next_el, id_manager=id_manager, config=config, preserve_paragraph_mark=False)
        current = next_el
