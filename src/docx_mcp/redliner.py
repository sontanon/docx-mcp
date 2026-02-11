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

    from docx_mcp.models import Change, ChangeType, RedlineConfig
    from docx_mcp.redliner import apply_redlines

    changes = [
        Change(fragment_id=3, change_type=ChangeType.MODIFY,
               new_text="The Company must provide notice.",
               justification="Strengthened obligation language."),
        Change(fragment_id=5, change_type=ChangeType.DELETE,
               justification="Removed redundant clause."),
        Change(fragment_id=7, change_type=ChangeType.APPEND_AFTER,
               new_text="The foregoing shall survive termination.",
               justification="Added survival provision."),
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
from docx_mcp.models import Change, ChangeType, RedlineConfig
from docx_mcp.namespaces import xpath

# Tag names of elements that carry visible text inside a run.
_TEXT_TAGS = frozenset({"t", "delText"})


def apply_redlines(
    source: Path | str | bytes,
    changes: list[Change],
    config: RedlineConfig | None = None,
) -> DocxDocument:
    """Apply tracked changes to a .docx document.

    Args:
        source: Path to a ``.docx`` file, or raw bytes of one.
        changes: List of changes to apply.
        config: Redline configuration (author, date).  Defaults to
            ``RedlineConfig()`` which uses "AI Review" and current time.

    Returns:
        A :class:`DocxDocument` with all changes applied as tracked
        changes with comments.  Call ``.save(path)`` or ``.to_bytes()``
        to produce the output file.

    Raises:
        ValueError: If a change references a non-existent fragment ID,
            or if a MODIFY/APPEND_AFTER change is missing ``new_text``.
    """
    if config is None:
        config = RedlineConfig()

    # --- 1. Load ---
    doc = DocxDocument(data=source) if isinstance(source, bytes) else DocxDocument(path=source)

    # --- 2. Fragment map ---
    fragment_map = doc.fragment_map()
    max_fid = max(fragment_map.keys()) if fragment_map else 0

    # --- 3. Validate changes ---
    _validate_changes(changes, max_fid)

    # --- 4. ID manager ---
    id_manager = IdManager(start_after=doc.max_annotation_id())

    # --- 5. Sort changes ---
    # Process in document order.  Within the same fragment:
    #   modify before delete (so we can diff the original text)
    #   delete before append_after (so appended para goes after the deleted one)
    # Process in reverse fragment order for append_after to maintain correct
    # positioning (appending after fragment 5 then 3 keeps positions stable).
    sorted_changes = _sort_changes(changes)

    # --- 6. Apply changes ---
    for change in sorted_changes:
        paragraph = fragment_map[change.fragment_id]

        if change.change_type == ChangeType.MODIFY:
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

        elif change.change_type == ChangeType.DELETE:
            handle_delete(
                paragraph,
                id_manager=id_manager,
                config=config,
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

        elif change.change_type == ChangeType.APPEND_AFTER:
            assert change.new_text is not None
            new_p, _ = handle_append_after(
                paragraph,
                change.new_text,
                id_manager=id_manager,
                config=config,
                blank_lines_before=change.blank_lines_before,
                blank_lines_after=change.blank_lines_after,
            )
            # Comment on the new paragraph
            add_comment(
                doc,
                new_p,
                change.justification,
                id_manager=id_manager,
                config=config,
            )

    return doc


def _validate_changes(changes: list[Change], max_fragment_id: int) -> None:
    """Validate all changes before applying any."""
    for change in changes:
        if change.fragment_id < 1 or change.fragment_id > max_fragment_id:
            msg = (
                f"Change references fragment_id={change.fragment_id}, "
                f"but document has fragments 1..{max_fragment_id}"
            )
            raise ValueError(msg)

        if (
            change.change_type in (ChangeType.MODIFY, ChangeType.APPEND_AFTER)
            and change.new_text is None
        ):
            msg = (
                f"Change type {change.change_type.value} on "
                f"fragment {change.fragment_id} requires new_text"
            )
            raise ValueError(msg)

        # blank_lines_before / blank_lines_after are only valid with append_after
        if change.change_type != ChangeType.APPEND_AFTER and (
            change.blank_lines_before > 0 or change.blank_lines_after > 0
        ):
            msg = (
                f"blank_lines_before/blank_lines_after are only valid with "
                f"append_after, but fragment {change.fragment_id} has "
                f"change_type={change.change_type.value}"
            )
            raise ValueError(msg)

        # delete_next_blanks is only valid with delete
        if change.change_type != ChangeType.DELETE and change.delete_next_blanks > 0:
            msg = (
                f"delete_next_blanks is only valid with delete, "
                f"but fragment {change.fragment_id} has "
                f"change_type={change.change_type.value}"
            )
            raise ValueError(msg)


def _sort_changes(changes: list[Change]) -> list[Change]:
    """Sort changes for correct processing order.

    Ordering rules:
    - Modify changes first (they need original text for diffing).
    - Then delete changes.
    - Then append_after changes (in reverse fragment order, so later
      appends don't shift earlier targets).
    """
    type_order = {
        ChangeType.MODIFY: 0,
        ChangeType.DELETE: 1,
        ChangeType.APPEND_AFTER: 2,
    }

    def sort_key(c: Change) -> tuple[int, int]:
        order = type_order[c.change_type]
        # For append_after, reverse fragment order (high IDs first)
        fid = -c.fragment_id if c.change_type == ChangeType.APPEND_AFTER else c.fragment_id
        return (order, fid)

    return sorted(changes, key=sort_key)


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

        handle_delete(next_el, id_manager=id_manager, config=config)
        current = next_el
