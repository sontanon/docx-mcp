"""Handler for ``ChangeType.MODIFY`` — word-level tracked changes within a paragraph.

This is the most complex handler.  It:

1. Extracts the raw text from the paragraph's existing XML runs.
2. Converts the new pseudo-Markdown text to raw text (strips formatting
   markers and unescapes).
3. Diffs the old raw text against the new raw text at word granularity.
4. Maps the diff chunks back onto the original XML runs (character-level
   alignment).
5. Rebuilds the paragraph children with proper ``<w:del>`` / ``<w:ins>``
   tracked-change wrappers around the changed regions.
6. Preserves ``<w:pPr>`` and any non-run children (bookmarks, etc.).

Equal regions are emitted as plain ``<w:r>`` elements (unchanged).
Deleted regions are wrapped in ``<w:del>`` with ``<w:delText>``.
Inserted regions are wrapped in ``<w:ins>`` with ``<w:t>``.

.. note::

   The diff operates on **raw text** (not pseudo-Markdown) because
   ``map_diff_to_runs()`` aligns diff characters against the raw run text
   from ``extract_runs()``.  Using pseudo-Markdown would introduce
   formatting markers (``**``, ``_``, ``\\_``) that have no counterpart in
   the XML run text, causing cascading alignment failures.
"""

from __future__ import annotations

from lxml import etree

from docx_mcp.converter import pseudo_markdown_to_raw
from docx_mcp.differ import DmpWordDiffer
from docx_mcp.id_manager import IdManager
from docx_mcp.models import DiffOp, RedlineConfig
from docx_mcp.namespaces import xpath
from docx_mcp.run_ops import (
    TaggedSegment,
    build_run_element,
    build_tracked_change_element,
    extract_runs,
    get_paragraph_text,
    map_diff_to_runs,
)

_differ = DmpWordDiffer()


def handle_modify(
    paragraph: etree._Element,
    new_text: str,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
) -> list[int]:
    """Apply word-level tracked changes to *paragraph* in-place.

    Args:
        paragraph: The ``<w:p>`` element to modify.
        new_text: The new paragraph content in pseudo-Markdown format.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.

    Returns:
        List of annotation IDs used (one per ``<w:del>``/``<w:ins>`` pair).
    """
    author = config.author
    date = config.date_iso()

    # --- 1. Get old text (raw, not pseudo-Markdown) ---
    # The diff must operate on raw text because map_diff_to_runs() aligns
    # diff characters against the raw run text from extract_runs().
    # Pseudo-Markdown markers (**bold**, \_escaped\_) would cause alignment
    # mismatches and garbled output.
    runs = extract_runs(paragraph)
    old_text = get_paragraph_text(runs)

    # --- 2. Convert new pseudo-Markdown to raw text for diffing ---
    new_raw = pseudo_markdown_to_raw(new_text)

    # --- 3. Diff (raw text vs raw text) ---
    diff_chunks = _differ.diff(old_text, new_raw)

    # If no changes, nothing to do
    if all(c.op == DiffOp.EQUAL for c in diff_chunks):
        return []

    # --- 4. Map diff to runs ---
    segments = map_diff_to_runs(diff_chunks, runs)

    # --- 5. Rebuild paragraph children ---
    annotation_ids = _rebuild_paragraph(
        paragraph,
        segments,
        author=author,
        date=date,
        id_manager=id_manager,
    )

    return annotation_ids


def _rebuild_paragraph(
    paragraph: etree._Element,
    segments: list[TaggedSegment],
    *,
    author: str,
    date: str,
    id_manager: IdManager,
) -> list[int]:
    """Replace the paragraph's run children with tracked-change elements.

    Preserves ``<w:pPr>`` and any elements before the first ``<w:r>``.

    Returns the list of annotation IDs used.
    """
    annotation_ids: list[int] = []

    # Save <w:pPr> if present
    ppr = None
    ppr_list = xpath(paragraph, "w:pPr")
    if ppr_list:
        ppr = ppr_list[0]

    # Save non-run children that appear before runs (e.g. bookmarkStart)
    # We'll preserve these by noting what's not a <w:r>
    pre_run_elements: list[etree._Element] = []
    for child in paragraph:
        tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
        if tag_local == "pPr":
            continue  # handled separately
        if tag_local == "r":
            break  # stop at first run
        pre_run_elements.append(child)

    # Clear paragraph children
    for child in list(paragraph):
        paragraph.remove(child)

    # Restore pPr
    if ppr is not None:
        paragraph.append(ppr)

    # Restore pre-run elements
    for el in pre_run_elements:
        paragraph.append(el)

    # Build new children from segments
    i = 0
    while i < len(segments):
        seg = segments[i]

        if seg.op == DiffOp.EQUAL:
            # Plain run — no tracked change wrapper
            r = build_run_element(seg.text, rpr=seg.rpr)
            paragraph.append(r)
            i += 1

        elif seg.op == DiffOp.DELETE:
            # Wrap in <w:del>
            del_id = id_manager.next_id()
            annotation_ids.append(del_id)
            del_el = build_tracked_change_element(
                "del",
                change_id=del_id,
                author=author,
                date_iso=date,
            )
            r = build_run_element(seg.text, rpr=seg.rpr, is_delete=True)
            del_el.append(r)

            # Check if next segment is INSERT (del+ins pair)
            if i + 1 < len(segments) and segments[i + 1].op == DiffOp.INSERT:
                paragraph.append(del_el)
                # Now the insert
                ins_seg = segments[i + 1]
                ins_id = id_manager.next_id()
                annotation_ids.append(ins_id)
                ins_el = build_tracked_change_element(
                    "ins",
                    change_id=ins_id,
                    author=author,
                    date_iso=date,
                )
                r_ins = build_run_element(ins_seg.text, rpr=ins_seg.rpr)
                ins_el.append(r_ins)
                paragraph.append(ins_el)
                i += 2
            else:
                paragraph.append(del_el)
                i += 1

        elif seg.op == DiffOp.INSERT:
            # Standalone insert (not paired with a preceding delete)
            ins_id = id_manager.next_id()
            annotation_ids.append(ins_id)
            ins_el = build_tracked_change_element(
                "ins",
                change_id=ins_id,
                author=author,
                date_iso=date,
            )
            r = build_run_element(seg.text, rpr=seg.rpr)
            ins_el.append(r)
            paragraph.append(ins_el)
            i += 1

    return annotation_ids
