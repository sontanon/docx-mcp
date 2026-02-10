"""Low-level operations on ``<w:r>`` run elements inside a paragraph.

This is the bridge between the word-level diff engine (which operates on
plain-text tokens) and the OOXML tracked-change XML structures.

Key responsibilities:

* **Extract** the text and formatting (``w:rPr``) from each run.
* **Split** a run at a character offset so diff boundaries align with
  run boundaries.
* **Map** diff chunks to runs, splitting as needed, and tag each
  resulting segment with its diff operation (equal / insert / delete).
* **Build** new ``<w:r>`` elements for inserted text, inheriting
  formatting from adjacent runs.
* **Clone** ``w:rPr`` elements (deep copy).
"""

from __future__ import annotations

import copy
from dataclasses import dataclass

from lxml import etree

from docx_mcp.models import DiffChunk, DiffOp
from docx_mcp.namespaces import XML, qn, xpath

# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class RunInfo:
    """Extracted information about a single ``<w:r>`` element.

    Attributes:
        text: The plain text content of the run (from ``<w:t>`` children).
        rpr: The ``<w:rPr>`` element (may be ``None`` if the run has no
            formatting).  This is the *original* element reference.
        element: The original ``<w:r>`` element in the document tree.
    """

    text: str
    rpr: etree._Element | None
    element: etree._Element


@dataclass
class TaggedSegment:
    """A text segment tagged with a diff operation and formatting info.

    Produced by :func:`map_diff_to_runs`.  Each segment becomes one
    ``<w:r>`` (or ``<w:delText>`` run) in the final output.

    Attributes:
        text: The text content for this segment.
        op: The diff operation (equal, insert, or delete).
        rpr: The ``<w:rPr>`` to apply (deep-copied).  ``None`` means no
            formatting (plain text).
    """

    text: str
    op: DiffOp
    rpr: etree._Element | None = None


# ---------------------------------------------------------------------------
# Run extraction
# ---------------------------------------------------------------------------


def extract_runs(paragraph: etree._Element) -> list[RunInfo]:
    """Extract all ``<w:r>`` elements from a paragraph.

    Returns a list of :class:`RunInfo` objects preserving document order.
    Runs with no text content (e.g. field codes, drawing anchors) are
    included with empty ``text`` so that element references stay valid.
    """
    runs: list[RunInfo] = []
    for r_el in xpath(paragraph, "w:r"):
        text_parts: list[str] = []
        for child in r_el:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else None
            if tag == "t" and child.text:
                text_parts.append(child.text)
            elif tag == "br":
                text_parts.append("\n")
            elif tag == "tab":
                text_parts.append("\t")
        rpr_list = xpath(r_el, "w:rPr")
        rpr = rpr_list[0] if rpr_list else None
        runs.append(RunInfo(text="".join(text_parts), rpr=rpr, element=r_el))
    return runs


def get_paragraph_text(runs: list[RunInfo]) -> str:
    """Concatenate the text of all runs into a single string."""
    return "".join(r.text for r in runs)


# ---------------------------------------------------------------------------
# rPr cloning
# ---------------------------------------------------------------------------


def clone_rpr(rpr: etree._Element | None) -> etree._Element | None:
    """Deep-copy a ``<w:rPr>`` element, or return ``None`` if input is ``None``."""
    if rpr is None:
        return None
    return copy.deepcopy(rpr)


# ---------------------------------------------------------------------------
# Run splitting
# ---------------------------------------------------------------------------


def split_run_at(run_info: RunInfo, offset: int) -> tuple[RunInfo, RunInfo]:
    """Split a :class:`RunInfo` at a character *offset*.

    Returns ``(left, right)`` where ``left.text = run_info.text[:offset]``
    and ``right.text = run_info.text[offset:]``.  Both halves receive
    independent deep-copies of the original ``w:rPr``.

    Raises:
        ValueError: If *offset* is out of range ``(0, len(text))``.
    """
    text = run_info.text
    if offset <= 0 or offset >= len(text):
        msg = f"split offset {offset} out of range for run text of length {len(text)}"
        raise ValueError(msg)

    left_text = text[:offset]
    right_text = text[offset:]

    # Build new RunInfo objects — we don't mutate the originals
    left = RunInfo(text=left_text, rpr=clone_rpr(run_info.rpr), element=run_info.element)
    right = RunInfo(text=right_text, rpr=clone_rpr(run_info.rpr), element=run_info.element)
    return left, right


# ---------------------------------------------------------------------------
# Diff → run mapping
# ---------------------------------------------------------------------------


def map_diff_to_runs(
    diff_chunks: list[DiffChunk],
    runs: list[RunInfo],
) -> list[TaggedSegment]:
    """Map word-level diff chunks onto the original paragraph runs.

    Walks the diff chunks and the run list in parallel.  When a diff
    boundary falls inside a run, the run is split.  Each resulting
    segment is tagged with its diff operation.

    For ``DiffOp.EQUAL`` and ``DiffOp.DELETE`` chunks, the text comes
    from the original runs (so we preserve original whitespace / XML
    structure within equal regions).

    For ``DiffOp.INSERT`` chunks, the text is from the diff (new text).
    Formatting is inherited from the nearest adjacent run.

    The mapping works at the *character level* on the concatenated run
    text vs. the concatenated diff text for EQUAL+DELETE chunks.  We
    do NOT re-tokenize; instead we walk character by character through
    both streams.

    Args:
        diff_chunks: Output of the differ (word-level diff).
        runs: Output of :func:`extract_runs`.

    Returns:
        Ordered list of :class:`TaggedSegment` objects ready to be
        turned into ``<w:r>`` elements.
    """
    # Build the "old text" from diff chunks (EQUAL + DELETE) and
    # the concatenated run text.  They should match after whitespace
    # normalisation; we map character-by-character.
    segments: list[TaggedSegment] = []

    # Flatten runs into a list of (char, rpr) pairs
    char_rpr_pairs: list[tuple[str, etree._Element | None]] = []
    for run in runs:
        for ch in run.text:
            char_rpr_pairs.append((ch, run.rpr))

    # Now walk both streams in parallel
    run_pos = 0  # position in char_rpr_pairs

    # Track the last rPr we saw (for INSERT formatting inheritance)
    last_rpr: etree._Element | None = None
    if char_rpr_pairs:
        last_rpr = char_rpr_pairs[0][1]

    for chunk in diff_chunks:
        if chunk.op == DiffOp.INSERT:
            # Inserted text doesn't exist in the original runs.
            # Inherit formatting from the last seen run.
            segments.append(
                TaggedSegment(
                    text=chunk.text,
                    op=DiffOp.INSERT,
                    rpr=clone_rpr(last_rpr),
                )
            )
            continue

        # For EQUAL and DELETE: consume characters from both streams
        chunk_text = chunk.text
        chunk_consumed = 0

        while chunk_consumed < len(chunk_text):
            if run_pos >= len(char_rpr_pairs):
                # Shouldn't happen if diff is correct, but handle gracefully
                remaining = chunk_text[chunk_consumed:]
                segments.append(
                    TaggedSegment(
                        text=remaining,
                        op=chunk.op,
                        rpr=clone_rpr(last_rpr),
                    )
                )
                chunk_consumed = len(chunk_text)
                break

            # Skip whitespace alignment between diff text and run text.
            # The diff text has single spaces between words, but the
            # original runs may have different whitespace.
            diff_char = chunk_text[chunk_consumed]
            run_char, run_rpr = char_rpr_pairs[run_pos]

            # If both are the same character, consume both
            if diff_char == run_char:
                # Find the longest contiguous segment with same rPr and same chunk
                seg_chars: list[str] = [run_char]
                chunk_consumed += 1
                run_pos += 1
                last_rpr = run_rpr

                while chunk_consumed < len(chunk_text) and run_pos < len(char_rpr_pairs):
                    dc = chunk_text[chunk_consumed]
                    rc, rp = char_rpr_pairs[run_pos]

                    if dc == rc and rp is run_rpr:
                        # Same formatting, same character — extend segment
                        seg_chars.append(rc)
                        chunk_consumed += 1
                        run_pos += 1
                    elif dc == rc and rp is not run_rpr:
                        # Character matches but formatting changed — break
                        break
                    elif dc == " " and rc != " ":
                        # Diff has a space (word separator) but run doesn't.
                        # This happens when the original had no space between
                        # tokens (rare). Skip the diff space.
                        chunk_consumed += 1
                    elif dc != " " and rc == " ":
                        # Run has a space but diff doesn't.
                        # Consume the run space as part of this segment.
                        seg_chars.append(rc)
                        run_pos += 1
                    else:
                        break

                segments.append(
                    TaggedSegment(
                        text="".join(seg_chars),
                        op=chunk.op,
                        rpr=clone_rpr(run_rpr),
                    )
                )

            elif diff_char == " " and run_char != " ":
                # Diff space = word boundary, but run has no space here.
                # Just skip the diff space.
                chunk_consumed += 1

            elif diff_char != " " and run_char == " ":
                # Run has whitespace that the diff doesn't see (it was
                # normalised during tokenization).  Consume it into the
                # current segment as whitespace.
                # Emit the space as part of the current op
                segments.append(
                    TaggedSegment(
                        text=run_char,
                        op=chunk.op,
                        rpr=clone_rpr(run_rpr),
                    )
                )
                run_pos += 1
                last_rpr = run_rpr

            else:
                # Characters differ — alignment error.  This shouldn't
                # normally happen.  Skip both and hope for re-alignment.
                chunk_consumed += 1
                run_pos += 1

    # If there's remaining run text (e.g. trailing whitespace), emit
    # as EQUAL segments.
    while run_pos < len(char_rpr_pairs):
        ch, rpr = char_rpr_pairs[run_pos]
        segments.append(TaggedSegment(text=ch, op=DiffOp.EQUAL, rpr=clone_rpr(rpr)))
        run_pos += 1
        last_rpr = rpr

    return _merge_tagged_segments(segments)


def _merge_tagged_segments(segments: list[TaggedSegment]) -> list[TaggedSegment]:
    """Merge adjacent segments with the same op and same rPr identity."""
    if not segments:
        return segments
    merged: list[TaggedSegment] = [segments[0]]
    for seg in segments[1:]:
        prev = merged[-1]
        if seg.op == prev.op and seg.rpr is prev.rpr:
            merged[-1] = TaggedSegment(
                text=prev.text + seg.text,
                op=seg.op,
                rpr=seg.rpr,
            )
        else:
            merged.append(seg)
    return merged


# ---------------------------------------------------------------------------
# XML element builders
# ---------------------------------------------------------------------------


def build_run_element(
    text: str,
    rpr: etree._Element | None = None,
    *,
    is_delete: bool = False,
) -> etree._Element:
    """Build a ``<w:r>`` element with text content.

    Args:
        text: The text content for the run.
        rpr: Optional ``<w:rPr>`` element to include (will be deep-copied).
        is_delete: If ``True``, use ``<w:delText>`` instead of ``<w:t>``.

    Returns:
        A new ``<w:r>`` element.
    """
    r_el = etree.SubElement(etree.Element("dummy"), qn("w", "r"))
    r_el = _detach(r_el)

    if rpr is not None:
        r_el.append(clone_rpr(rpr))

    text_tag = qn("w", "delText") if is_delete else qn("w", "t")
    t_el = etree.SubElement(r_el, text_tag)
    t_el.text = text

    # Preserve leading/trailing whitespace
    if text and (text[0] == " " or text[-1] == " " or "\t" in text):
        t_el.set(f"{{{XML}}}space", "preserve")

    return r_el


def build_tracked_change_element(
    tag_local: str,
    *,
    change_id: int,
    author: str,
    date_iso: str,
) -> etree._Element:
    """Build a tracked-change wrapper element (``w:ins`` or ``w:del``).

    Args:
        tag_local: ``"ins"`` or ``"del"``.
        change_id: Unique annotation ID.
        author: Author name.
        date_iso: ISO 8601 timestamp string.

    Returns:
        A new ``<w:ins>`` or ``<w:del>`` element with the required attributes.
    """
    el = etree.Element(qn("w", tag_local))
    el.set(qn("w", "id"), str(change_id))
    el.set(qn("w", "author"), author)
    el.set(qn("w", "date"), date_iso)
    return el


def _detach(element: etree._Element) -> etree._Element:
    """Remove *element* from its parent (if any) and return it."""
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)
    return element
