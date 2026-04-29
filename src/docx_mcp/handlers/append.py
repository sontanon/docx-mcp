"""Handler for ``ChangeType.APPEND_AFTER`` — insert a new paragraph.

OOXML tracked insertion of a new paragraph requires:

1. A new ``<w:p>`` element inserted after the reference paragraph.
2. The paragraph properties must contain ``<w:pPr><w:rPr><w:ins .../>``
   to mark the *paragraph mark* as inserted.
3. All content runs must be wrapped in ``<w:ins>`` elements.
4. The new paragraph copies ``<w:pPr>`` (style, numbering, etc.) from
   the reference paragraph to maintain document consistency.
5. Run-level formatting (font, size, etc.) is inherited from the reference
   paragraph's runs so that appended text matches the surrounding document.

The new text is provided as pseudo-Markdown and must be parsed into
formatted runs.
"""

from __future__ import annotations

import copy
import re
from collections.abc import Callable

from lxml import etree

from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig
from docx_mcp.namespaces import qn, xpath
from docx_mcp.run_ops import build_run_element


def handle_append_after(
    reference_paragraph: etree._Element,
    new_text: str,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
    blank_lines_before: int = 0,
    blank_lines_after: int = 0,
    hyperlink_creator: Callable[[str], str] | None = None,
) -> tuple[etree._Element, int]:
    """Insert a new tracked-change paragraph after *reference_paragraph*.

    Args:
        reference_paragraph: The ``<w:p>`` element after which to insert.
        new_text: The new paragraph content in pseudo-Markdown format.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.
        blank_lines_before: Number of blank paragraphs to insert between
            the reference and the new content paragraph.
        blank_lines_after: Number of blank paragraphs to insert after the
            new content paragraph.

    Returns:
        A tuple of ``(new_paragraph, ins_id)`` — the new content element
        and its annotation ID.
    """
    ins_id = id_manager.next_id()
    author = config.author
    date = config.date_iso()

    # --- 1. Derive formatting from reference paragraph ---
    ref_ppr = xpath(reference_paragraph, "w:pPr")
    base_ppr = copy.deepcopy(ref_ppr[0]) if ref_ppr else None
    base_rpr = _extract_base_rpr(reference_paragraph)

    # --- 2. Build the content paragraph ---
    new_p = _build_inserted_paragraph(
        new_text,
        ins_id=ins_id,
        base_ppr=base_ppr,
        base_rpr=base_rpr,
        id_manager=id_manager,
        author=author,
        date=date,
        hyperlink_creator=hyperlink_creator,
    )

    # --- 3. Insert into the tree: blanks_before → content → blanks_after ---
    # addnext inserts immediately after the reference. We insert
    # blanks_before first (closest to reference), then content, then
    # blanks_after (furthest from reference), updating the anchor each time.
    anchor = reference_paragraph

    # Insert blank lines *before* the content paragraph
    for _ in range(blank_lines_before):
        blank_p = _build_blank_inserted_paragraph(
            base_ppr=base_ppr,
            id_manager=id_manager,
            author=author,
            date=date,
        )
        anchor.addnext(blank_p)
        anchor = blank_p

    # Insert the content paragraph
    anchor.addnext(new_p)
    anchor = new_p

    # Insert blank lines *after* the content paragraph
    for _ in range(blank_lines_after):
        blank_p = _build_blank_inserted_paragraph(
            base_ppr=base_ppr,
            id_manager=id_manager,
            author=author,
            date=date,
        )
        anchor.addnext(blank_p)
        anchor = blank_p

    return new_p, ins_id


# ---------------------------------------------------------------------------
# Formatting inheritance
# ---------------------------------------------------------------------------


def _extract_base_rpr(paragraph: etree._Element) -> etree._Element | None:
    """Extract a base ``<w:rPr>`` from *paragraph* for formatting inheritance.

    Strategy (first match wins):

    1. The ``<w:rPr>`` of the first run that contains text — this represents
       the formatting actually visible in the paragraph.
    2. The ``<w:rPr>`` inside ``<w:pPr>`` — the paragraph-mark formatting,
       which often carries the font even when no runs exist.
    3. ``None`` if neither is available.

    The returned element is a **deep copy** safe to mutate.
    """
    # Try runs first
    for run in xpath(paragraph, "w:r"):
        # Check if this run has any text content
        has_text = False
        for child in run:
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
            if tag in ("t", "tab", "br") and (tag != "t" or child.text):
                has_text = True
                break
        if has_text:
            rpr_list = xpath(run, "w:rPr")
            if rpr_list:
                return copy.deepcopy(rpr_list[0])

    # Fall back to paragraph-mark rPr
    ppr_rpr_list = xpath(paragraph, "w:pPr/w:rPr")
    if ppr_rpr_list:
        rpr = copy.deepcopy(ppr_rpr_list[0])
        # Remove tracked-change marks (ins/del) — they're not formatting
        for mark in xpath(rpr, "w:ins") + xpath(rpr, "w:del"):
            rpr.remove(mark)
        return rpr if len(rpr) > 0 else None

    return None


def _merge_rpr(
    base: etree._Element | None,
    *,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
) -> etree._Element | None:
    """Merge pseudo-Markdown formatting onto a base ``<w:rPr>``.

    Starts from a deep copy of *base* (or a fresh ``<w:rPr>`` if base is
    ``None`` and formatting is requested).  Then sets or clears the bold,
    italic, and underline elements.

    Returns ``None`` if the result would be an empty ``<w:rPr>``.
    """
    if base is not None:
        result = copy.deepcopy(base)
    elif bold or italic or underline:
        result = etree.Element(qn("w", "rPr"))
    else:
        return None

    # Bold: set or remove <w:b>
    existing_b = xpath(result, "w:b")
    if bold and not existing_b:
        # Insert <w:b> early (convention: b comes before i, u)
        etree.SubElement(result, qn("w", "b"))
    elif not bold and existing_b:
        for el in existing_b:
            result.remove(el)

    # Italic: set or remove <w:i>
    existing_i = xpath(result, "w:i")
    if italic and not existing_i:
        etree.SubElement(result, qn("w", "i"))
    elif not italic and existing_i:
        for el in existing_i:
            result.remove(el)

    # Underline: set or remove <w:u>
    existing_u = xpath(result, "w:u")
    if underline and not existing_u:
        u = etree.SubElement(result, qn("w", "u"))
        u.set(qn("w", "val"), "single")
    elif not underline and existing_u:
        for el in existing_u:
            result.remove(el)

    return result if len(result) > 0 else None


# ---------------------------------------------------------------------------
# Paragraph builders
# ---------------------------------------------------------------------------


def _build_inserted_paragraph(
    text: str,
    *,
    ins_id: int,
    base_ppr: etree._Element | None,
    base_rpr: etree._Element | None,
    id_manager: IdManager,
    author: str,
    date: str,
    hyperlink_creator: Callable[[str], str] | None,
) -> etree._Element:
    """Build a ``<w:p>`` with tracked-insertion markup and content runs.

    Args:
        text: New paragraph content in pseudo-Markdown.
        ins_id: Annotation ID for the paragraph-mark insertion.
        base_ppr: Deep copy of the reference paragraph's ``<w:pPr>``
            (may be ``None``).
        base_rpr: Deep copy of the reference paragraph's run formatting
            (may be ``None``).
        id_manager: ID allocator for the ``<w:ins>`` wrapper.
        author: Author for tracked-change attributes.
        date: ISO 8601 date for tracked-change attributes.

    Returns:
        A new ``<w:p>`` element ready to insert into the document tree.
    """
    new_p = etree.Element(qn("w", "p"))

    # --- Paragraph properties ---
    if base_ppr is not None:
        new_ppr = copy.deepcopy(base_ppr)
        # Remove any existing tracked-change marks from the copied pPr
        for old_mark in xpath(new_ppr, "w:rPr/w:ins") + xpath(new_ppr, "w:rPr/w:del"):
            old_mark.getparent().remove(old_mark)
    else:
        new_ppr = etree.Element(qn("w", "pPr"))

    # Mark the paragraph mark as inserted
    ppr_rpr = xpath(new_ppr, "w:rPr")
    rpr = ppr_rpr[0] if ppr_rpr else etree.SubElement(new_ppr, qn("w", "rPr"))

    ins_mark = etree.SubElement(rpr, qn("w", "ins"))
    ins_mark.set(qn("w", "id"), str(ins_id))
    ins_mark.set(qn("w", "author"), author)
    ins_mark.set(qn("w", "date"), date)

    new_p.append(new_ppr)

    # --- Content runs inside <w:ins> ---
    ins_wrapper = etree.SubElement(new_p, qn("w", "ins"))
    ins_wrapper.set(qn("w", "id"), str(ins_id))
    ins_wrapper.set(qn("w", "author"), author)
    ins_wrapper.set(qn("w", "date"), date)

    run_specs = _parse_pseudo_markdown(text)
    for run_text, _md_rpr, md_bold, md_italic, md_underline, md_url in run_specs:
        # Merge pseudo-Markdown formatting onto the inherited base rPr
        merged = _merge_rpr(
            base_rpr,
            bold=md_bold,
            italic=md_italic,
            underline=md_underline,
        )
        rel_id: str | None = None
        if md_url is not None:
            if md_url and hyperlink_creator is not None:
                rel_id = hyperlink_creator(md_url)
            elif md_url == "":
                # [text] without URL — no hyperlink created (plain text)
                pass
        r = build_run_element(run_text, rpr=merged, hyperlink_rel_id=rel_id)
        ins_wrapper.append(r)

    return new_p


def _build_blank_inserted_paragraph(
    *,
    base_ppr: etree._Element | None,
    id_manager: IdManager,
    author: str,
    date: str,
) -> etree._Element:
    """Build an empty ``<w:p>`` marked as a tracked insertion.

    Used for inserting blank separator lines around appended content.
    """
    blank_id = id_manager.next_id()
    blank_p = etree.Element(qn("w", "p"))

    if base_ppr is not None:
        blank_ppr = copy.deepcopy(base_ppr)
        for old_mark in xpath(blank_ppr, "w:rPr/w:ins") + xpath(blank_ppr, "w:rPr/w:del"):
            old_mark.getparent().remove(old_mark)
    else:
        blank_ppr = etree.Element(qn("w", "pPr"))

    ppr_rpr = xpath(blank_ppr, "w:rPr")
    rpr = ppr_rpr[0] if ppr_rpr else etree.SubElement(blank_ppr, qn("w", "rPr"))

    ins_mark = etree.SubElement(rpr, qn("w", "ins"))
    ins_mark.set(qn("w", "id"), str(blank_id))
    ins_mark.set(qn("w", "author"), author)
    ins_mark.set(qn("w", "date"), date)

    blank_p.append(blank_ppr)
    return blank_p


# ---------------------------------------------------------------------------
# Pseudo-Markdown parser
# ---------------------------------------------------------------------------

# Regex patterns for pseudo-Markdown.
# Order matters: we try the most specific patterns first.
# Hyperlink: [text](url) or [text]
# Bold+Underline+Italic: **___text___**
# Bold+Underline: **__text__**
# Bold+Italic: **_text_**
# Bold: **text**
# Underline+Italic: ___text___  (ambiguous, we treat as underline+italic)
# Underline: __text__
# Italic: _text_

# We use a single regex that matches any formatting sequence and captures
# the formatting markers and the inner text.

_FORMATTING_RE = re.compile(
    r"(\[([^\]]+)\]\(([^)]+)\))"  # hyperlink with URL: [text](url)
    r"|(\[([^\]]+)\])"  # hyperlink without URL: [text]
    r"|(\*\*__(_.*?)_\*\*\*)"  # bold+underline+italic: **___text___**
    r"|(\*\*__(.*?)__\*\*)"  # bold+underline: **__text__**
    r"|(\*\*_(.*?)_\*\*)"  # bold+italic: **_text_**
    r"|(\*\*(.*?)\*\*)"  # bold: **text**
    r"|(__(_(.*?)_)__)"  # underline+italic: ___text___
    r"|(__(.*?)__)"  # underline: __text__
    r"|(_(.*?)_)"  # italic: _text_
)


def _parse_pseudo_markdown(
    text: str,
) -> list[tuple[str, etree._Element | None, bool, bool, bool, str | None]]:
    """Parse pseudo-Markdown into run specifications.

    Returns a list of ``(plain_text, rPr_element_or_None, bold, italic,
    underline, hyperlink_url_or_None)`` tuples.  The ``rPr`` is built with
    *only* the Markdown formatting (not inherited formatting); the
    ``bold``/``italic``/``underline`` flags indicate which markers were
    present so the caller can merge with inherited formatting.

    The text has Markdown escapes removed (``\\*`` → ``*``, ``\\_`` → ``_``).
    """
    if not text:
        return []

    result: list[tuple[str, etree._Element | None, bool, bool, bool, str | None]] = []
    pos = 0

    while pos < len(text):
        m = _FORMATTING_RE.search(text, pos)
        if m is None:
            # Rest is plain text
            remaining = text[pos:]
            if remaining:
                result.append((_unescape(remaining), None, False, False, False, None))
            break

        # Plain text before the match
        if m.start() > pos:
            plain = text[pos : m.start()]
            if plain:
                result.append((_unescape(plain), None, False, False, False, None))

        # Determine which group matched
        if m.group(1):  # hyperlink with URL: [text](url)
            inner = m.group(2)
            url = m.group(3)
            result.append((_unescape(inner), None, False, False, False, url))
        elif m.group(4):  # hyperlink without URL: [text]
            inner = m.group(5)
            result.append((_unescape(inner), None, False, False, False, ""))
        elif m.group(6):  # bold+underline+italic
            inner = m.group(7)
            inner = inner[1:-1] if inner.startswith("_") and inner.endswith("_") else inner
            rpr = _make_rpr(bold=True, italic=True, underline=True)
            result.append((_unescape(inner), rpr, True, True, True, None))
        elif m.group(8):  # bold+underline
            inner = m.group(9)
            rpr = _make_rpr(bold=True, underline=True)
            result.append((_unescape(inner), rpr, True, False, True, None))
        elif m.group(10):  # bold+italic
            inner = m.group(11)
            rpr = _make_rpr(bold=True, italic=True)
            result.append((_unescape(inner), rpr, True, True, False, None))
        elif m.group(12):  # bold
            inner = m.group(13)
            rpr = _make_rpr(bold=True)
            result.append((_unescape(inner), rpr, True, False, False, None))
        elif m.group(14):  # underline+italic
            inner = m.group(16)
            rpr = _make_rpr(italic=True, underline=True)
            result.append((_unescape(inner), rpr, False, True, True, None))
        elif m.group(17):  # underline
            inner = m.group(18)
            rpr = _make_rpr(underline=True)
            result.append((_unescape(inner), rpr, False, False, True, None))
        elif m.group(19):  # italic
            inner = m.group(20)
            rpr = _make_rpr(italic=True)
            result.append((_unescape(inner), rpr, False, True, False, None))

        pos = m.end()

    return result


def _unescape(text: str) -> str:
    """Remove Markdown escapes: ``\\*`` → ``*``, ``\\_`` → ``_``, ``\\\\`` → ``\\``."""
    text = text.replace("\\_", "_")
    text = text.replace("\\*", "*")
    text = text.replace("\\\\", "\\")
    return text


def _make_rpr(
    *,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
) -> etree._Element | None:
    """Build a ``<w:rPr>`` element with the specified formatting."""
    if not (bold or italic or underline):
        return None

    rpr = etree.Element(qn("w", "rPr"))
    if bold:
        etree.SubElement(rpr, qn("w", "b"))
    if italic:
        etree.SubElement(rpr, qn("w", "i"))
    if underline:
        u = etree.SubElement(rpr, qn("w", "u"))
        u.set(qn("w", "val"), "single")
    return rpr
