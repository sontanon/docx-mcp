"""Handler for ``ChangeType.APPEND_AFTER`` — insert a new paragraph.

OOXML tracked insertion of a new paragraph requires:

1. A new ``<w:p>`` element inserted after the reference paragraph.
2. The paragraph properties must contain ``<w:pPr><w:rPr><w:ins .../>``
   to mark the *paragraph mark* as inserted.
3. All content runs must be wrapped in ``<w:ins>`` elements.
4. The new paragraph copies ``<w:pPr>`` (style, numbering, etc.) from
   the reference paragraph to maintain document consistency.

The new text is provided as pseudo-Markdown and must be parsed into
formatted runs.
"""

from __future__ import annotations

import copy
import re

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
) -> tuple[etree._Element, int]:
    """Insert a new tracked-change paragraph after *reference_paragraph*.

    Args:
        reference_paragraph: The ``<w:p>`` element after which to insert.
        new_text: The new paragraph content in pseudo-Markdown format.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.

    Returns:
        A tuple of ``(new_paragraph, ins_id)`` — the new element and its
        annotation ID.
    """
    ins_id = id_manager.next_id()
    author = config.author
    date = config.date_iso()

    # --- 1. Create new <w:p> ---
    new_p = etree.Element(qn("w", "p"))

    # --- 2. Copy <w:pPr> from reference (style, numbering, etc.) ---
    ref_ppr = xpath(reference_paragraph, "w:pPr")
    if ref_ppr:
        new_ppr = copy.deepcopy(ref_ppr[0])
        # Remove any existing tracked-change marks from the copied pPr
        for old_mark in xpath(new_ppr, "w:rPr/w:ins") + xpath(new_ppr, "w:rPr/w:del"):
            old_mark.getparent().remove(old_mark)
    else:
        new_ppr = etree.Element(qn("w", "pPr"))

    # --- 3. Mark paragraph mark as inserted ---
    ppr_rpr = xpath(new_ppr, "w:rPr")
    rpr = ppr_rpr[0] if ppr_rpr else etree.SubElement(new_ppr, qn("w", "rPr"))

    ins_mark = etree.SubElement(rpr, qn("w", "ins"))
    ins_mark.set(qn("w", "id"), str(ins_id))
    ins_mark.set(qn("w", "author"), author)
    ins_mark.set(qn("w", "date"), date)

    new_p.append(new_ppr)

    # --- 4. Parse pseudo-Markdown and create runs inside <w:ins> ---
    ins_wrapper = etree.SubElement(new_p, qn("w", "ins"))
    ins_wrapper.set(qn("w", "id"), str(id_manager.next_id()))
    ins_wrapper.set(qn("w", "author"), author)
    ins_wrapper.set(qn("w", "date"), date)

    run_specs = _parse_pseudo_markdown(new_text)
    for text, rpr_el in run_specs:
        r = build_run_element(text, rpr=rpr_el)
        ins_wrapper.append(r)

    # --- 5. Insert after reference ---
    reference_paragraph.addnext(new_p)

    return new_p, ins_id


# ---------------------------------------------------------------------------
# Pseudo-Markdown parser
# ---------------------------------------------------------------------------

# Regex patterns for pseudo-Markdown.
# Order matters: we try the most specific patterns first.
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
    r"(\*\*__(_.*?_)__\*\*)"  # bold+underline+italic: **___text___**
    r"|(\*\*__(.*?)__\*\*)"  # bold+underline: **__text__**
    r"|(\*\*_(.*?)_\*\*)"  # bold+italic: **_text_**
    r"|(\*\*(.*?)\*\*)"  # bold: **text**
    r"|(__(_(.*?)_)__)"  # underline+italic: ___text___
    r"|(__(.*?)__)"  # underline: __text__
    r"|(_(.*?)_)"  # italic: _text_
)


def _parse_pseudo_markdown(text: str) -> list[tuple[str, etree._Element | None]]:
    """Parse pseudo-Markdown into ``(text, rPr)`` pairs.

    Returns a list of ``(plain_text, rPr_element_or_None)`` tuples.
    The text has Markdown escapes removed (``\\*`` → ``*``, ``\\_`` → ``_``).
    """
    if not text:
        return []

    result: list[tuple[str, etree._Element | None]] = []
    pos = 0

    while pos < len(text):
        m = _FORMATTING_RE.search(text, pos)
        if m is None:
            # Rest is plain text
            remaining = text[pos:]
            if remaining:
                result.append((_unescape(remaining), None))
            break

        # Plain text before the match
        if m.start() > pos:
            plain = text[pos : m.start()]
            if plain:
                result.append((_unescape(plain), None))

        # Determine which group matched
        if m.group(1):  # bold+underline+italic
            inner = m.group(2)
            # Strip the inner italic markers
            inner = inner[1:-1] if inner.startswith("_") and inner.endswith("_") else inner
            rpr = _make_rpr(bold=True, italic=True, underline=True)
            result.append((_unescape(inner), rpr))
        elif m.group(3):  # bold+underline
            inner = m.group(4)
            rpr = _make_rpr(bold=True, underline=True)
            result.append((_unescape(inner), rpr))
        elif m.group(5):  # bold+italic
            inner = m.group(6)
            rpr = _make_rpr(bold=True, italic=True)
            result.append((_unescape(inner), rpr))
        elif m.group(7):  # bold
            inner = m.group(8)
            rpr = _make_rpr(bold=True)
            result.append((_unescape(inner), rpr))
        elif m.group(9):  # underline+italic
            inner = m.group(11)
            rpr = _make_rpr(italic=True, underline=True)
            result.append((_unescape(inner), rpr))
        elif m.group(12):  # underline
            inner = m.group(13)
            rpr = _make_rpr(underline=True)
            result.append((_unescape(inner), rpr))
        elif m.group(14):  # italic
            inner = m.group(15)
            rpr = _make_rpr(italic=True)
            result.append((_unescape(inner), rpr))

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
