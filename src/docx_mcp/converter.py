"""Convert .docx paragraphs to pseudo-Markdown fragment representation.

This module converts the runs within each paragraph to a pseudo-Markdown format
that preserves bold (**), italic (_), and underline (__) formatting while being
suitable for LLM consumption.

Unicode characters (smart quotes, em-dashes, section symbols, etc.) are preserved
as-is. Only Markdown syntax characters (*, _) are escaped with backslashes.
"""

from __future__ import annotations

import re

from lxml import etree

from docx_mcp.namespaces import qn, xpath


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


def _extract_run_text(run: etree._Element) -> str:
    """Extract text content from a <w:r> element.

    Concatenates all <w:t> children. Handles <w:br/> as newline
    and <w:tab/> as tab (though tabs are rare in legal docs).
    """
    parts: list[str] = []
    for child in run:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else None
        if tag == "t" and child.text:
            parts.append(child.text)
        elif tag == "br":
            parts.append("\n")
        elif tag == "tab":
            parts.append("\t")
    return "".join(parts)


def paragraph_to_pseudo_markdown(paragraph: etree._Element) -> str:
    """Convert a single <w:p> element to pseudo-Markdown text.

    Walks the runs, extracts text and formatting, and produces a
    pseudo-Markdown string with **, _, __ markers.

    Unicode characters (smart quotes, em-dashes, etc.) are preserved as-is.
    Markdown syntax characters (* and _) in the source text are escaped.
    Multiple consecutive spaces are collapsed to single spaces.
    """
    segments: list[str] = []

    for run in xpath(paragraph, "w:r"):
        text = _extract_run_text(run)
        if not text:
            continue

        # Get formatting
        rpr_list = xpath(run, "w:rPr")
        rpr = rpr_list[0] if rpr_list else None

        bold = _is_bold(rpr)
        italic = _is_italic(rpr)
        underline = _is_underline(rpr)

        # Escape markdown syntax characters in the raw text
        escaped = _escape_markdown(text)

        # Wrap with formatting markers
        formatted = _wrap_formatting(escaped, bold=bold, italic=italic, underline=underline)
        segments.append(formatted)

    result = "".join(segments)

    # Collapse multiple consecutive spaces to single space
    # (but preserve non-breaking spaces and other unicode whitespace as-is)
    result = re.sub(r"  +", " ", result)

    return result


def document_to_fragments(
    paragraphs: list[etree._Element],
) -> list[tuple[int, str]]:
    """Convert a list of paragraph elements to (fragment_id, pseudo_markdown) pairs.

    Args:
        paragraphs: List of <w:p> elements (typically from DocxDocument.paragraphs).

    Returns:
        List of (fragment_id, pseudo_markdown) tuples, 1-indexed.
        Empty paragraphs are included with empty string text.
    """
    fragments: list[tuple[int, str]] = []
    for i, para in enumerate(paragraphs, start=1):
        md = paragraph_to_pseudo_markdown(para)
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
