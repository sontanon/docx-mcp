"""Handler for ``ChangeType.DELETE`` — mark an entire paragraph as deleted.

OOXML tracked deletion of a paragraph requires:

1. Every ``<w:r>`` in the paragraph is wrapped in a ``<w:del>`` element.
2. Each ``<w:t>`` inside those runs is replaced with ``<w:delText>``.
3. The paragraph properties ``<w:pPr>`` must contain a ``<w:rPr><w:del .../>``
   element to mark the *paragraph mark* itself as deleted (otherwise Word
   shows the paragraph as still existing even though all its text is struck).

The resulting paragraph is still present in the XML — Word just renders it
with strikethrough and hides it when "Show Markup" is off.
"""

from lxml import etree

from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig
from docx_mcp.namespaces import XML, qn, xpath


def handle_delete(
    paragraph: etree._Element,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
    preserve_paragraph_mark: bool = False,
) -> int:
    """Mark *paragraph* as a tracked deletion in-place.

    Args:
        paragraph: The ``<w:p>`` element to mark as deleted.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.
        preserve_paragraph_mark: If True, skip marking the paragraph mark
            as deleted. Used for table cells where ``<w:tc>`` must contain
            at least one ``<w:p>`` element.

    Returns:
        The annotation ID used for the ``<w:del>`` wrapper.
    """
    del_id = id_manager.next_id()
    author = config.author
    date = config.date_iso()

    # --- 1. Wrap all runs in <w:del> ---
    # We handle both direct <w:r> children and <w:r> inside <w:hyperlink>.
    # For hyperlinks, the <w:del> goes inside the <w:hyperlink> wrapper.
    _wrap_runs_in_del(paragraph, del_id=del_id, author=author, date=date)

    # --- 2. Mark the paragraph mark as deleted ---
    if not preserve_paragraph_mark:
        _mark_paragraph_mark_deleted(paragraph, del_id=del_id, author=author, date=date)

    return del_id


def _wrap_runs_in_del(
    paragraph: etree._Element,
    *,
    del_id: int,
    author: str,
    date: str,
) -> None:
    """Wrap all text-bearing runs in *paragraph* with ``<w:del>``.

    Direct ``<w:r>`` children are replaced by a ``<w:del>`` containing
    the run.  Runs inside ``<w:hyperlink>`` are wrapped in a ``<w:del>``
    that is placed *inside* the hyperlink wrapper so the relationship is
    preserved.
    """
    del_attribs = {
        qn("w", "id"): str(del_id),
        qn("w", "author"): author,
        qn("w", "date"): date,
    }

    for child in list(paragraph):
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

        if tag == "r":
            # Direct run — wrap in <w:del>
            _convert_t_to_del_text(child)
            del_el = etree.Element(qn("w", "del"))
            for k, v in del_attribs.items():
                del_el.set(k, v)
            child.addprevious(del_el)
            del_el.append(child)

        elif tag == "hyperlink":
            # Wrap the entire <w:hyperlink> in <w:del> (consistent with modify)
            for run in xpath(child, "w:r"):
                _convert_t_to_del_text(run)
            del_el = etree.Element(qn("w", "del"))
            for k, v in del_attribs.items():
                del_el.set(k, v)
            child.addprevious(del_el)
            del_el.append(child)


def _convert_t_to_del_text(run: etree._Element) -> None:
    """Convert all ``<w:t>`` elements in *run* to ``<w:delText>``.

    Preserves ``xml:space`` and text content.
    """
    for t_el in list(xpath(run, "w:t")):
        # Create <w:delText> with same content
        dt = etree.Element(qn("w", "delText"))
        dt.text = t_el.text
        space = t_el.get(f"{{{XML}}}space")
        if space:
            dt.set(f"{{{XML}}}space", space)
        # Replace <w:t> with <w:delText> at the same position
        t_el.addprevious(dt)
        t_el.getparent().remove(t_el)


def _mark_paragraph_mark_deleted(
    paragraph: etree._Element,
    *,
    del_id: int,
    author: str,
    date: str,
) -> None:
    """Add ``<w:pPr><w:rPr><w:del .../>`` to mark the paragraph mark as deleted.

    This is required for Word to properly show the paragraph as deleted
    (otherwise just the text is deleted but the empty paragraph remains).
    """
    # Get or create <w:pPr>
    ppr_list = xpath(paragraph, "w:pPr")
    if ppr_list:
        ppr = ppr_list[0]
    else:
        ppr = etree.Element(qn("w", "pPr"))
        paragraph.insert(0, ppr)

    # Get or create <w:rPr> inside <w:pPr>
    rpr_list = xpath(ppr, "w:rPr")
    rpr = rpr_list[0] if rpr_list else etree.SubElement(ppr, qn("w", "rPr"))

    # Add <w:del> to mark the paragraph mark
    del_mark = etree.SubElement(rpr, qn("w", "del"))
    del_mark.set(qn("w", "id"), str(del_id))
    del_mark.set(qn("w", "author"), author)
    del_mark.set(qn("w", "date"), date)
