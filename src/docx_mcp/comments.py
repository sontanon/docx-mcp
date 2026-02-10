"""Comment management for tracked-change justifications.

Each change in the redlined document gets an associated comment explaining
*why* the change was made.  This module handles:

* Creating the ``<w:comment>`` element in ``word/comments.xml``.
* Inserting ``commentRangeStart`` / ``commentRangeEnd`` / ``commentReference``
  markers in ``word/document.xml`` around the changed content.
* Ensuring the ``comments.xml`` part exists in the ZIP package (with correct
  relationship and content type entries).

Comment placement strategy:

* **Modify**: Comment range spans the changed runs (from first
  ``<w:del>``/``<w:ins>`` to the last in the paragraph).
* **Delete**: Comment range spans the entire paragraph.
* **Append**: Comment range spans the entire new paragraph.
"""

from __future__ import annotations

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig
from docx_mcp.namespaces import (
    CT,
    CT_COMMENTS,
    REL_COMMENTS,
    RELS,
    W,
    qn,
)


def ensure_comments_part(doc: DocxDocument) -> etree._Element:
    """Ensure ``word/comments.xml`` exists and return its root element.

    If the part doesn't exist yet, creates:
    1. An empty ``<w:comments>`` tree.
    2. A relationship entry in ``word/_rels/document.xml.rels``.
    3. A content-type override in ``[Content_Types].xml``.

    Returns:
        The ``<w:comments>`` root element.
    """
    if doc.comments_tree is not None:
        return doc.comments_tree

    # Create <w:comments> root with proper namespace declarations
    comments_root = etree.Element(qn("w", "comments"), nsmap={"w": W})
    doc.comments_tree = comments_root

    # Add relationship
    if doc.rels_tree is not None:
        _add_comment_relationship(doc.rels_tree)

    # Add content type
    if doc.content_types_tree is not None:
        _add_comment_content_type(doc.content_types_tree)

    return comments_root


def add_comment(
    doc: DocxDocument,
    paragraph: etree._Element,
    justification: str,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
    range_elements: list[etree._Element] | None = None,
) -> int:
    """Add a comment to the document associated with *paragraph*.

    Args:
        doc: The document to modify.
        paragraph: The ``<w:p>`` element to attach the comment to.
        justification: The comment text (explanation of the change).
        id_manager: ID allocator.
        config: Author / date configuration.
        range_elements: If provided, the comment range will span from
            the first to the last element in this list (they must be
            children of *paragraph*).  If ``None``, the comment spans
            the entire paragraph.

    Returns:
        The comment ID.
    """
    comment_id = id_manager.next_id()
    author = config.author
    date = config.date_iso()

    # --- 1. Create <w:comment> in comments.xml ---
    comments_root = ensure_comments_part(doc)
    comment_el = etree.SubElement(comments_root, qn("w", "comment"))
    comment_el.set(qn("w", "id"), str(comment_id))
    comment_el.set(qn("w", "author"), author)
    comment_el.set(qn("w", "date"), date)
    comment_el.set(qn("w", "initials"), _initials(author))

    # Add a paragraph with the justification text
    cp = etree.SubElement(comment_el, qn("w", "p"))
    cr = etree.SubElement(cp, qn("w", "r"))
    ct = etree.SubElement(cr, qn("w", "t"))
    ct.text = justification

    # --- 2. Insert markers in the paragraph ---
    _insert_comment_markers(paragraph, comment_id, range_elements=range_elements)

    return comment_id


def _insert_comment_markers(
    paragraph: etree._Element,
    comment_id: int,
    *,
    range_elements: list[etree._Element] | None = None,
) -> None:
    """Insert ``commentRangeStart``, ``commentRangeEnd``, and ``commentReference``.

    If *range_elements* is provided, the range spans from before the first
    element to after the last.  Otherwise it spans the entire paragraph
    (after ``<w:pPr>``, before the last child).
    """
    crs = etree.Element(qn("w", "commentRangeStart"))
    crs.set(qn("w", "id"), str(comment_id))

    cre = etree.Element(qn("w", "commentRangeEnd"))
    cre.set(qn("w", "id"), str(comment_id))

    # Comment reference run
    ref_run = etree.Element(qn("w", "r"))
    ref_rpr = etree.SubElement(ref_run, qn("w", "rPr"))
    etree.SubElement(ref_rpr, qn("w", "rStyle")).set(qn("w", "val"), "CommentReference")
    etree.SubElement(ref_run, qn("w", "commentReference")).set(
        qn("w", "id"),
        str(comment_id),
    )

    if range_elements:
        # Insert commentRangeStart before the first range element
        first = range_elements[0]
        first.addprevious(crs)

        # Insert commentRangeEnd after the last range element
        last = range_elements[-1]
        last.addnext(ref_run)
        last.addnext(cre)
    else:
        # Span the entire paragraph: after pPr, before end
        children = list(paragraph)
        if children:
            # Find first non-pPr child
            insert_before = None
            for child in children:
                tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""
                if tag_local != "pPr":
                    insert_before = child
                    break

            if insert_before is not None:
                insert_before.addprevious(crs)
            else:
                paragraph.append(crs)

            # Add range end and reference at the end
            paragraph.append(cre)
            paragraph.append(ref_run)
        else:
            paragraph.append(crs)
            paragraph.append(cre)
            paragraph.append(ref_run)


def _initials(author: str) -> str:
    """Extract initials from an author name.

    "AI Review" → "AR"
    "John Smith" → "JS"
    "Single" → "S"
    """
    parts = author.split()
    return "".join(p[0].upper() for p in parts if p)


def _add_comment_relationship(rels_tree: etree._Element) -> None:
    """Add a relationship entry for ``word/comments.xml`` if not present."""
    # Check if relationship already exists
    for rel in rels_tree:
        if isinstance(rel.tag, str) and rel.get("Type") == REL_COMMENTS:
            return  # Already exists

    # Generate a unique rId
    existing_ids = {el.get("Id", "") for el in rels_tree if isinstance(el.tag, str)}
    rid_num = 1
    while f"rId{rid_num}" in existing_ids:
        rid_num += 1

    rel = etree.SubElement(rels_tree, f"{{{RELS}}}Relationship")
    rel.set("Id", f"rId{rid_num}")
    rel.set("Type", REL_COMMENTS)
    rel.set("Target", "comments.xml")


def _add_comment_content_type(ct_tree: etree._Element) -> None:
    """Add a content-type override for ``/word/comments.xml`` if not present."""
    for child in ct_tree:
        if isinstance(child.tag, str):
            part_name = child.get("PartName", "")
            if part_name == "/word/comments.xml":
                return  # Already exists

    override = etree.SubElement(ct_tree, f"{{{CT}}}Override")
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", CT_COMMENTS)
