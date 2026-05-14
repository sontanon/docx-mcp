"""Validation utilities for redlined .docx documents.

Provides structural checks to catch errors before the file reaches Word:

* **Annotation ID isolation** — tracked-change IDs and comment IDs
  must not collide with each other (within each group, ID sharing
  is normal OOXML behaviour).
* **Comment integrity** — every ``<w:comment>`` must have matching
  ``commentRangeStart``, ``commentRangeEnd``, and ``commentReference``
  in the document body, and vice-versa.
* **Tracked-change structure** — ``<w:ins>`` and ``<w:del>`` elements
  must carry required ``w:id``, ``w:author``, and ``w:date`` attributes.
* **Content-type / relationship consistency** — if ``word/comments.xml``
  exists, the corresponding content-type override and relationship entry
  must also be present.
"""

import contextlib
from dataclasses import dataclass, field

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.namespaces import CT_COMMENTS, REL_COMMENTS, qn


@dataclass
class ValidationResult:
    """Outcome of a document validation pass.

    Attributes:
        errors: Fatal issues that will likely cause Word to trigger a
            repair prompt.
        warnings: Non-fatal issues that may cause unexpected visual
            behaviour but won't corrupt the file.
    """

    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        """True if no errors were found (warnings are acceptable)."""
        return len(self.errors) == 0


def validate_document(doc: DocxDocument) -> ValidationResult:
    """Run all validation checks on a redlined document.

    Args:
        doc: The document to validate (typically the output of
            :func:`~docx_mcp.redliner.apply_redlines`).

    Returns:
        A :class:`ValidationResult` with any errors and warnings.
    """
    result = ValidationResult()
    _check_annotation_id_uniqueness(doc, result)
    _check_comment_integrity(doc, result)
    _check_tracked_change_attributes(doc, result)
    _check_package_consistency(doc, result)
    return result


# ---------------------------------------------------------------------------
# Individual checks
# ---------------------------------------------------------------------------


def _check_annotation_id_uniqueness(
    doc: DocxDocument,
    result: ValidationResult,
) -> None:
    """Check that annotation IDs are used correctly.

    In OOXML, some elements intentionally share the same ``w:id``:

    * **Comment groups**: ``commentRangeStart``, ``commentRangeEnd``,
      ``commentReference`` (in document.xml) and ``comment`` (in
      comments.xml) all share one ID to link together.
    * **Tracked-change groups**: a ``<w:del>`` or ``<w:ins>`` wrapper
      around runs may share its ID with the paragraph-mark deletion/
      insertion marker in ``<w:pPr><w:rPr><w:del>`` or ``<w:ins>``.

    This check verifies that IDs are not reused across *unrelated*
    annotation groups (e.g., a tracked change ID should not collide
    with a comment ID).
    """
    # Classify every w:id by its group type.
    # Elements in the same group type can share an ID.
    _COMMENT_TAGS = {"commentRangeStart", "commentRangeEnd", "commentReference", "comment"}
    _TC_TAGS = {"ins", "del"}

    # Map: (group, id) -> list of locations
    groups: dict[str, dict[int, list[str]]] = {
        "comment": {},
        "tracked_change": {},
        "other": {},
    }

    def _classify(tag_local: str) -> str:
        if tag_local in _COMMENT_TAGS:
            return "comment"
        if tag_local in _TC_TAGS:
            return "tracked_change"
        return "other"

    def _scan(tree: etree._Element, source: str) -> None:
        for el in tree.iter():
            raw = el.get(qn("w", "id"))
            if raw is None:
                continue
            with contextlib.suppress(ValueError):
                wid = int(raw)
                tag = etree.QName(el.tag).localname if isinstance(el.tag, str) else ""
                group = _classify(tag)
                groups[group].setdefault(wid, []).append(f"{source}/<{tag}>")

    _scan(doc.document_tree, "document.xml")
    if doc.comments_tree is not None:
        _scan(doc.comments_tree, "comments.xml")

    # Check for cross-group collisions
    all_groups = list(groups.items())
    for i, (group_a, ids_a) in enumerate(all_groups):
        for group_b, ids_b in all_groups[i + 1 :]:
            shared = set(ids_a) & set(ids_b)
            for wid in sorted(shared):
                locs_a = ", ".join(ids_a[wid])
                locs_b = ", ".join(ids_b[wid])
                result.errors.append(
                    f"w:id={wid} used in both {group_a} ({locs_a}) and {group_b} ({locs_b})"
                )


def _check_comment_integrity(
    doc: DocxDocument,
    result: ValidationResult,
) -> None:
    """Verify that comment IDs are consistent across all parts.

    Every ``<w:comment w:id="N">`` in ``comments.xml`` must have a
    corresponding ``commentRangeStart``, ``commentRangeEnd``, and
    ``commentReference`` in ``document.xml`` — and vice-versa.
    """
    # Collect comment IDs from comments.xml
    comment_ids: set[int] = set()
    if doc.comments_tree is not None:
        for comment_el in doc.comments_tree.iter(qn("w", "comment")):
            raw = comment_el.get(qn("w", "id"))
            if raw is not None:
                with contextlib.suppress(ValueError):
                    comment_ids.add(int(raw))

    # Collect IDs from document.xml markers
    range_start_ids: set[int] = set()
    range_end_ids: set[int] = set()
    reference_ids: set[int] = set()

    for el in doc.document_tree.iter():
        tag = etree.QName(el.tag).localname if isinstance(el.tag, str) else ""
        raw = el.get(qn("w", "id"))
        if raw is None:
            continue
        with contextlib.suppress(ValueError):
            wid = int(raw)
            if tag == "commentRangeStart":
                range_start_ids.add(wid)
            elif tag == "commentRangeEnd":
                range_end_ids.add(wid)
            elif tag == "commentReference":
                reference_ids.add(wid)

    # Check: every comment in comments.xml should have all three markers
    for cid in sorted(comment_ids):
        if cid not in range_start_ids:
            result.errors.append(
                f"Comment id={cid} in comments.xml has no commentRangeStart in document.xml"
            )
        if cid not in range_end_ids:
            result.errors.append(
                f"Comment id={cid} in comments.xml has no commentRangeEnd in document.xml"
            )
        if cid not in reference_ids:
            result.warnings.append(
                f"Comment id={cid} in comments.xml has no "
                f"commentReference in document.xml (comment won't be "
                f"visually anchored)"
            )

    # Check: every marker in document.xml should reference a real comment
    all_marker_ids = range_start_ids | range_end_ids | reference_ids
    for mid in sorted(all_marker_ids - comment_ids):
        # Only flag if comments.xml actually exists (otherwise the markers
        # are just orphaned from a previous edit — common in real .docx)
        if doc.comments_tree is not None:
            result.warnings.append(
                f"Comment marker id={mid} in document.xml has no "
                f"matching <w:comment> in comments.xml"
            )

    # Check: rangeStart and rangeEnd should be paired
    for sid in sorted(range_start_ids - range_end_ids):
        result.errors.append(f"commentRangeStart id={sid} has no matching commentRangeEnd")
    for eid in sorted(range_end_ids - range_start_ids):
        result.errors.append(f"commentRangeEnd id={eid} has no matching commentRangeStart")


def _check_tracked_change_attributes(
    doc: DocxDocument,
    result: ValidationResult,
) -> None:
    """Verify that ``<w:ins>`` and ``<w:del>`` elements have required attributes.

    Word expects ``w:id``, ``w:author``, and ``w:date`` on every tracked
    change wrapper element.
    """
    tc_tags = {qn("w", "ins"), qn("w", "del")}

    for el in doc.document_tree.iter():
        if el.tag not in tc_tags:
            continue

        tag_local = etree.QName(el.tag).localname
        wid = el.get(qn("w", "id"))
        author = el.get(qn("w", "author"))
        date = el.get(qn("w", "date"))

        # Some <w:ins>/<w:del> inside <w:rPr> mark paragraph-mark
        # changes — these always need the attributes.
        missing = []
        if wid is None:
            missing.append("w:id")
        if author is None:
            missing.append("w:author")
        if date is None:
            missing.append("w:date")

        if missing:
            result.errors.append(
                f"<w:{tag_local}> element missing required attributes: {', '.join(missing)}"
            )


def _check_package_consistency(
    doc: DocxDocument,
    result: ValidationResult,
) -> None:
    """If comments.xml exists, ensure content-type and relationship entries exist."""
    if doc.comments_tree is None:
        return

    # Check content type
    if doc.content_types_tree is not None:
        found_ct = False
        for child in doc.content_types_tree:
            if (
                isinstance(child.tag, str)
                and child.get("PartName") == "/word/comments.xml"
                and child.get("ContentType") == CT_COMMENTS
            ):
                found_ct = True
                break
        if not found_ct:
            result.errors.append("[Content_Types].xml missing Override for /word/comments.xml")
    else:
        result.errors.append("Document has comments.xml but no [Content_Types].xml")

    # Check relationship
    if doc.rels_tree is not None:
        found_rel = False
        for child in doc.rels_tree:
            if isinstance(child.tag, str) and child.get("Type") == REL_COMMENTS:
                found_rel = True
                break
        if not found_rel:
            result.errors.append(
                "word/_rels/document.xml.rels missing relationship for comments.xml"
            )
    else:
        result.warnings.append(
            "Document has comments.xml but no document.xml.rels "
            "(relationship entry cannot be verified)"
        )
