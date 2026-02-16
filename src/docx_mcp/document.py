"""Document parsing and manipulation for .docx files.

Provides the DocxDocument class which handles:
- Unzipping and parsing .docx files
- Indexing paragraphs with fragment IDs
- In-memory XML tree manipulation
- Repacking to .docx
"""

from __future__ import annotations

import contextlib
import io
import zipfile
from pathlib import Path

from lxml import etree

from docx_mcp.namespaces import qn, xpath


class DocxDocument:
    """A parsed .docx file with in-memory XML trees.

    The document is represented as:
    - A dict of ZIP entry paths → bytes (for non-XML parts like images)
    - Parsed lxml trees for key XML parts (document.xml, comments.xml, etc.)

    Modifications are made in-place on the lxml trees. Call save() or to_bytes()
    to produce the final .docx output.
    """

    def __init__(
        self,
        path: Path | str | None = None,
        data: bytes | None = None,
    ) -> None:
        """Load a .docx file from a path or raw bytes.

        Exactly one of path or data must be provided.
        """
        if path is not None and data is not None:
            msg = "Provide either path or data, not both"
            raise ValueError(msg)
        if path is None and data is None:
            msg = "Provide either path or data"
            raise ValueError(msg)

        raw = Path(path).read_bytes() if path is not None else data
        assert raw is not None

        self._zip_entries: dict[str, bytes] = {}
        self._document_tree: etree._Element | None = None
        self._comments_tree: etree._Element | None = None
        self._content_types_tree: etree._Element | None = None
        self._rels_tree: etree._Element | None = None  # word/_rels/document.xml.rels

        self._parse_zip(raw)

    def _parse_zip(self, raw: bytes) -> None:
        """Parse the ZIP archive and extract key XML parts."""
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            for entry in zf.namelist():
                self._zip_entries[entry] = zf.read(entry)

        # Parse document.xml (required)
        doc_xml = self._zip_entries.get("word/document.xml")
        if doc_xml is None:
            msg = "Invalid .docx file: missing word/document.xml"
            raise ValueError(msg)
        self._document_tree = etree.fromstring(doc_xml)

        # Parse comments.xml (optional — may not exist yet)
        comments_xml = self._zip_entries.get("word/comments.xml")
        if comments_xml is not None:
            self._comments_tree = etree.fromstring(comments_xml)

        # Parse [Content_Types].xml (required)
        ct_xml = self._zip_entries.get("[Content_Types].xml")
        if ct_xml is not None:
            self._content_types_tree = etree.fromstring(ct_xml)

        # Parse word/_rels/document.xml.rels (required)
        rels_xml = self._zip_entries.get("word/_rels/document.xml.rels")
        if rels_xml is not None:
            self._rels_tree = etree.fromstring(rels_xml)

    @property
    def document_tree(self) -> etree._Element:
        """The root element of word/document.xml."""
        assert self._document_tree is not None
        return self._document_tree

    @property
    def body(self) -> etree._Element:
        """The <w:body> element containing all document content."""
        bodies = xpath(self.document_tree, ".//w:body")
        if not bodies:
            msg = "Invalid document.xml: missing w:body element"
            raise ValueError(msg)
        return bodies[0]

    @property
    def paragraphs(self) -> list[etree._Element]:
        """All <w:p> elements that are direct children of <w:body>.

        These are the top-level paragraphs, excluding those inside tables,
        text boxes, or other nested containers.
        """
        return xpath(self.body, "./w:p")

    @property
    def body_elements(self) -> list[etree._Element]:
        """All direct children of <w:body> in document order.

        Includes both ``<w:p>`` (paragraphs) and ``<w:tbl>`` (tables).
        Section properties ``<w:sectPr>`` and other non-content elements
        are excluded.
        """
        result: list[etree._Element] = []
        for child in self.body:
            tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag_local in ("p", "tbl"):
                result.append(child)
        return result

    def interleaved_element_map(self) -> dict[int, etree._Element]:
        """Build a 1-based ID → element map for paragraphs and tables.

        Tables and paragraphs share the same ID space in document order.
        Fragment ID 1 corresponds to the first element in body_elements,
        fragment ID 2 to the second, etc.

        Returns:
            Dict mapping element_id (1..N) to the lxml element.
        """
        return {i: el for i, el in enumerate(self.body_elements, start=1)}

    def fragment_map(self) -> dict[int, etree._Element]:
        """Build a mapping of 1-based fragment IDs to paragraph elements.

        Fragment ID 1 corresponds to the first paragraph in the body,
        fragment ID 2 to the second, etc.

        Returns:
            Dict mapping fragment_id (1..N) to the <w:p> lxml element.
        """
        return {i: para for i, para in enumerate(self.paragraphs, start=1)}

    @property
    def comments_tree(self) -> etree._Element | None:
        """The root element of word/comments.xml, or None if it doesn't exist."""
        return self._comments_tree

    @comments_tree.setter
    def comments_tree(self, tree: etree._Element) -> None:
        """Set or replace the comments.xml tree."""
        self._comments_tree = tree

    @property
    def content_types_tree(self) -> etree._Element | None:
        """The root element of [Content_Types].xml."""
        return self._content_types_tree

    @property
    def rels_tree(self) -> etree._Element | None:
        """The root element of word/_rels/document.xml.rels."""
        return self._rels_tree

    def max_annotation_id(self) -> int:
        """Find the highest annotation ID used in the document.

        Scans all w:id attributes across document.xml and comments.xml
        to find the maximum. Returns 0 if no annotations exist.
        """
        max_id = 0

        # Scan document.xml for w:id attributes
        for el in self.document_tree.iter():
            wid = el.get(qn("w", "id"))
            if wid is not None:
                with contextlib.suppress(ValueError):
                    max_id = max(max_id, int(wid))

        # Also scan comments.xml if it exists
        if self._comments_tree is not None:
            for el in self._comments_tree.iter():
                wid = el.get(qn("w", "id"))
                if wid is not None:
                    with contextlib.suppress(ValueError):
                        max_id = max(max_id, int(wid))

        return max_id

    def deepcopy(self) -> DocxDocument:
        """Create a deep copy of this document (for validation comparisons)."""
        return DocxDocument(data=self.to_bytes())

    def to_bytes(self) -> bytes:
        """Serialize the document back to .docx bytes.

        Modified XML parts are re-serialized from their lxml trees.
        Unmodified parts pass through as-is.
        """
        buf = io.BytesIO()

        # Build the set of entries we'll write, updating XML parts
        entries = dict(self._zip_entries)

        # Re-serialize modified XML parts
        if self._document_tree is not None:
            entries["word/document.xml"] = etree.tostring(
                self._document_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        if self._comments_tree is not None:
            entries["word/comments.xml"] = etree.tostring(
                self._comments_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        if self._content_types_tree is not None:
            entries["[Content_Types].xml"] = etree.tostring(
                self._content_types_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        if self._rels_tree is not None:
            entries["word/_rels/document.xml.rels"] = etree.tostring(
                self._rels_tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for entry_path, entry_bytes in entries.items():
                zf.writestr(entry_path, entry_bytes)

        return buf.getvalue()

    def save(self, path: Path | str) -> None:
        """Save the document to a .docx file."""
        Path(path).write_bytes(self.to_bytes())
