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
        self._header_trees: dict[str, etree._Element] = {}
        self._footer_trees: dict[str, etree._Element] = {}
        self._header_paths: dict[str, str] = {}  # rel_id -> zip entry path
        self._footer_paths: dict[str, str] = {}  # rel_id -> zip entry path

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
            self._load_header_footer_trees()

    def _load_header_footer_trees(self) -> None:
        """Discover and parse header/footer XML parts from relationships."""
        if self._rels_tree is None:
            return

        ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        for rel in self._rels_tree:
            rel_type = rel.get("Type", "")
            target = rel.get("Target", "")
            rel_id = rel.get("Id", "")
            if not target:
                continue

            # Resolve target path relative to word/
            target_path = target if not target.startswith("/") else target[1:]
            if not target_path.startswith("word/"):
                target_path = f"word/{target_path}"

            if rel_type == f"{ns}/header":
                xml_bytes = self._zip_entries.get(target_path)
                if xml_bytes is not None:
                    self._header_trees[rel_id] = etree.fromstring(xml_bytes)
                    self._header_paths[rel_id] = target_path
            elif rel_type == f"{ns}/footer":
                xml_bytes = self._zip_entries.get(target_path)
                if xml_bytes is not None:
                    self._footer_trees[rel_id] = etree.fromstring(xml_bytes)
                    self._footer_paths[rel_id] = target_path

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

    def full_element_map(self) -> dict[str, etree._Element]:
        """Build a unified element map for body, headers, and footers.

        Body elements use plain numeric keys (``"1"``, ``"2"``, …).
        Header paragraphs use ``"header_{part_index}.{element_index}"``.
        Footer paragraphs use ``"footer_{part_index}.{element_index}"``.

        Only ``<w:p>`` elements are included for headers/footers.
        Tables inside headers/footers are omitted and should be reported
        in ``skipped_elements`` by the caller.

        Returns:
            Dict mapping fragment_id (str) to the lxml element.
        """
        result: dict[str, etree._Element] = {}

        # Body elements (paragraphs and tables)
        for i, el in enumerate(self.body_elements, start=1):
            result[str(i)] = el

        # Header paragraphs
        for part_idx, (_rel_id, tree) in enumerate(self._header_trees.items(), start=1):
            para_idx = 0
            for child in tree:
                tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if tag_local == "p":
                    para_idx += 1
                    result[f"header_{part_idx}.{para_idx}"] = child
                elif tag_local == "tbl":
                    # Tables in headers are not editable; skip
                    pass

        # Footer paragraphs
        for part_idx, (_rel_id, tree) in enumerate(self._footer_trees.items(), start=1):
            para_idx = 0
            for child in tree:
                tag_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                if tag_local == "p":
                    para_idx += 1
                    result[f"footer_{part_idx}.{para_idx}"] = child
                elif tag_local == "tbl":
                    pass

        return result

    def resolve_fragment_id(self, fragment_id: str) -> tuple[etree._Element, str]:
        """Resolve a fragment ID to its element and parent tree type.

        Args:
            fragment_id: Fragment ID (e.g. ``"5"``, ``"header_1.3"``).

        Returns:
            Tuple of (element, tree_type) where tree_type is one of
            ``"body"``, ``"header"``, ``"footer"``.

        Raises:
            ValueError: If the fragment_id is unknown or malformed.
        """
        element_map = self.full_element_map()
        element = element_map.get(fragment_id)
        if element is not None:
            if fragment_id.startswith("header_"):
                return element, "header"
            if fragment_id.startswith("footer_"):
                return element, "footer"
            return element, "body"

        msg = f"Unknown fragment_id: {fragment_id!r}"
        raise ValueError(msg)

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

    @property
    def header_trees(self) -> dict[str, etree._Element]:
        """Dict of relationship ID → header XML tree."""
        return self._header_trees

    @property
    def footer_trees(self) -> dict[str, etree._Element]:
        """Dict of relationship ID → footer XML tree."""
        return self._footer_trees

    def has_tracked_changes(self) -> list[str]:
        """Check for pre-existing tracked changes in any document part.

        Scans document.xml, all header parts, all footer parts, and
        comments.xml for ``<w:ins>``, ``<w:del>``, ``<w:moveFrom>``,
        and ``<w:moveTo>`` elements.

        Returns:
            List of part names that contain tracked changes.
            Empty list if the document is clean.
        """
        tracked_tags = {"ins", "del", "moveFrom", "moveTo"}
        dirty_parts: list[str] = []

        def _part_has_changes(tree: etree._Element, part_name: str) -> bool:
            for el in tree.iter():
                tag_local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if tag_local in tracked_tags:
                    return True
            return False

        if _part_has_changes(self.document_tree, "word/document.xml"):
            dirty_parts.append("word/document.xml")

        for rel_id, tree in self._header_trees.items():
            if _part_has_changes(tree, f"header ({rel_id})"):
                dirty_parts.append(f"word/header ({rel_id})")

        for rel_id, tree in self._footer_trees.items():
            if _part_has_changes(tree, f"footer ({rel_id})"):
                dirty_parts.append(f"word/footer ({rel_id})")

        if self._comments_tree is not None and _part_has_changes(
            self._comments_tree, "word/comments.xml"
        ):
            dirty_parts.append("word/comments.xml")

        return dirty_parts

    def max_annotation_id(self) -> int:
        """Find the highest annotation ID used in the document.

        Scans all w:id attributes across document.xml, comments.xml,
        and all header/footer parts to find the maximum. Returns 0 if
        no annotations exist.
        """
        max_id = 0

        trees = [self.document_tree]
        trees.extend(self._header_trees.values())
        trees.extend(self._footer_trees.values())
        if self._comments_tree is not None:
            trees.append(self._comments_tree)

        for tree in trees:
            for el in tree.iter():
                wid = el.get(qn("w", "id"))
                if wid is not None:
                    with contextlib.suppress(ValueError):
                        max_id = max(max_id, int(wid))

        return max_id

    def resolve_hyperlink_url(self, rel_id: str) -> str | None:
        """Resolve a hyperlink relationship ID to its target URL.

        Looks up the ``r:id`` in ``word/_rels/document.xml.rels``.

        Args:
            rel_id: The relationship ID (e.g., ``"rId4"``).

        Returns:
            The target URL, or ``None`` if the relationship is not found
            or is not an external hyperlink.
        """
        if self._rels_tree is None:
            return None

        hyperlink_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        )
        for rel in self._rels_tree:
            if rel.get("Id") == rel_id:
                if rel.get("Type") == hyperlink_type:
                    return rel.get("Target")
                return None
        return None

    def create_hyperlink_relationship(self, url: str) -> str:
        """Create a new hyperlink relationship and return its ``r:id``.

        Appends a new ``<Relationship>`` element to
        ``word/_rels/document.xml.rels`` with a unique ID.

        Args:
            url: The target URL for the hyperlink.

        Returns:
            The newly allocated relationship ID (e.g., ``"rId999"``).
        """
        if self._rels_tree is None:
            msg = "Cannot create hyperlink relationship: document has no rels tree"
            raise RuntimeError(msg)

        hyperlink_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        )

        # Find the highest existing rId number
        max_num = 0
        for rel in self._rels_tree:
            rid = rel.get("Id", "")
            if rid.startswith("rId"):
                with contextlib.suppress(ValueError):
                    max_num = max(max_num, int(rid[3:]))

        new_rid = f"rId{max_num + 1}"
        new_rel = etree.SubElement(self._rels_tree, qn("r", "Relationship"))
        new_rel.set("Id", new_rid)
        new_rel.set("Type", hyperlink_type)
        new_rel.set("Target", url)
        new_rel.set("TargetMode", "External")

        return new_rid

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

        # Re-serialize modified header/footer parts
        for rel_id, tree in self._header_trees.items():
            path = self._header_paths.get(rel_id)
            if path is not None:
                entries[path] = etree.tostring(
                    tree, xml_declaration=True, encoding="UTF-8", standalone=True
                )

        for rel_id, tree in self._footer_trees.items():
            path = self._footer_paths.get(rel_id)
            if path is not None:
                entries[path] = etree.tostring(
                    tree, xml_declaration=True, encoding="UTF-8", standalone=True
                )

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for entry_path, entry_bytes in entries.items():
                zf.writestr(entry_path, entry_bytes)

        return buf.getvalue()

    def save(self, path: Path | str) -> None:
        """Save the document to a .docx file."""
        Path(path).write_bytes(self.to_bytes())
