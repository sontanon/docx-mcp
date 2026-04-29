"""Document auditing — structural diagnostics without text extraction.

Reports counts of headers, footers, images, tables, section breaks,
and unsupported elements. Used by the ``audit_document`` CLI and MCP tool.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from docx_mcp.document import DocxDocument
from docx_mcp.namespaces import qn, xpath


@dataclass
class AuditReport:
    """Structural audit results for a .docx file."""

    header_parts: int = 0
    header_paragraphs: int = 0
    footer_parts: int = 0
    footer_paragraphs: int = 0
    images_body: int = 0
    images_header: int = 0
    images_footer: int = 0
    simple_tables: int = 0
    skipped_tables: int = 0
    skipped_table_reasons: list[str] = field(default_factory=list)
    section_breaks: int = 0
    multi_column: bool = False
    has_tracked_changes: bool = False
    has_comments: bool = False
    unsupported_elements: list[str] = field(default_factory=list)

    def to_text(self) -> str:
        """Format as human-readable text."""
        lines: list[str] = ["Document Audit Report", "=" * 40]
        lines.append(
            f"Headers: {self.header_parts} part(s), {self.header_paragraphs} paragraph(s)"
        )
        lines.append(
            f"Footers: {self.footer_parts} part(s), {self.footer_paragraphs} paragraph(s)"
        )
        total_images = self.images_body + self.images_header + self.images_footer
        lines.append(
            f"Images: {total_images} "
            f"({self.images_body} body, {self.images_header} header, {self.images_footer} footer)"
        )
        lines.append(f"Tables: {self.simple_tables} simple, {self.skipped_tables} skipped")
        for reason in self.skipped_table_reasons:
            lines.append(f"  - skipped: {reason}")
        lines.append(f"Section breaks: {self.section_breaks}")
        lines.append(f"Multi-column: {'yes' if self.multi_column else 'no'}")
        lines.append(f"Tracked changes: {'yes' if self.has_tracked_changes else 'no'}")
        lines.append(f"Comments: {'yes' if self.has_comments else 'no'}")
        if self.unsupported_elements:
            lines.append(f"Unsupported elements: {', '.join(self.unsupported_elements)}")
        return "\n".join(lines)

    def to_dict(self) -> dict:
        """Format as JSON-serializable dict."""
        return {
            "header_parts": self.header_parts,
            "header_paragraphs": self.header_paragraphs,
            "footer_parts": self.footer_parts,
            "footer_paragraphs": self.footer_paragraphs,
            "images": {
                "body": self.images_body,
                "header": self.images_header,
                "footer": self.images_footer,
                "total": self.images_body + self.images_header + self.images_footer,
            },
            "tables": {
                "simple": self.simple_tables,
                "skipped": self.skipped_tables,
                "skipped_reasons": self.skipped_table_reasons,
            },
            "section_breaks": self.section_breaks,
            "multi_column": self.multi_column,
            "has_tracked_changes": self.has_tracked_changes,
            "has_comments": self.has_comments,
            "unsupported_elements": self.unsupported_elements,
        }


def audit_document(doc: DocxDocument) -> AuditReport:
    """Perform a structural audit of *doc*.

    Args:
        doc: Parsed document.

    Returns:
        :class:`AuditReport` with counts and flags.
    """
    report = AuditReport()

    # Headers / footers
    report.header_parts = len(doc.header_trees)
    report.footer_parts = len(doc.footer_trees)

    for tree in doc.header_trees.values():
        for child in tree:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                report.header_paragraphs += 1

    for tree in doc.footer_trees.values():
        for child in tree:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                report.footer_paragraphs += 1

    # Images in body
    for para in doc.paragraphs:
        if xpath(para, ".//w:drawing") or xpath(para, ".//w:pict"):
            report.images_body += 1
    for tbl in xpath(doc.body, "w:tbl"):
        for _drawing in xpath(tbl, ".//w:drawing"):
            report.images_body += 1
        for _pict in xpath(tbl, ".//w:pict"):
            report.images_body += 1

    # Images in headers / footers
    for tree in doc.header_trees.values():
        for para in xpath(tree, ".//w:p"):
            if xpath(para, ".//w:drawing") or xpath(para, ".//w:pict"):
                report.images_header += 1
    for tree in doc.footer_trees.values():
        for para in xpath(tree, ".//w:p"):
            if xpath(para, ".//w:drawing") or xpath(para, ".//w:pict"):
                report.images_footer += 1

    # Tables
    from docx_mcp.table_utils import is_simple_table

    for tbl in xpath(doc.body, "w:tbl"):
        is_simple, reason = is_simple_table(tbl)
        if is_simple:
            report.simple_tables += 1
        else:
            report.skipped_tables += 1
            report.skipped_table_reasons.append(reason)

    # Section breaks / multi-column
    for sectPr in xpath(doc.body, "w:sectPr"):
        report.section_breaks += 1
        cols = xpath(sectPr, "w:cols")
        if cols:
            num = cols[0].get(qn("w", "num"))
            if num is not None and int(num) > 1:
                report.multi_column = True

    # Tracked changes
    report.has_tracked_changes = bool(doc.has_tracked_changes())

    # Comments
    report.has_comments = doc.comments_tree is not None

    # Unsupported elements
    if xpath(doc.document_tree, ".//w:footnoteReference"):
        report.unsupported_elements.append("footnotes")
    if xpath(doc.document_tree, ".//w:endnoteReference"):
        report.unsupported_elements.append("endnotes")
    if xpath(doc.document_tree, ".//w:txbxContent"):
        report.unsupported_elements.append("text boxes")

    return report
