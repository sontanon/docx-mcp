"""Tests for T2.2: section-break handling in append_after."""

from __future__ import annotations

from docx_mcp.document import DocxDocument
from docx_mcp.models import ParagraphChange, ParagraphChangeType
from docx_mcp.redliner import apply_redlines


class TestSectionBreakHandling:
    """Tests that append_after behaves correctly near section breaks."""

    def test_append_before_section_break(self, two_section_path, tmp_path):
        """Appending after the last paragraph before a section break
        should place the new paragraph in the same section."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Appended paragraph in first section.",
                justification="Test section break handling.",
            ),
        ]
        doc = apply_redlines(two_section_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        # Verify the appended paragraph is before the section-break paragraph
        result = DocxDocument(path=out)
        paragraphs = result.paragraphs
        texts = [
            "".join(
                t.text or ""
                for t in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            )
            for p in paragraphs
        ]

        # Document has: p1, p_break (empty, contains sectPr), p2
        # After append: p1, p_appended, p_break, p2
        assert "Paragraph in first section." in texts[0]
        assert "Appended paragraph in first section." in texts[1]
        # The section break paragraph should still be present (empty text)
        assert texts[2] == ""
        assert "Paragraph in second section." in texts[3]

    def test_extraction_on_multi_section_document(self, two_section_path):
        """extract_fragments should work on multi-section documents."""
        from docx_mcp.converter import full_to_fragments

        doc = DocxDocument(path=two_section_path)
        result = full_to_fragments(doc)

        # Should have at least 2 body paragraphs
        body_fragments = [
            item for item in result.items
            if isinstance(item, tuple) and not item[0].startswith(("header", "footer"))
        ]
        assert len(body_fragments) >= 2

    def test_audit_reports_section_breaks(self, two_section_path):
        """audit_document should report section breaks."""
        from docx_mcp.audit import audit_document

        doc = DocxDocument(path=two_section_path)
        report = audit_document(doc)
        assert report.section_breaks >= 1
