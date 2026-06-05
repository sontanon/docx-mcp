"""Tests for T2.1: header/footer extraction and redlining."""

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.models import ParagraphChange, ParagraphChangeType
from docx_mcp.namespaces import xpath
from docx_mcp.redliner import apply_redlines


class TestHeaderFooterExtraction:
    """Tests for extracting header/footer text as fragments."""

    def test_header_footer_in_tagged_format(self, header_footer_text_path):
        """Tagged extraction includes header and footer paragraphs."""
        from docx_mcp.converter import fragments_to_tagged_text_interleaved, full_to_fragments

        doc = DocxDocument(path=header_footer_text_path)
        result = full_to_fragments(doc)
        text = fragments_to_tagged_text_interleaved(result.items)

        assert "<f=header_1.1>Header paragraph text.</f=header_1.1>" in text
        assert "<f=footer_1.1>Footer paragraph text.</f=footer_1.1>" in text
        assert "<f=1>First body paragraph.</f=1>" in text
        assert "<f=2>Second body paragraph.</f=2>" in text

    def test_header_footer_in_json_format(self, header_footer_text_path):
        """JSON extraction includes header and footer fragments."""
        from docx_mcp.converter import fragments_to_json_interleaved, full_to_fragments

        doc = DocxDocument(path=header_footer_text_path)
        result = full_to_fragments(doc)
        data = fragments_to_json_interleaved(result.items)

        header_frag = next((f for f in data if f.get("fragment_id") == "header_1.1"), None)
        footer_frag = next((f for f in data if f.get("fragment_id") == "footer_1.1"), None)

        assert header_frag is not None
        assert header_frag["text"] == "Header paragraph text."
        assert footer_frag is not None
        assert footer_frag["text"] == "Footer paragraph text."

    def test_full_element_map_keys(self, header_footer_text_path):
        """full_element_map contains body, header, and footer keys."""
        doc = DocxDocument(path=header_footer_text_path)
        element_map = doc.full_element_map()

        assert "1" in element_map
        assert "2" in element_map
        assert "header_1.1" in element_map
        assert "footer_1.1" in element_map

    def test_resolve_fragment_id(self, header_footer_text_path):
        """resolve_fragment_id returns correct element and tree type."""
        doc = DocxDocument(path=header_footer_text_path)

        el, tree_type = doc.resolve_fragment_id("1")
        assert tree_type == "body"
        tag = etree.QName(el.tag).localname
        assert tag == "p"

        el, tree_type = doc.resolve_fragment_id("header_1.1")
        assert tree_type == "header"

        el, tree_type = doc.resolve_fragment_id("footer_1.1")
        assert tree_type == "footer"


class TestHeaderFooterRedlining:
    """Tests for applying tracked changes to headers and footers."""

    def test_modify_header(self, header_footer_text_path, tmp_path):
        """Modify a header paragraph."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id="header_1.1",
                change_type=ParagraphChangeType.MODIFY,
                new_text="Modified header text.",
                justification="Test header modify.",
            ),
        ]
        doc = apply_redlines(header_footer_text_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        # Re-read and verify
        result = DocxDocument(path=out)
        header_tree = next(iter(result.header_trees.values()))
        header_paras = xpath(header_tree, ".//w:p")
        header_text = "".join(t.text or "" for t in header_paras[0].iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
        assert "Modified header" in header_text
        assert "text." in header_text

    def test_delete_footer(self, header_footer_text_path, tmp_path):
        """Delete a footer paragraph."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id="footer_1.1",
                change_type=ParagraphChangeType.DELETE,
                justification="Test footer delete.",
            ),
        ]
        doc = apply_redlines(header_footer_text_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        result = DocxDocument(path=out)
        footer_tree = next(iter(result.footer_trees.values()))
        footer_paras = xpath(footer_tree, ".//w:p")
        # The paragraph should contain a <w:del> wrapper
        del_els = xpath(footer_paras[0], "w:del")
        assert len(del_els) > 0

    def test_append_after_header(self, header_footer_text_path, tmp_path):
        """Append a paragraph after a header paragraph."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id="header_1.1",
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Appended header paragraph.",
                justification="Test header append.",
            ),
        ]
        doc = apply_redlines(header_footer_text_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        result = DocxDocument(path=out)
        header_tree = next(iter(result.header_trees.values()))
        header_paras = xpath(header_tree, ".//w:p")
        assert len(header_paras) == 2
        texts = [
            "".join(t.text or "" for t in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))
            for p in header_paras
        ]
        assert "Header paragraph text." in texts[0]
        assert "Appended header paragraph." in texts[1]

    def test_header_serialization_roundtrip(self, header_footer_text_path, tmp_path):
        """Modifications to headers survive save/load roundtrip."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id="header_1.1",
                change_type=ParagraphChangeType.MODIFY,
                new_text="Roundtrip header.",
                justification="Test serialization.",
            ),
        ]
        doc = apply_redlines(header_footer_text_path, changes)
        out = tmp_path / "roundtrip.docx"
        doc.save(out)

        result = DocxDocument(path=out)
        header_tree = next(iter(result.header_trees.values()))
        header_paras = xpath(header_tree, ".//w:p")
        header_text = "".join(
            t.text or "" for t in header_paras[0].iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        )
        assert "Roundtrip header." in header_text


class TestHeaderFooterComments:
    """Tests for comments on header/footer paragraphs."""

    def test_comment_on_header_paragraph_is_skipped_with_warning(
        self,
        header_footer_text_path,
        tmp_path,
    ):
        """Comments in headers/footers are skipped with a warning.

        Word and LibreOffice do not support commentRangeStart/End in
        header/footer XML parts. LibreOffice rejects the file as corrupt.
        """
        import warnings

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id="header_1.1",
                change_type=ParagraphChangeType.MODIFY,
                new_text="Modified header with comment.",
                justification="Header comment test.",
            ),
        ]
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            doc = apply_redlines(header_footer_text_path, changes)

        out = tmp_path / "output.docx"
        doc.save(out)

        # Warning should be raised
        assert len(w) == 1
        assert "header_1.1" in str(w[0].message)
        assert "silently dropped" in str(w[0].message)

        # No comments.xml should be created for header-only changes
        result = DocxDocument(path=out)
        assert result.comments_tree is None

        # But the tracked change itself should still work
        header_text = ""
        for tree in result._header_trees.values():
            header_text += "".join(
                t.text or "" for t in tree.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            )
        assert "Modified header with comment." in header_text
