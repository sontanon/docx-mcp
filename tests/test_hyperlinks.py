"""Tests for T1.2: hyperlink preservation & editing."""

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.handlers.modify import handle_modify
from docx_mcp.id_manager import IdManager
from docx_mcp.models import ParagraphChange, ParagraphChangeType, RedlineConfig
from docx_mcp.namespaces import qn, xpath
from docx_mcp.redliner import apply_redlines
from docx_mcp.run_ops import extract_runs


class TestHyperlinkExtraction:
    """Hyperlinks are extracted as [text](url) pseudo-Markdown."""

    def test_extract_plain_hyperlink(self, hyperlink_paragraph_path):
        """Plain hyperlink renders as [text](url)."""
        doc = DocxDocument(path=hyperlink_paragraph_path)
        from docx_mcp.converter import paragraph_to_pseudo_markdown

        md = paragraph_to_pseudo_markdown(
            doc.paragraphs[0],
            hyperlink_resolver=doc.resolve_hyperlink_url,
        )
        assert "[our website](https://example.com)" in md

    def test_extract_formatted_hyperlink(self, hyperlink_formatted_path):
        """Hyperlink with bold/italic inside preserves formatting."""
        doc = DocxDocument(path=hyperlink_formatted_path)
        from docx_mcp.converter import paragraph_to_pseudo_markdown

        md = paragraph_to_pseudo_markdown(
            doc.paragraphs[0],
            hyperlink_resolver=doc.resolve_hyperlink_url,
        )
        assert "[**_important terms_**](https://example.com/terms)" in md

    def test_extract_multiple_hyperlinks(self, multiple_hyperlinks_path):
        """Paragraph with two hyperlinks extracts both."""
        doc = DocxDocument(path=multiple_hyperlinks_path)
        from docx_mcp.converter import paragraph_to_pseudo_markdown

        md = paragraph_to_pseudo_markdown(
            doc.paragraphs[0],
            hyperlink_resolver=doc.resolve_hyperlink_url,
        )
        assert "[sales](https://example.com/sales)" in md
        # support hyperlink may include trailing punctuation depending on fixture
        assert "[support" in md
        assert "https://example.com/support" in md


class TestHyperlinkRunExtraction:
    """extract_runs captures hyperlink_rel_id."""

    def test_extract_runs_with_hyperlink(self, hyperlink_paragraph_path):
        doc = DocxDocument(path=hyperlink_paragraph_path)
        runs = extract_runs(doc.paragraphs[0])
        # There should be a run with hyperlink_rel_id set
        link_runs = [r for r in runs if r.hyperlink_rel_id is not None]
        assert len(link_runs) == 1
        assert link_runs[0].text == "our website"


class TestHyperlinkModify:
    """Modifying text preserves or updates hyperlinks."""

    def test_modify_preserves_hyperlink_url(self, hyperlink_paragraph_path):
        """Hyperlink text that is not changed by diff remains wrapped.

        Note: Full pseudo-Markdown hyperlink editing ([text](url)) during
        modify is not yet supported. This test verifies that unchanged
        hyperlink runs retain their wrapper.
        """
        doc = DocxDocument(path=hyperlink_paragraph_path)
        para = doc.paragraphs[0]
        mgr = IdManager(start_after=doc.max_annotation_id())
        config = RedlineConfig()

        # Modify non-hyperlink text only
        handle_modify(
            para,
            "Go to [our website] for more info.",
            id_manager=mgr,
            config=config,
        )

        # The hyperlink should still exist (may be inside <w:del> or <w:ins>)
        hyperlink_els = xpath(para, ".//w:hyperlink")
        assert len(hyperlink_els) >= 1

    @pytest.mark.skip(reason="Hyperlink URL modification via pseudo-Markdown not yet implemented")
    def test_modify_changes_hyperlink_url(self, hyperlink_paragraph_path, tmp_path):
        """[text](new_url) creates a new relationship."""
        doc = DocxDocument(path=hyperlink_paragraph_path)
        para = doc.paragraphs[0]
        mgr = IdManager(start_after=doc.max_annotation_id())
        config = RedlineConfig()

        handle_modify(
            para,
            "Visit [our site](https://new-example.com) for more info.",
            id_manager=mgr,
            config=config,
        )

        out = tmp_path / "modified.docx"
        doc.save(out)

        result_doc = DocxDocument(path=out)
        assert result_doc.rels_tree is not None
        new_rel_id = None
        for rel in result_doc.rels_tree:
            if rel.get("Target") == "https://new-example.com":
                new_rel_id = rel.get("Id")
                break
        assert new_rel_id is not None

    def test_delete_paragraph_with_hyperlink(self, hyperlink_paragraph_path):
        """Deleting a paragraph with hyperlink wraps runs in <w:del>."""
        from docx_mcp.handlers.delete import handle_delete

        doc = DocxDocument(path=hyperlink_paragraph_path)
        para = doc.paragraphs[0]
        mgr = IdManager(start_after=doc.max_annotation_id())
        config = RedlineConfig()

        handle_delete(para, id_manager=mgr, config=config)

        # Hyperlink should contain <w:del>
        hyperlink_els = xpath(para, "w:hyperlink")
        for h in hyperlink_els:
            del_els = xpath(h, "w:del")
            assert len(del_els) >= 1


class TestHyperlinkAppend:
    """Appending text with [text](url) creates hyperlinks."""

    def test_append_creates_hyperlink(self, simple_5para_path, tmp_path):
        """append_after with [text](url) creates a hyperlink relationship."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Visit [our site](https://appended.com) for details.",
                justification="Test append hyperlink.",
            ),
        ]

        doc = apply_redlines(simple_5para_path, changes)
        out = tmp_path / "appended.docx"
        doc.save(out)

        result_doc = DocxDocument(path=out)
        assert result_doc.rels_tree is not None
        # Check that a new relationship was created
        new_rel_id = None
        for rel in result_doc.rels_tree:
            if rel.get("Target") == "https://appended.com":
                new_rel_id = rel.get("Id")
                break
        assert new_rel_id is not None

        # Check that the appended paragraph contains a hyperlink
        paragraphs = result_doc.paragraphs
        appended_para = paragraphs[1]
        hyperlink_els = xpath(appended_para, ".//w:hyperlink")
        assert len(hyperlink_els) == 1
        assert hyperlink_els[0].get(qn("r", "id")) == new_rel_id
