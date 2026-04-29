"""Tests for T1.3: fix append_after bugs."""

from __future__ import annotations

from docx_mcp.document import DocxDocument
from docx_mcp.handlers.append import handle_append_after
from docx_mcp.id_manager import IdManager
from docx_mcp.models import ParagraphChange, ParagraphChangeType, RedlineConfig
from docx_mcp.namespaces import qn, xpath
from docx_mcp.redliner import _sort_paragraph_changes


class TestSortParagraphChanges:
    """Tests for _sort_paragraph_changes ordering."""

    def _make_element_map(self, ids):
        """Build a minimal element map for sorting tests."""
        return {str(i): None for i in ids}

    def test_append_same_fragment_reversed(self):
        """Multiple appends to the same fragment should be reversed."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=5,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="First append",
                justification="Test 1",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=5,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Second append",
                justification="Test 2",
            ),
        ]
        element_map = self._make_element_map([1, 2, 3, 4, 5])
        sorted_changes = _sort_paragraph_changes(changes, element_map)
        assert sorted_changes[0].new_text == "Second append"
        assert sorted_changes[1].new_text == "First append"

    def test_append_different_fragments_ordered(self):
        """Appends to different fragments should be in reverse fragment order."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Append to 3",
                justification="Test",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=5,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Append to 5",
                justification="Test",
            ),
        ]
        element_map = self._make_element_map([1, 2, 3, 4, 5])
        sorted_changes = _sort_paragraph_changes(changes, element_map)
        assert sorted_changes[0].fragment_id == "5"
        assert sorted_changes[1].fragment_id == "3"

    def test_mixed_changes_ordered_correctly(self):
        """Modify comes before delete comes before append."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=5,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Append",
                justification="Test",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.MODIFY,
                new_text="Modified",
                justification="Test",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=4,
                change_type=ParagraphChangeType.DELETE,
                justification="Test",
            ),
        ]
        element_map = self._make_element_map([1, 2, 3, 4, 5])
        sorted_changes = _sort_paragraph_changes(changes, element_map)
        assert sorted_changes[0].change_type == ParagraphChangeType.MODIFY
        assert sorted_changes[1].change_type == ParagraphChangeType.DELETE
        assert sorted_changes[2].change_type == ParagraphChangeType.APPEND_AFTER


class TestAppendAnnotationIdConsistency:
    """Tests that paragraph mark and content share the same annotation ID."""

    def test_inserted_paragraph_single_id(self, simple_5para_path):
        """Paragraph mark and content wrapper must share the same w:id."""
        doc = DocxDocument(path=simple_5para_path)
        paragraph = doc.paragraphs[0]
        id_manager = IdManager(start_after=doc.max_annotation_id())
        config = RedlineConfig()

        new_p, ins_id = handle_append_after(
            paragraph,
            "Appended text.",
            id_manager=id_manager,
            config=config,
        )

        # Paragraph mark <w:ins> inside w:pPr/w:rPr
        ppr = xpath(new_p, "w:pPr")[0]
        rpr = xpath(ppr, "w:rPr")[0]
        mark_ins = xpath(rpr, "w:ins")[0]
        mark_id = int(mark_ins.get(qn("w", "id")))

        # Content wrapper <w:ins> direct child of w:p
        content_ins = xpath(new_p, "w:ins")[0]
        content_id = int(content_ins.get(qn("w", "id")))

        assert mark_id == content_id == ins_id

    def test_multiple_appends_same_fragment_order(self, simple_5para_path, tmp_path):
        """Two appends to the same fragment should appear in correct text order."""
        from docx_mcp.redliner import apply_redlines

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="First appended paragraph.",
                justification="Test first.",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Second appended paragraph.",
                justification="Test second.",
            ),
        ]

        doc = apply_redlines(simple_5para_path, changes)
        out = tmp_path / "appended.docx"
        doc.save(out)

        # Re-read and check order
        result_doc = DocxDocument(path=out)
        paragraphs = result_doc.paragraphs

        # Find the two appended paragraphs (after fragment 1)
        texts = ["".join(t.text or "" for t in xpath(p, ".//w:t")) for p in paragraphs]

        # The original paragraph 1 should be first
        # Then "First appended paragraph."
        # Then "Second appended paragraph."
        assert texts[1] == "First appended paragraph."
        assert texts[2] == "Second appended paragraph."
