"""Tests for table cell change application."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.converter import body_to_fragments, paragraph_to_pseudo_markdown
from docx_mcp.models import ChangeType, RedlineConfig, TableChange, TableInfo
from docx_mcp.namespaces import xpath
from docx_mcp.redliner import apply_redlines
from docx_mcp.table_utils import get_cell_element, get_cell_paragraphs


def _default_config() -> RedlineConfig:
    return RedlineConfig(author="AI Review")


class TestModifyCell:
    """Tests for modify_cell operations."""

    def test_modify_single_para_cell(self, simple_table_path: Path) -> None:
        """Modify a cell with a single paragraph."""
        changes = [
            TableChange(
                table_id=2,
                row=1,
                col=1,
                cell_id="2.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Modified Header",
                justification="Update header text",
            ),
        ]

        doc = apply_redlines(simple_table_path, [], table_changes=changes)

        # Verify the cell was modified
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) >= 1
        tbl = tables[0]
        tc = get_cell_element(tbl, row=1, col=1)
        paras = get_cell_paragraphs(tc)
        assert len(paras) == 1

        # Check for tracked changes
        md = paragraph_to_pseudo_markdown(paras[0], markup=True)
        assert "++" in md or "~~" in md  # Should have insertions or deletions

        # Check comment was added
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) >= 1
        # Check justification text is in one of the comments
        comment_texts = ["".join(t.text for t in xpath(c, ".//w:t") if t.text) for c in comments]
        assert any("Update header text" in text for text in comment_texts)

    def test_modify_multi_para_cell_equal_count(self, table_multi_para_path: Path) -> None:
        """Modify a cell where new text has same paragraph count as old."""
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=1,
                cell_id="2.2.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="New first para.\nNew second para.",
                justification="Update both paragraphs",
            ),
        ]

        doc = apply_redlines(table_multi_para_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=2, col=1)
        paras = get_cell_paragraphs(tc)

        # Should still have 2 paragraphs
        assert len(paras) == 2

        # Both should have tracked changes
        for para in paras:
            md = paragraph_to_pseudo_markdown(para, markup=True)
            assert "++" in md or "~~" in md

    def test_modify_multi_para_cell_reduce_count(self, table_multi_para_path: Path) -> None:
        """Modify a cell where new text has fewer paragraphs than old."""
        # Original cell 2.2.2 has 3 paragraphs
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=2,
                cell_id="2.2.2",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Single combined paragraph",
                justification="Consolidate paragraphs",
            ),
        ]

        doc = apply_redlines(table_multi_para_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=2, col=2)
        paras = get_cell_paragraphs(tc)

        # Should still have 3 paragraphs (excess marked as deleted with preserved marks)
        assert len(paras) == 3

        # First para should have modifications
        md0 = paragraph_to_pseudo_markdown(paras[0], markup=True)
        assert "++" in md0 or "~~" in md0

        # Second and third paras should be fully deleted
        for para in paras[1:]:
            md = paragraph_to_pseudo_markdown(para, markup=True)
            # Should have deletion markers
            assert "~~" in md or len(md) == 0

    def test_modify_multi_para_cell_increase_count(self, table_multi_para_path: Path) -> None:
        """Modify a cell where new text has more paragraphs than old."""
        # Original cell 2.2.1 has 2 paragraphs
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=1,
                cell_id="2.2.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="First.\nSecond.\nThird.\nFourth.",
                justification="Add more paragraphs",
            ),
        ]

        doc = apply_redlines(table_multi_para_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=2, col=1)
        paras = get_cell_paragraphs(tc)

        # Should now have 4 paragraphs (2 modified + 2 inserted)
        assert len(paras) == 4

        # All should have tracked changes
        for para in paras:
            md = paragraph_to_pseudo_markdown(para, markup=True)
            # New ones will be entirely inside ++, modified ones will have ++ and/or ~~
            assert "++" in md or "~~" in md

    def test_modify_cell_with_formatting(self, formatted_table_path: Path) -> None:
        """Modify a cell that has formatting."""
        changes = [
            TableChange(
                table_id=2,
                row=1,
                col=1,
                cell_id="2.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="**Updated Bold** Header",
                justification="Change header",
            ),
        ]

        doc = apply_redlines(formatted_table_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=1, col=1)
        paras = get_cell_paragraphs(tc)

        md = paragraph_to_pseudo_markdown(paras[0], markup=True)
        assert "++" in md or "~~" in md

    def test_modify_cell_without_new_text_raises(self, simple_table_path: Path) -> None:
        """Modify cell without new_text should raise error."""
        changes = [
            TableChange(
                table_id=2,
                row=1,
                col=1,
                cell_id="2.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text=None,
                justification="Missing text",
            ),
        ]

        with pytest.raises(ValueError, match="requires new_text"):
            apply_redlines(simple_table_path, [], table_changes=changes)


class TestClearCell:
    """Tests for clear_cell operations."""

    def test_clear_single_para_cell(self, simple_table_path: Path) -> None:
        """Clear a cell with a single paragraph."""
        changes = [
            TableChange(
                table_id=2,
                row=1,
                col=1,
                cell_id="2.1.1",
                change_type=ChangeType.CLEAR_CELL,
                justification="Remove header",
            ),
        ]

        doc = apply_redlines(simple_table_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=1, col=1)
        paras = get_cell_paragraphs(tc)

        # Cell should still have paragraph (preserved mark)
        assert len(paras) == 1

        # Content should be deleted
        md = paragraph_to_pseudo_markdown(paras[0], markup=True)
        assert "~~" in md or len(md) == 0

        # Check comment
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) >= 1
        comment_texts = ["".join(t.text for t in xpath(c, ".//w:t") if t.text) for c in comments]
        assert any("Remove header" in text for text in comment_texts)

    def test_clear_multi_para_cell(self, table_multi_para_path: Path) -> None:
        """Clear a cell with multiple paragraphs."""
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=2,
                cell_id="2.2.2",
                change_type=ChangeType.CLEAR_CELL,
                justification="Clear all content",
            ),
        ]

        doc = apply_redlines(table_multi_para_path, [], table_changes=changes)

        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=2, col=2)
        paras = get_cell_paragraphs(tc)

        # All 3 paragraphs should still exist (preserved marks)
        assert len(paras) == 3

        # All should be deleted
        for para in paras:
            md = paragraph_to_pseudo_markdown(para, markup=True)
            # Each paragraph should have deletion or be empty
            assert "~~" in md or len(md) == 0


class TestTableChangeValidation:
    """Tests for validation of table changes."""

    def test_invalid_table_id_raises(self, simple_table_path: Path) -> None:
        """Table ID that doesn't exist should raise error."""
        changes = [
            TableChange(
                table_id=99,
                row=1,
                col=1,
                cell_id="99.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Text",
                justification="Bad table ID",
            ),
        ]

        with pytest.raises(ValueError, match="table_id=99"):
            apply_redlines(simple_table_path, [], table_changes=changes)

    def test_table_id_pointing_to_paragraph_raises(self, mixed_content_path: Path) -> None:
        """Table ID that points to a paragraph should raise error."""
        changes = [
            TableChange(
                table_id=1,  # First element is a paragraph
                row=1,
                col=1,
                cell_id="1.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Text",
                justification="Wrong element type",
            ),
        ]

        with pytest.raises(ValueError, match="is <w:p>, not <w:tbl>"):
            apply_redlines(mixed_content_path, [], table_changes=changes)

    def test_row_out_of_range_raises(self, simple_table_path: Path) -> None:
        """Row index beyond table bounds should raise error."""
        changes = [
            TableChange(
                table_id=2,
                row=99,
                col=1,
                cell_id="2.99.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Text",
                justification="Bad row",
            ),
        ]

        with pytest.raises(ValueError, match="row 99 out of range"):
            apply_redlines(simple_table_path, [], table_changes=changes)

    def test_col_out_of_range_raises(self, simple_table_path: Path) -> None:
        """Column index beyond table bounds should raise error."""
        changes = [
            TableChange(
                table_id=2,
                row=1,
                col=99,
                cell_id="2.1.99",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Text",
                justification="Bad col",
            ),
        ]

        with pytest.raises(ValueError, match="col 99 out of range"):
            apply_redlines(simple_table_path, [], table_changes=changes)


class TestMixedChanges:
    """Tests for applying both paragraph and table changes together."""

    def test_paragraph_and_table_changes_together(self, mixed_content_path: Path) -> None:
        """Apply both paragraph and table changes in same operation."""
        from docx_mcp.models import Change

        para_changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Modified intro paragraph.",
                justification="Update intro",
            ),
        ]

        table_changes = [
            TableChange(
                table_id=2,
                row=1,
                col=1,
                cell_id="2.1.1",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Modified cell",
                justification="Update cell",
            ),
        ]

        doc = apply_redlines(mixed_content_path, para_changes, table_changes=table_changes)

        # Check paragraph change
        paras = doc.paragraphs
        md_para = paragraph_to_pseudo_markdown(paras[0], markup=True)
        assert "++" in md_para or "~~" in md_para

        # Check table change
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=1, col=1)
        cell_paras = get_cell_paragraphs(tc)
        md_cell = paragraph_to_pseudo_markdown(cell_paras[0], markup=True)
        assert "++" in md_cell or "~~" in md_cell

        # Check both comments exist
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) >= 2
        comment_texts = ["".join(t.text for t in xpath(c, ".//w:t") if t.text) for c in comments]
        assert any("Update intro" in text for text in comment_texts)
        assert any("Update cell" in text for text in comment_texts)


class TestCommentAttachment:
    """Tests for comment attachment to table cells."""

    def test_comment_attached_to_first_para_in_cell(self, table_multi_para_path: Path) -> None:
        """Comment should be attached to first paragraph of multi-para cell."""
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=2,
                cell_id="2.2.2",
                change_type=ChangeType.MODIFY_CELL,
                new_text="Modified",
                justification="Test comment attachment",
            ),
        ]

        doc = apply_redlines(table_multi_para_path, [], table_changes=changes)

        # Get first paragraph of the cell
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=2, col=2)
        paras = get_cell_paragraphs(tc)
        first_para = paras[0]

        # Check for comment range markers
        comment_starts = xpath(first_para, ".//w:commentRangeStart")
        comment_ends = xpath(first_para, ".//w:commentRangeEnd")
        assert len(comment_starts) > 0
        assert len(comment_ends) > 0


class TestRoundTrip:
    """Tests for extracting modified table content."""

    def test_extract_modified_table_with_markup(self, simple_table_path: Path) -> None:
        """Extract table after modification shows tracked changes in markup mode."""
        changes = [
            TableChange(
                table_id=2,
                row=2,
                col=2,
                cell_id="2.2.2",
                change_type=ChangeType.MODIFY_CELL,
                new_text="New content",
                justification="Test extraction",
            ),
        ]

        doc = apply_redlines(simple_table_path, [], table_changes=changes)

        # Extract with markup=True
        items = body_to_fragments(doc.body_elements, markup=True)
        assert len(items) == 3

        # Second item is the table
        table_info = items[1]
        assert isinstance(table_info, TableInfo)

        # Check cell 2.2.2 has tracked-change markers
        cell_text = table_info.cells[1][1].text
        assert "++" in cell_text or "~~" in cell_text
