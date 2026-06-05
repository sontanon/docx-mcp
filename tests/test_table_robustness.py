"""Tests for T2.3: table robustness improvements."""

from docx_mcp.document import DocxDocument
from docx_mcp.models import (
    TableChange,
    TableChangeType,
    TableInfo,
)
from docx_mcp.redliner import apply_redlines


class TestEmptyCellTable:
    """Tests for tables with empty cells."""

    def test_empty_cell_extraction(self, table_empty_cell_path):
        """Empty cell should appear as empty string in extraction."""
        from docx_mcp.converter import body_to_fragments

        doc = DocxDocument(path=table_empty_cell_path)
        result = body_to_fragments(doc.body_elements)

        # Find the table in items
        table_info = next(item for item in result.items if isinstance(item, TableInfo))
        assert table_info.cells[1][1].text == ""  # Row 2, Col 2 is empty

    def test_modify_empty_cell(self, table_empty_cell_path, tmp_path):
        """Modify an empty cell — should insert text as tracked change."""
        changes = [
            TableChange(
                kind="table",
                table_id=1,
                row=2,
                col=2,
                change_type=TableChangeType.MODIFY_CELL,
                new_text="Filled cell.",
                justification="Fill empty cell.",
            ),
        ]
        doc = apply_redlines(table_empty_cell_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        result = DocxDocument(path=out)
        # The table should have the new text
        table = result.body_elements[0]
        from docx_mcp.table_utils import get_cell_element
        cell = get_cell_element(table, 2, 2)
        cell_text = "".join(
            t.text or "" for t in cell.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        )
        assert "Filled cell." in cell_text


class TestWideTable:
    """Tests for very wide tables."""

    def test_wide_table_extraction(self, wide_table_path):
        """10-column table should be fully extracted."""
        from docx_mcp.converter import body_to_fragments

        doc = DocxDocument(path=wide_table_path)
        result = body_to_fragments(doc.body_elements)

        table_info = next(item for item in result.items if isinstance(item, TableInfo))
        assert table_info.cols == 10
        assert table_info.cells[0][9].text == "Col 10"
        assert table_info.cells[1][9].text == "Data 10"

    def test_modify_last_cell_wide_table(self, wide_table_path, tmp_path):
        """Modify the last cell (1.2.10) of a 10-column table."""
        changes = [
            TableChange(
                kind="table",
                table_id=1,
                row=2,
                col=10,
                change_type=TableChangeType.MODIFY_CELL,
                new_text="Updated last cell.",
                justification="Test wide table.",
            ),
        ]
        doc = apply_redlines(wide_table_path, changes)
        out = tmp_path / "output.docx"
        doc.save(out)

        result = DocxDocument(path=out)
        table = result.body_elements[0]
        from docx_mcp.table_utils import get_cell_element
        cell = get_cell_element(table, 2, 10)
        cell_text = "".join(
            t.text or "" for t in cell.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        )
        assert "Updated last cell." in cell_text


class TestErrorMessages:
    """Tests for improved error messages."""

    def test_skipped_table_reason_includes_table_id(self, merged_cell_table_path):
        """is_simple_table reason should include table ID when provided."""
        from docx_mcp.table_utils import is_simple_table

        doc = DocxDocument(path=merged_cell_table_path)
        table = doc.body_elements[0]
        is_simple, reason = is_simple_table(table, table_id=5)
        assert not is_simple
        assert "table 5," in reason
