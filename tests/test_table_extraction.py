"""Tests for table extraction functions in converter.py."""

import json
from pathlib import Path

from docx_mcp.converter import (
    body_to_fragments,
    fragments_to_json_interleaved,
    fragments_to_tagged_text_interleaved,
)
from docx_mcp.document import DocxDocument
from docx_mcp.models import TableInfo


class TestBodyToFragments:
    """Tests for body_to_fragments()."""

    def test_simple_table_extraction(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Simple table doc has: para, table, para
        assert len(items) == 3
        assert isinstance(items[0], tuple)
        item = items[1]
        assert isinstance(item, TableInfo)
        assert item.table_id == 2  # Second element
        assert item.rows == 3
        assert item.cols == 3
        assert len(item.cells) == 3
        assert len(item.cells[0]) == 3

    def test_formatted_table_extraction(self, formatted_table_path: Path) -> None:
        doc = DocxDocument(path=formatted_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Formatted table doc has: para, table
        assert len(items) == 2
        item = items[1]
        assert isinstance(item, TableInfo)
        assert item.table_id == 2
        assert item.rows == 2
        assert item.cols == 2

        # Check that formatting is preserved in cell text
        cell_1_1 = item.cells[0][0]
        assert "**" in cell_1_1.text or "__" in cell_1_1.text or "_" in cell_1_1.text

    def test_multi_paragraph_cell_extraction(self, table_multi_para_path: Path) -> None:
        doc = DocxDocument(path=table_multi_para_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Multi-para doc has: para, table
        assert len(items) == 2
        item = items[1]
        assert isinstance(item, TableInfo)

        # Row 2, Col 1 has 2 paragraphs
        cell_2_1 = item.cells[1][0]
        assert "\n" in cell_2_1.text
        para_count = cell_2_1.text.count("\n") + 1
        assert para_count == 2

        # Row 2, Col 2 has 3 paragraphs
        cell_2_2 = item.cells[1][1]
        para_count = cell_2_2.text.count("\n") + 1
        assert para_count == 3

    def test_merged_cell_table_extracted(self, merged_cell_table_path: Path) -> None:
        doc = DocxDocument(path=merged_cell_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Merged cell doc has: para, table
        assert len(items) == 2
        item = items[1]
        assert isinstance(item, TableInfo)
        assert item.table_id == 2
        assert item.rows == 2
        assert item.cols == 3

        # Row 1: cell 1 spans 2 columns, cell 2 is spanned over, cell 3 is normal
        assert item.cells[0][0].span == 2
        assert item.cells[0][0].text == "A\nB"
        assert item.cells[0][1].span == 0
        assert item.cells[0][1].text == ""
        assert item.cells[0][2].span == 1
        assert item.cells[0][2].text == "C"

    def test_mixed_content_interleaved(self, mixed_content_path: Path) -> None:
        doc = DocxDocument(path=mixed_content_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Mixed content has: para, table(2x2), para, table(1x3), para
        assert len(items) == 5

        # First item: paragraph
        assert isinstance(items[0], tuple)
        assert items[0][0] == "1"

        # Second item: table
        assert isinstance(items[1], TableInfo)
        assert items[1].table_id == 2
        assert items[1].rows == 2
        assert items[1].cols == 2

        # Third item: paragraph
        assert isinstance(items[2], tuple)
        assert items[2][0] == "3"

        # Fourth item: table
        assert isinstance(items[3], TableInfo)
        assert items[3].table_id == 4
        assert items[3].rows == 1
        assert items[3].cols == 3

        # Fifth item: paragraph
        assert isinstance(items[4], tuple)
        assert items[4][0] == "5"

    def test_cell_ids_are_correct(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items

        # Table is second element
        item = items[1]
        assert isinstance(item, TableInfo)

        # Check a few cell IDs
        assert item.cells[0][0].cell_id == "2.1.1"
        assert item.cells[0][1].cell_id == "2.1.2"
        assert item.cells[1][1].cell_id == "2.2.2"
        assert item.cells[2][2].cell_id == "2.3.3"


class TestFragmentsToTaggedTextInterleaved:
    """Tests for fragments_to_tagged_text_interleaved()."""

    def test_simple_table_tagged_output(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        tagged = fragments_to_tagged_text_interleaved(items)

        # Should have table tags (table is ID 2)
        assert "<table=2 rows=3 cols=3>" in tagged
        assert "</table=2>" in tagged

        # Should have cell tags
        assert "<cell=2.1.1>" in tagged
        assert "</cell=2.1.1>" in tagged
        assert "<cell=2.3.3>" in tagged
        assert "</cell=2.3.3>" in tagged

    def test_mixed_content_tagged_output(self, mixed_content_path: Path) -> None:
        doc = DocxDocument(path=mixed_content_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        tagged = fragments_to_tagged_text_interleaved(items)

        # Should have paragraph tags
        assert "<f=1>" in tagged
        assert "</f=1>" in tagged
        assert "<f=3>" in tagged
        assert "<f=5>" in tagged

        # Should have table tags
        assert "<table=2 rows=2 cols=2>" in tagged
        assert "</table=2>" in tagged
        assert "<table=4 rows=1 cols=3>" in tagged
        assert "</table=4>" in tagged

    def test_merged_cell_table_tagged_output(self, merged_cell_table_path: Path) -> None:
        doc = DocxDocument(path=merged_cell_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        tagged = fragments_to_tagged_text_interleaved(items)

        # Should have table with span markers (spanned-over cells omitted)
        assert "<table=2 rows=2 cols=3>" in tagged
        assert '<cell=2.1.1 span="2">A\nB</cell=2.1.1>' in tagged
        assert '<cell=2.1.2' not in tagged  # spanned-over cell omitted
        assert "</table=2>" in tagged

    def test_multi_paragraph_cell_preserves_newlines(self, table_multi_para_path: Path) -> None:
        doc = DocxDocument(path=table_multi_para_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        _tagged = fragments_to_tagged_text_interleaved(items)

        # Multi-paragraph cells should have newlines in their text
        # Table is second element
        assert isinstance(items[1], TableInfo)
        cell_2_1 = items[1].cells[1][0]
        assert "\n" in cell_2_1.text


class TestFragmentsToJsonInterleaved:
    """Tests for fragments_to_json_interleaved()."""

    def test_simple_table_json_output(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        json_list = fragments_to_json_interleaved(items)

        # Should have 3 items: para, table, para
        assert len(json_list) == 3
        table_dict = json_list[1]

        assert table_dict["type"] == "table"
        assert table_dict["table_id"] == 2
        assert table_dict["rows"] == 3
        assert table_dict["cols"] == 3
        assert "cells" in table_dict
        assert len(table_dict["cells"]) == 3
        assert len(table_dict["cells"][0]) == 3

        # Check cell structure
        cell_1_1 = table_dict["cells"][0][0]
        assert cell_1_1["cell_id"] == "2.1.1"
        assert cell_1_1["row"] == 1
        assert cell_1_1["col"] == 1
        assert "text" in cell_1_1

    def test_mixed_content_json_output(self, mixed_content_path: Path) -> None:
        doc = DocxDocument(path=mixed_content_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        json_list = fragments_to_json_interleaved(items)

        assert len(json_list) == 5

        # First item: paragraph
        assert json_list[0]["type"] == "paragraph"
        assert json_list[0]["fragment_id"] == "1"
        assert "text" in json_list[0]

        # Second item: table
        assert json_list[1]["type"] == "table"
        assert json_list[1]["table_id"] == 2
        assert json_list[1]["rows"] == 2
        assert json_list[1]["cols"] == 2

        # Third item: paragraph
        assert json_list[2]["type"] == "paragraph"
        assert json_list[2]["fragment_id"] == "3"

        # Fourth item: table
        assert json_list[3]["type"] == "table"
        assert json_list[3]["table_id"] == 4

        # Fifth item: paragraph
        assert json_list[4]["type"] == "paragraph"
        assert json_list[4]["fragment_id"] == "5"

    def test_merged_cell_table_json_output(self, merged_cell_table_path: Path) -> None:
        doc = DocxDocument(path=merged_cell_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        json_list = fragments_to_json_interleaved(items)

        # Should have 2 items: para, table
        assert len(json_list) == 2
        table_dict = json_list[1]

        assert table_dict["type"] == "table"
        assert table_dict["table_id"] == 2
        assert table_dict["rows"] == 2
        assert table_dict["cols"] == 3

        # Check span fields in cells (spanned-over cells omitted)
        cells = table_dict["cells"]
        assert cells[0][0]["span"] == 2
        assert len(cells[0]) == 2  # cell 1.1.2 omitted, only 1.1.1 and 1.1.3 remain
        assert "span" not in cells[0][1]  # span=1 is omitted

    def test_json_is_serializable(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        json_list = fragments_to_json_interleaved(items)

        # Should be JSON-serializable
        json_str = json.dumps(json_list)
        assert len(json_str) > 0

        # Should be JSON-deserializable
        parsed = json.loads(json_str)
        assert len(parsed) == len(json_list)

    def test_multi_paragraph_cell_json(self, table_multi_para_path: Path) -> None:
        doc = DocxDocument(path=table_multi_para_path)
        result = body_to_fragments(doc.body_elements)
        items = result.items
        json_list = fragments_to_json_interleaved(items)

        # Table is second item
        table_dict = json_list[1]
        # Row 2, Col 1 has 2 paragraphs (joined with \n)
        cell_2_1 = table_dict["cells"][1][0]
        assert "\n" in cell_2_1["text"]

        # Row 2, Col 2 has 3 paragraphs
        cell_2_2 = table_dict["cells"][1][1]
        para_count = cell_2_2["text"].count("\n") + 1
        assert para_count == 3
