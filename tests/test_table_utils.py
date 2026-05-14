"""Tests for table utility functions."""

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.namespaces import xpath
from docx_mcp.table_utils import (
    get_cell_element,
    get_cell_paragraphs,
    is_simple_table,
    parse_cell_id,
    table_dimensions,
)


class TestIsSimpleTable:
    """Tests for is_simple_table()."""

    def test_simple_table_returns_true(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 1

        is_simple, reason = is_simple_table(tables[0])
        assert is_simple is True
        assert reason == ""

    def test_merged_cell_table_returns_false(self, merged_cell_table_path: Path) -> None:
        doc = DocxDocument(path=merged_cell_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 1

        is_simple, reason = is_simple_table(tables[0])
        assert is_simple is False
        assert "gridSpan" in reason or "merge" in reason.lower()

    def test_formatted_table_is_simple(self, formatted_table_path: Path) -> None:
        doc = DocxDocument(path=formatted_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 1

        is_simple, _reason = is_simple_table(tables[0])
        assert is_simple is True


class TestTableDimensions:
    """Tests for table_dimensions()."""

    def test_simple_table_dimensions(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 1

        rows, cols = table_dimensions(tables[0])
        assert rows == 3
        assert cols == 3

    def test_formatted_table_dimensions(self, formatted_table_path: Path) -> None:
        doc = DocxDocument(path=formatted_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 1

        rows, cols = table_dimensions(tables[0])
        assert rows == 2
        assert cols == 2

    def test_mixed_content_first_table(self, mixed_content_path: Path) -> None:
        doc = DocxDocument(path=mixed_content_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 2

        rows, cols = table_dimensions(tables[0])
        assert rows == 2
        assert cols == 2

    def test_mixed_content_second_table(self, mixed_content_path: Path) -> None:
        doc = DocxDocument(path=mixed_content_path)
        tables = xpath(doc.body, ".//w:tbl")
        assert len(tables) == 2

        rows, cols = table_dimensions(tables[1])
        assert rows == 1
        assert cols == 3


class TestGetCellElement:
    """Tests for get_cell_element()."""

    def test_get_first_cell(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        tc = get_cell_element(tbl, row=1, col=1)
        assert tc is not None
        assert tc.tag.endswith("}tc")

    def test_get_last_cell(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        tc = get_cell_element(tbl, row=3, col=3)
        assert tc is not None
        assert tc.tag.endswith("}tc")

    def test_get_middle_cell(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        tc = get_cell_element(tbl, row=2, col=2)
        assert tc is not None
        assert tc.tag.endswith("}tc")

    def test_row_out_of_range_raises(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        with pytest.raises(ValueError, match="row 4 out of range"):
            get_cell_element(tbl, row=4, col=1)

    def test_col_out_of_range_raises(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        with pytest.raises(ValueError, match="col 4 out of range"):
            get_cell_element(tbl, row=1, col=4)

    def test_row_zero_raises(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        with pytest.raises(ValueError, match="row 0 out of range"):
            get_cell_element(tbl, row=0, col=1)


class TestGetCellParagraphs:
    """Tests for get_cell_paragraphs()."""

    def test_single_paragraph_cell(self, simple_table_path: Path) -> None:
        doc = DocxDocument(path=simple_table_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]
        tc = get_cell_element(tbl, row=1, col=1)

        paras = get_cell_paragraphs(tc)
        assert len(paras) >= 1
        assert all(p.tag.endswith("}p") for p in paras)

    def test_multi_paragraph_cell(self, table_multi_para_path: Path) -> None:
        doc = DocxDocument(path=table_multi_para_path)
        tables = xpath(doc.body, ".//w:tbl")
        tbl = tables[0]

        tc = get_cell_element(tbl, row=2, col=1)
        paras = get_cell_paragraphs(tc)
        assert len(paras) == 2

        tc = get_cell_element(tbl, row=2, col=2)
        paras = get_cell_paragraphs(tc)
        assert len(paras) == 3


class TestParseCellId:
    """Tests for parse_cell_id()."""

    def test_valid_cell_id(self) -> None:
        table_id, row, col = parse_cell_id("2.1.3")
        assert table_id == 2
        assert row == 1
        assert col == 3

    def test_single_digit_components(self) -> None:
        table_id, row, col = parse_cell_id("1.2.3")
        assert table_id == 1
        assert row == 2
        assert col == 3

    def test_multi_digit_components(self) -> None:
        table_id, row, col = parse_cell_id("10.20.30")
        assert table_id == 10
        assert row == 20
        assert col == 30

    def test_missing_component_raises(self) -> None:
        with pytest.raises(ValueError, match="must have format"):
            parse_cell_id("2.1")

    def test_too_many_components_raises(self) -> None:
        with pytest.raises(ValueError, match="must have format"):
            parse_cell_id("2.1.3.4")

    def test_non_numeric_component_raises(self) -> None:
        with pytest.raises(ValueError, match="must be integers"):
            parse_cell_id("2.a.3")

    def test_zero_component_raises(self) -> None:
        with pytest.raises(ValueError, match="must be positive"):
            parse_cell_id("2.0.3")

    def test_negative_component_raises(self) -> None:
        with pytest.raises(ValueError, match="must be positive"):
            parse_cell_id("2.1.-3")
