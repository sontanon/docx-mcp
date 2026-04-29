"""Table inspection and cell resolution utilities.

Provides functions to:
- Check if a table is "simple" (no merged cells, no nested tables).
- Get table dimensions (rows, cols).
- Resolve cell references to XML elements.
- Parse cell_id strings.
"""

from __future__ import annotations

from lxml import etree

from docx_mcp.namespaces import xpath


def is_simple_table(
    tbl: etree._Element,
    *,
    table_id: int | None = None,
) -> tuple[bool, str]:
    """Check if *tbl* is a simple rectangular table without merged cells.

    A "simple" table has:
    - No ``<w:gridSpan>`` (horizontal merges).
    - No ``<w:vMerge>`` (vertical merges).
    - No nested ``<w:tbl>`` elements.
    - All rows have the same number of cells.

    Args:
        tbl: A ``<w:tbl>`` element.
        table_id: Optional table fragment ID to include in error messages.

    Returns:
        A tuple of (is_simple, reason). If simple, returns (True, "").
        If not simple, returns (False, "reason message").
    """
    prefix = f"table {table_id}, " if table_id is not None else ""

    rows = list(xpath(tbl, "./w:tr"))
    if not rows:
        return False, f"{prefix}table has no rows"

    expected_cells: int | None = None

    for row_idx, row in enumerate(rows, start=1):
        cells = list(xpath(row, "./w:tc"))

        if expected_cells is None:
            expected_cells = len(cells)
        elif len(cells) != expected_cells:
            msg = f"{prefix}row {row_idx} has {len(cells)} cells, expected {expected_cells}"
            return False, msg

        for col_idx, cell in enumerate(cells, start=1):
            tcPr_list = xpath(cell, "./w:tcPr")
            if tcPr_list:
                tcPr = tcPr_list[0]

                if xpath(tcPr, "./w:gridSpan"):
                    msg = f"{prefix}cell {row_idx}.{col_idx} has horizontal merge (gridSpan)"
                    return False, msg

                if xpath(tcPr, "./w:vMerge"):
                    msg = f"{prefix}cell {row_idx}.{col_idx} has vertical merge (vMerge)"
                    return False, msg

            nested_tables = xpath(cell, ".//w:tbl")
            if nested_tables:
                msg = f"{prefix}cell {row_idx}.{col_idx} contains nested table"
                return False, msg

    return True, ""


def table_dimensions(tbl: etree._Element) -> tuple[int, int]:
    """Return the (rows, cols) dimensions of *tbl*.

    Assumes the table has been validated as simple (uniform row lengths).
    If called on an invalid table, returns the count from the first row.

    Args:
        tbl: A ``<w:tbl>`` element.

    Returns:
        A tuple of (row_count, column_count).
    """
    rows = list(xpath(tbl, "./w:tr"))
    row_count = len(rows)

    if not rows:
        return (0, 0)

    first_row_cells = list(xpath(rows[0], "./w:tc"))
    col_count = len(first_row_cells)

    return (row_count, col_count)


def get_cell_element(tbl: etree._Element, row: int, col: int) -> etree._Element:
    """Return the ``<w:tc>`` element at 1-based (row, col) position.

    Args:
        tbl: A ``<w:tbl>`` element.
        row: 1-based row index.
        col: 1-based column index.

    Returns:
        The ``<w:tc>`` element at the specified position.

    Raises:
        ValueError: If row or col is out of range.
    """
    rows = list(xpath(tbl, "./w:tr"))

    if row < 1 or row > len(rows):
        msg = f"row {row} out of range (table has {len(rows)} rows)"
        raise ValueError(msg)

    row_el = rows[row - 1]
    cells = list(xpath(row_el, "./w:tc"))

    if col < 1 or col > len(cells):
        msg = f"col {col} out of range (row {row} has {len(cells)} cells)"
        raise ValueError(msg)

    return cells[col - 1]


def get_cell_paragraphs(cell: etree._Element) -> list[etree._Element]:
    """Return all ``<w:p>`` children of *cell*.

    Args:
        cell: A ``<w:tc>`` element.

    Returns:
        List of ``<w:p>`` elements that are direct children of the cell.
    """
    return list(xpath(cell, "./w:p"))


def parse_cell_id(cell_id: str) -> tuple[int, int, int]:
    """Parse a dotted cell reference into (table_id, row, col).

    Args:
        cell_id: Dotted cell reference like "2.1.3".

    Returns:
        A tuple of (table_id, row, col), all 1-based integers.

    Raises:
        ValueError: If the format is invalid or values are not positive integers.
    """
    parts = cell_id.split(".")

    if len(parts) != 3:
        msg = f"cell_id '{cell_id}' must have format 'table_id.row.col'"
        raise ValueError(msg)

    try:
        table_id = int(parts[0])
        row = int(parts[1])
        col = int(parts[2])
    except ValueError:
        msg = f"cell_id '{cell_id}' components must be integers"
        raise ValueError(msg) from None

    if table_id < 1 or row < 1 or col < 1:
        msg = f"cell_id '{cell_id}' components must be positive integers"
        raise ValueError(msg)

    return (table_id, row, col)
