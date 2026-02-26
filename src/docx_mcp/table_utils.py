"""Table inspection and cell resolution utilities.

Provides functions to:
- Check if a table is supported (horizontal merges OK, vertical merges not).
- Get table dimensions (rows, grid columns).
- Resolve cell references to XML elements via grid-based addressing.
- Parse cell_id strings.
"""

from __future__ import annotations

from lxml import etree

from docx_mcp.namespaces import qn, xpath


def get_grid_span(cell: etree._Element) -> int:
    """Return the ``w:gridSpan`` value for *cell*, defaulting to 1.

    Args:
        cell: A ``<w:tc>`` element.

    Returns:
        The number of grid columns this cell spans. Returns 1 if
        ``<w:gridSpan>`` is not present.
    """
    tcPr_list = xpath(cell, "./w:tcPr")
    if not tcPr_list:
        return 1
    gridSpan_list = xpath(tcPr_list[0], "./w:gridSpan")
    if not gridSpan_list:
        return 1
    val = gridSpan_list[0].get(qn("w", "val"))
    return int(val) if val else 1


def get_grid_col_count(tbl: etree._Element) -> int:
    """Return the number of grid columns from ``<w:tblGrid>``.

    Args:
        tbl: A ``<w:tbl>`` element.

    Returns:
        The number of ``<w:gridCol>`` elements, or 0 if not found.
    """
    gridCols = xpath(tbl, "./w:tblGrid/w:gridCol")
    return len(gridCols)


def get_cell_at_grid(tbl: etree._Element, row: int, grid_col: int) -> etree._Element:
    """Return the ``<w:tc>`` element covering a grid column position.

    Uses grid-based addressing: a cell spanning grid columns 1-2 can be
    addressed by either grid_col=1 or grid_col=2.

    Args:
        tbl: A ``<w:tbl>`` element.
        row: 1-based row index.
        grid_col: 1-based grid column position.

    Returns:
        The ``<w:tc>`` element covering the specified grid column.

    Raises:
        ValueError: If row or grid_col is out of range.
    """
    rows = list(xpath(tbl, "./w:tr"))

    if row < 1 or row > len(rows):
        msg = f"row {row} out of range (table has {len(rows)} rows)"
        raise ValueError(msg)

    row_el = rows[row - 1]
    current_col = 1

    for cell in xpath(row_el, "./w:tc"):
        span = get_grid_span(cell)
        if grid_col >= current_col and grid_col < current_col + span:
            return cell
        current_col += span

    msg = f"grid column {grid_col} out of range (row {row} has {current_col - 1} grid columns)"
    raise ValueError(msg)


def is_simple_table(tbl: etree._Element) -> tuple[bool, str]:
    """Check if *tbl* is a supported table structure.

    A supported table has:
    - No ``<w:vMerge>`` (vertical merges).
    - No nested ``<w:tbl>`` elements.
    - All rows span the same number of grid columns.

    Horizontal merges (``<w:gridSpan>``) are allowed.

    Args:
        tbl: A ``<w:tbl>`` element.

    Returns:
        A tuple of (is_supported, reason). If supported, returns (True, "").
        If not supported, returns (False, "reason message").
    """
    rows = list(xpath(tbl, "./w:tr"))
    if not rows:
        return False, "table has no rows"

    expected_grid_cols: int | None = None

    for row_idx, row in enumerate(rows, start=1):
        cells = list(xpath(row, "./w:tc"))
        grid_cols_in_row = 0

        for col_idx, cell in enumerate(cells, start=1):
            # Check for vertical merge
            tcPr_list = xpath(cell, "./w:tcPr")
            if tcPr_list:
                tcPr = tcPr_list[0]
                if xpath(tcPr, "./w:vMerge"):
                    msg = f"cell {row_idx}.{col_idx} has vertical merge (vMerge)"
                    return False, msg

            # Check for nested table
            nested_tables = xpath(cell, ".//w:tbl")
            if nested_tables:
                msg = f"cell {row_idx}.{col_idx} contains nested table"
                return False, msg

            # Accumulate grid columns
            grid_cols_in_row += get_grid_span(cell)

        if expected_grid_cols is None:
            expected_grid_cols = grid_cols_in_row
        elif grid_cols_in_row != expected_grid_cols:
            msg = (
                f"row {row_idx} spans {grid_cols_in_row} grid columns, "
                f"expected {expected_grid_cols}"
            )
            return False, msg

    return True, ""


def table_dimensions(tbl: etree._Element) -> tuple[int, int]:
    """Return the (rows, grid_cols) dimensions of *tbl*.

    Uses grid columns (from ``<w:tblGrid>``) rather than cell count,
    so horizontally merged cells are properly represented.

    Args:
        tbl: A ``<w:tbl>`` element.

    Returns:
        A tuple of (row_count, grid_column_count).
    """
    rows = list(xpath(tbl, "./w:tr"))
    row_count = len(rows)

    if not rows:
        return (0, 0)

    # Get grid column count from tblGrid if available
    grid_col_count = get_grid_col_count(tbl)
    if grid_col_count > 0:
        return (row_count, grid_col_count)

    # Fallback: compute from first row cells
    first_row_cells = list(xpath(rows[0], "./w:tc"))
    col_count = sum(get_grid_span(cell) for cell in first_row_cells)
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
