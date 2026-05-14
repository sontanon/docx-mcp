"""Table inspection and cell resolution utilities.

Provides functions to:
- Check if a table is "simple" (no merged cells, no nested tables).
- Build a logical grid for tables with merged cells.
- Get table dimensions (rows, cols).
- Resolve cell references to XML elements.
- Parse cell_id strings.
"""

import contextlib
from dataclasses import dataclass

from lxml import etree

from docx_mcp.namespaces import qn, xpath


@dataclass
class GridCell:
    """A cell in the logical table grid."""

    element: etree._Element
    row: int
    col: int
    span: int = 1
    vspan: int = 1
    is_spanned_over: bool = False


def _get_grid_span(cell: etree._Element) -> int:
    """Read gridSpan from a <w:tc> element (default 1)."""
    tcPr_list = xpath(cell, "./w:tcPr")
    if not tcPr_list:
        return 1
    tcPr = tcPr_list[0]
    grid_span_els = xpath(tcPr, "./w:gridSpan")
    if not grid_span_els:
        return 1
    val = grid_span_els[0].get(qn("w", "val"))
    if val is None:
        return 1
    with contextlib.suppress(ValueError):
        return int(val)
    return 1


def _has_vmerge(cell: etree._Element) -> bool:
    """Check if cell has a vMerge element."""
    tcPr_list = xpath(cell, "./w:tcPr")
    if not tcPr_list:
        return False
    return bool(xpath(tcPr_list[0], "./w:vMerge"))


def _is_vmerge_restart(cell: etree._Element) -> bool:
    """Check if cell has vMerge with val='restart'."""
    tcPr_list = xpath(cell, "./w:tcPr")
    if not tcPr_list:
        return False
    vmerge_els = xpath(tcPr_list[0], "./w:vMerge")
    if not vmerge_els:
        return False
    return vmerge_els[0].get(qn("w", "val")) == "restart"


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


def build_table_grid(
    tbl: etree._Element,
    table_id: int,
) -> tuple[list[list[GridCell | None]], str | None]:
    """Build a logical grid from a table, handling horizontal merges.

    Args:
        tbl: A ``<w:tbl>`` element.
        table_id: Table ID for error messages.

    Returns:
        Tuple of (grid, error). If successful, grid is a 2D list of GridCell
        and error is None. If the table is too complex, grid is empty and
        error is a reason string.
    """

    rows = list(xpath(tbl, "./w:tr"))
    if not rows:
        return [], f"table {table_id}, table has no rows"

    # Get total grid columns from tblGrid
    tbl_grid = xpath(tbl, "./w:tblGrid")
    total_cols = len(xpath(tbl_grid[0], "./w:gridCol")) if tbl_grid else 0

    grid: list[list[GridCell | None]] = []

    for row_idx, row in enumerate(rows, start=1):
        grid_row: list[GridCell | None] = [None] * total_cols
        cells = list(xpath(row, "./w:tc"))
        col_idx = 0

        for cell in cells:
            # Find next available column
            while col_idx < total_cols and grid_row[col_idx] is not None:
                col_idx += 1

            if col_idx >= total_cols:
                return [], f"table {table_id}, row {row_idx} has too many cells"

            span = _get_grid_span(cell)

            # Check for vertical merge
            vspan = 1
            is_spanned_over = False
            if _has_vmerge(cell):
                if _is_vmerge_restart(cell):
                    # Count continuation rows
                    vspan = _count_vmerge_span(tbl, row_idx, col_idx)
                else:
                    # This is a continuation cell — skip it, the parent handles it
                    is_spanned_over = True

            # Place the cell
            gc = GridCell(
                element=cell,
                row=row_idx,
                col=col_idx + 1,  # 1-based
                span=span,
                vspan=vspan,
                is_spanned_over=is_spanned_over,
            )
            grid_row[col_idx] = gc

            # Mark spanned-over horizontal positions
            for s in range(1, span):
                if col_idx + s < total_cols:
                    grid_row[col_idx + s] = GridCell(
                        element=cell,
                        row=row_idx,
                        col=col_idx + s + 1,
                        span=0,
                        vspan=0,
                        is_spanned_over=True,
                    )

            col_idx += span

        # Fill remaining columns with spanned-over markers (from above-row vMerge)
        for c in range(total_cols):
            if (
                grid_row[c] is None
                and row_idx > 1
                and len(grid) >= row_idx - 1
            ):
                above = grid[row_idx - 2][c]
                if above is not None and above.vspan > 1:
                        # This position is vertically spanned
                        remaining_vspan = above.vspan - (row_idx - above.row)
                        if remaining_vspan > 0:
                            grid_row[c] = GridCell(
                                element=above.element,
                                row=row_idx,
                                col=c + 1,
                                span=0,
                                vspan=0,
                                is_spanned_over=True,
                            )

        grid.append(grid_row)

    # Check for nested tables
    for row_idx, row in enumerate(grid, start=1):
        for col_idx, cell in enumerate(row, start=1):
            if cell is not None and not cell.is_spanned_over:
                nested = xpath(cell.element, ".//w:tbl")
                if nested:
                    return [], f"table {table_id}, cell {row_idx}.{col_idx} contains nested table"

    return grid, None


def _count_vmerge_span(tbl: etree._Element, start_row: int, start_col: int) -> int:
    """Count how many rows a vMerge spans starting at start_row (1-based)."""
    rows = list(xpath(tbl, "./w:tr"))
    count = 1
    for r in range(start_row, len(rows)):
        next_row = rows[r]
        cells = list(xpath(next_row, "./w:tc"))
        col_idx = 0
        for cell in cells:
            span = _get_grid_span(cell)
            if col_idx == start_col:
                if _has_vmerge(cell) and not _is_vmerge_restart(cell):
                    count += 1
                else:
                    return count
            col_idx += span
            if col_idx > start_col:
                return count
    return count


def table_dimensions(tbl: etree._Element) -> tuple[int, int]:
    """Return the (rows, cols) dimensions of *tbl*.

    Uses ``<w:tblGrid>`` for column count if available, otherwise
    falls back to the first row's cell count.

    Args:
        tbl: A ``<w:tbl>`` element.

    Returns:
        A tuple of (row_count, column_count).
    """
    rows = list(xpath(tbl, "./w:tr"))
    row_count = len(rows)

    tbl_grid = xpath(tbl, "./w:tblGrid")
    if tbl_grid:
        col_count = len(xpath(tbl_grid[0], "./w:gridCol"))
    elif rows:
        first_row_cells = list(xpath(rows[0], "./w:tc"))
        col_count = len(first_row_cells)
    else:
        col_count = 0

    return (row_count, col_count)


def get_cell_element(tbl: etree._Element, row: int, col: int) -> etree._Element:
    """Return the ``<w:tc>`` element at 1-based (row, col) position.

    For tables with merged cells, this returns the *starting* cell that
    covers the requested position.

    Args:
        tbl: A ``<w:tbl>`` element.
        row: 1-based row index.
        col: 1-based column index.

    Returns:
        The ``<w:tc>`` element at the specified position.

    Raises:
        ValueError: If row or col is out of range.
    """
    grid, error = build_table_grid(tbl, table_id=0)
    if error:
        msg = f"Cannot resolve cell: {error}"
        raise ValueError(msg)

    if row < 1 or row > len(grid):
        msg = f"row {row} out of range (table has {len(grid)} rows)"
        raise ValueError(msg)

    grid_row = grid[row - 1]
    if col < 1 or col > len(grid_row):
        msg = f"col {col} out of range (row {row} has {len(grid_row)} cells)"
        raise ValueError(msg)

    cell = grid_row[col - 1]
    if cell is None:
        msg = f"No cell at row {row}, col {col}"
        raise ValueError(msg)

    if cell.is_spanned_over:
        # Find the starting cell position
        start_col = cell.col
        for c in range(col - 1, -1, -1):
            prev = grid_row[c]
            if prev is not None and not prev.is_spanned_over:
                start_col = c + 1
                break
        msg = (
            f"Cell at row {row}, col {col} is spanned over by another cell. "
            f"Target the starting cell (row {row}, col {start_col}) instead."
        )
        raise ValueError(msg)

    return cell.element


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
