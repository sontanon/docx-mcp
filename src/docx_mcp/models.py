"""Data models for the docx-mcp redlining engine."""

from __future__ import annotations

from datetime import UTC, datetime
from enum import StrEnum

from pydantic import BaseModel, Field


class ChangeType(StrEnum):
    """Type of change to apply to a document fragment."""

    MODIFY = "modify"
    DELETE = "delete"
    APPEND_AFTER = "append_after"
    MODIFY_CELL = "modify_cell"
    CLEAR_CELL = "clear_cell"


class Change(BaseModel):
    """A single change to apply to a document.

    Attributes:
        fragment_id: 1-based paragraph index identifying the target paragraph.
        change_type: The type of change (modify, delete, or append_after).
        new_text: The new text content in pseudo-Markdown format.
            Required for MODIFY and APPEND_AFTER. None for DELETE.
        justification: Explanation of why this change is being made.
            Used as the comment text in the redlined document.
    """

    fragment_id: int
    change_type: ChangeType
    new_text: str | None = None
    justification: str

    # -- Optional spacing controls -------------------------------------------
    blank_lines_before: int = Field(
        default=0,
        ge=0,
        description=(
            "Number of blank paragraphs to insert before the appended paragraph. "
            "Only used with append_after."
        ),
    )
    blank_lines_after: int = Field(
        default=0,
        ge=0,
        description=(
            "Number of blank paragraphs to insert after the appended paragraph. "
            "Only used with append_after."
        ),
    )
    delete_next_blanks: int = Field(
        default=0,
        ge=0,
        description=(
            "Number of blank paragraphs immediately following the deleted "
            "paragraph to also mark as deleted.  Only used with delete.  "
            "Each targeted paragraph must be blank (whitespace-only), otherwise "
            "an error is raised."
        ),
    )


class DiffOp(StrEnum):
    """Type of operation in a word-level diff."""

    EQUAL = "equal"
    INSERT = "insert"
    DELETE = "delete"


class TableChange(BaseModel):
    """A change to apply to a single table cell.

    Attributes:
        table_id: 1-based table index in document order.
        row: 1-based row index within the table.
        col: 1-based column index within the row.
        cell_id: Dotted cell reference "table_id.row.col" (e.g., "2.1.3").
        change_type: The type of change (modify_cell or clear_cell).
        new_text: New cell text in pseudo-Markdown format. Use \\n for
            paragraph breaks within a cell. Required for modify_cell.
            None for clear_cell.
        justification: Explanation of why this change is being made.
    """

    table_id: int
    row: int
    col: int
    cell_id: str
    change_type: ChangeType
    new_text: str | None = None
    justification: str


class CellInfo(BaseModel):
    """Information about a single table cell.

    Attributes:
        cell_id: Dotted cell reference "table_id.row.col".
        row: 1-based row index.
        col: 1-based column index.
        text: Cell text in pseudo-Markdown format. Multiple paragraphs
            within the cell are joined with \\n.
    """

    cell_id: str
    row: int
    col: int
    text: str

    model_config = {"frozen": True}


class TableInfo(BaseModel):
    """Information about a simple (rectangular) table.

    Attributes:
        table_id: 1-based table index in document order.
        rows: Number of rows in the table.
        cols: Number of columns in the table.
        cells: 2D list of CellInfo, indexed as cells[row][col].
    """

    table_id: int
    rows: int
    cols: int
    cells: list[list[CellInfo]]

    model_config = {"frozen": True}


class SkippedTableInfo(BaseModel):
    """Information about a table that was skipped during extraction.

    Attributes:
        table_id: 1-based table index in document order.
        skipped: Always True (for discriminated union detection).
        reason: Human-readable explanation of why the table was skipped.
    """

    table_id: int
    skipped: bool = True
    reason: str

    model_config = {"frozen": True}


class DiffChunk(BaseModel):
    """A single chunk in a word-level diff result.

    Attributes:
        op: The type of operation (equal, insert, or delete).
        text: The text content of this chunk (words joined by spaces).
    """

    op: DiffOp
    text: str

    model_config = {"frozen": True}


class RedlineConfig(BaseModel):
    """Configuration for redline generation.

    Attributes:
        author: The author name to use for tracked changes and comments.
        date: The timestamp to use. Defaults to the current UTC time.
    """

    author: str = "AI Review"
    date: datetime = Field(default_factory=lambda: datetime.now(UTC))

    def date_iso(self) -> str:
        """Return the date in ISO 8601 format for OOXML attributes."""
        return self.date.strftime("%Y-%m-%dT%H:%M:%SZ")
