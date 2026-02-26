"""Data models for the docx-mcp redlining engine.

This module defines the core domain models for representing document changes.
The architecture uses a discriminated union pattern to maintain type safety
internally while allowing flexible APIs at the interface layer.

Core Models:
    ParagraphChange: Changes to document paragraphs (modify/delete/append).
    TableChange: Changes to table cells (modify/clear).
    Change: Discriminated union of ParagraphChange | TableChange.

The explicit `kind` field enables Pydantic's discriminated union validation,
ensuring type-safe processing throughout the pipeline.
"""

from __future__ import annotations

from datetime import UTC, datetime
from enum import StrEnum
from typing import Annotated, Literal

from pydantic import BaseModel, Discriminator, Field, Tag, model_validator


class ParagraphChangeType(StrEnum):
    """Type of change to apply to a paragraph."""

    MODIFY = "modify"
    DELETE = "delete"
    APPEND_AFTER = "append_after"


class TableChangeType(StrEnum):
    """Type of change to apply to a table cell."""

    MODIFY_CELL = "modify_cell"
    CLEAR_CELL = "clear_cell"


class ParagraphChange(BaseModel):
    """A single change to apply to a document paragraph.

    Represents one modification, deletion, or insertion operation targeting a
    specific paragraph. Fragment IDs are 1-based indices corresponding to
    top-level ``<w:p>`` elements in ``<w:body>`` in document order.

    The explicit `kind` field enables discriminated union validation and
    type-safe processing.

    Validation Rules:
        - `new_text` is required for MODIFY and APPEND_AFTER; must be None for DELETE
        - `blank_lines_before`/`blank_lines_after` only valid with APPEND_AFTER
        - `delete_next_blanks` only valid with DELETE

    Attributes:
        kind: Discriminator field, always "paragraph".
        fragment_id: 1-based paragraph index identifying the target paragraph.
        change_type: The type of change (modify, delete, or append_after).
        new_text: The new text content in pseudo-Markdown format.
            Required for MODIFY and APPEND_AFTER. None for DELETE.
        justification: Explanation of why this change is being made.
            Used as the comment text in the redlined document.
        blank_lines_before: Number of blank paragraphs to insert before the
            appended paragraph. Only used with APPEND_AFTER.
        blank_lines_after: Number of blank paragraphs to insert after the
            appended paragraph. Only used with APPEND_AFTER.
        delete_next_blanks: Number of blank paragraphs immediately following
            the deleted paragraph to also mark as deleted. Only used with DELETE.
    """

    kind: Literal["paragraph"] = "paragraph"
    fragment_id: int
    change_type: ParagraphChangeType
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

    @model_validator(mode="after")
    def validate_change_type_constraints(self) -> ParagraphChange:
        """Validate that fields are consistent with change_type."""
        if self.change_type == ParagraphChangeType.MODIFY:
            if self.new_text is None:
                msg = "new_text is required for MODIFY changes"
                raise ValueError(msg)
            if self.blank_lines_before != 0 or self.blank_lines_after != 0:
                msg = "blank_lines_before/blank_lines_after only allowed for APPEND_AFTER"
                raise ValueError(msg)
            if self.delete_next_blanks != 0:
                msg = "delete_next_blanks only allowed for DELETE"
                raise ValueError(msg)

        elif self.change_type == ParagraphChangeType.DELETE:
            if self.new_text is not None:
                msg = "new_text must be None for DELETE changes"
                raise ValueError(msg)
            if self.blank_lines_before != 0 or self.blank_lines_after != 0:
                msg = "blank_lines_before/blank_lines_after only allowed for APPEND_AFTER"
                raise ValueError(msg)

        elif self.change_type == ParagraphChangeType.APPEND_AFTER:
            if self.new_text is None:
                msg = "new_text is required for APPEND_AFTER changes"
                raise ValueError(msg)
            if self.delete_next_blanks != 0:
                msg = "delete_next_blanks only allowed for DELETE"
                raise ValueError(msg)

        return self


class TableChange(BaseModel):
    """A change to apply to a single table cell.

    Targets a specific cell within a table using numeric coordinates.
    The `cell_id` property provides a computed dotted reference for display.

    The explicit `kind` field enables discriminated union validation and
    type-safe processing.

    Validation Rules:
        - `new_text` is required for MODIFY_CELL; must be None for CLEAR_CELL

    Attributes:
        kind: Discriminator field, always "table".
        table_id: 1-based table index in document order.
        row: 1-based row index within the table.
        col: 1-based column index within the row.
        change_type: The type of change (modify_cell or clear_cell).
        new_text: New cell text in pseudo-Markdown format. Use \n for
            paragraph breaks within a cell. Required for MODIFY_CELL.
            None for CLEAR_CELL.
        justification: Explanation of why this change is being made.
    """

    kind: Literal["table"] = "table"
    table_id: int
    row: int
    col: int
    change_type: TableChangeType
    new_text: str | None = None
    justification: str

    @property
    def cell_id(self) -> str:
        """Computed cell ID in format 'table_id.row.col'."""
        return f"{self.table_id}.{self.row}.{self.col}"

    @model_validator(mode="after")
    def validate_change_type_constraints(self) -> TableChange:
        """Validate that fields are consistent with change_type."""
        if self.change_type == TableChangeType.MODIFY_CELL:
            if self.new_text is None:
                msg = "new_text is required for MODIFY_CELL changes"
                raise ValueError(msg)

        elif self.change_type == TableChangeType.CLEAR_CELL and self.new_text is not None:
            msg = "new_text must be None for CLEAR_CELL changes"
            raise ValueError(msg)

        return self


def _discriminate_change(v: dict | BaseModel) -> str:
    """Discriminator function for Change union based on explicit 'kind' field."""
    kind = v.get("kind", "") if isinstance(v, dict) else getattr(v, "kind", "")
    if kind in ("paragraph", "table"):
        return kind
    msg = f"Invalid or missing 'kind' field. Expected 'paragraph' or 'table', got: {kind}"
    raise ValueError(msg)


# Discriminated union type for all changes
Change = Annotated[
    Annotated[ParagraphChange, Tag("paragraph")] | Annotated[TableChange, Tag("table")],
    Discriminator(_discriminate_change),
]


class CellInfo(BaseModel):
    """Information about a single table cell.

    Attributes:
        cell_id: Dotted cell reference "table_id.row.col".
        row: 1-based row index.
        col: 1-based grid column index (starting position).
        text: Cell text in pseudo-Markdown format. Multiple paragraphs
            within the cell are joined with \n.
        grid_span: Number of grid columns this cell spans. Default is 1.
            For horizontally merged cells, this is the value of w:gridSpan.
    """

    cell_id: str
    row: int
    col: int
    text: str
    grid_span: int = 1

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


class DiffOp(StrEnum):
    """Type of operation in a word-level diff."""

    EQUAL = "equal"
    INSERT = "insert"
    DELETE = "delete"


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
