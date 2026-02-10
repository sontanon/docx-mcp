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
