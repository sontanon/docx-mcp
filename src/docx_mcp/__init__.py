"""docx-mcp: Legal document redlining engine.

Apply AI-generated changes to .docx files as professional tracked changes
with comments, indistinguishable from a lawyer's Word redlines.

Basic usage::

    from docx_mcp import apply_redlines, Change, ChangeType, RedlineConfig

    changes = [
        Change(
            fragment_id=3,
            change_type=ChangeType.MODIFY,
            new_text="The Company **shall** provide written notice.",
            justification="Strengthened obligation language.",
        ),
    ]
    doc = apply_redlines("contract.docx", changes)
    doc.save("contract_redlined.docx")
"""

__version__ = "0.1.0"

from docx_mcp.converter import (
    document_to_fragments,
    fragments_to_tagged_text,
    paragraph_to_pseudo_markdown,
    pseudo_markdown_to_raw,
)
from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType, DiffChunk, DiffOp, RedlineConfig
from docx_mcp.redliner import apply_redlines
from docx_mcp.validator import ValidationResult, validate_document

__all__ = [
    "Change",
    "ChangeType",
    "DiffChunk",
    "DiffOp",
    "DocxDocument",
    "RedlineConfig",
    "ValidationResult",
    "apply_redlines",
    "document_to_fragments",
    "fragments_to_tagged_text",
    "paragraph_to_pseudo_markdown",
    "pseudo_markdown_to_raw",
    "validate_document",
]
