"""docx-mcp: Legal document redlining engine.

Apply AI-generated changes to .docx files as professional tracked changes
with comments, indistinguishable from a lawyer's Word redlines.

Basic usage::

    from docx_mcp import (
        apply_redlines,
        ParagraphChange,
        ParagraphChangeType,
        RedlineConfig,
    )

    changes = [
        ParagraphChange(
            kind="paragraph",
            fragment_id="3",
            change_type=ParagraphChangeType.MODIFY,
            new_text="The Company **shall** provide written notice.",
            justification="Strengthened obligation language.",
        ),
    ]
    doc = apply_redlines("contract.docx", changes)
    doc.save("contract_redlined.docx")
"""

__version__ = "0.2.0"

from docx_mcp.audit import AuditReport, audit_document
from docx_mcp.converter import (
    FragmentResult,
    body_to_fragments,
    document_to_fragments,
    fragments_to_json_interleaved,
    fragments_to_tagged_text,
    fragments_to_tagged_text_interleaved,
    full_to_fragments,
    paragraph_to_pseudo_markdown,
    pseudo_markdown_to_raw,
)
from docx_mcp.document import DocxDocument
from docx_mcp.models import (
    CellInfo,
    Change,
    DiffChunk,
    DiffOp,
    ParagraphChange,
    ParagraphChangeType,
    RedlineConfig,
    SkippedTableInfo,
    TableChange,
    TableChangeType,
    TableInfo,
)
from docx_mcp.redliner import apply_redlines
from docx_mcp.validator import ValidationResult, validate_document

__all__ = [
    "AuditReport",
    "CellInfo",
    "Change",
    "DiffChunk",
    "DiffOp",
    "DocxDocument",
    "FragmentResult",
    "ParagraphChange",
    "ParagraphChangeType",
    "RedlineConfig",
    "SkippedTableInfo",
    "TableChange",
    "TableChangeType",
    "TableInfo",
    "ValidationResult",
    "apply_redlines",
    "audit_document",
    "body_to_fragments",
    "document_to_fragments",
    "fragments_to_json_interleaved",
    "fragments_to_tagged_text",
    "fragments_to_tagged_text_interleaved",
    "full_to_fragments",
    "paragraph_to_pseudo_markdown",
    "pseudo_markdown_to_raw",
    "validate_document",
]
