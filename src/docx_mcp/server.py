"""MCP server for the docx-mcp legal-document redlining engine.

Exposes the redlining pipeline as MCP tools so that LLM clients (Claude Desktop,
Cursor, ChatGPT, etc.) can extract document fragments, apply tracked changes,
validate output, and compare document versions -- all via filesystem paths.

Tools:
    extract_fragments   Read a .docx and return paragraph text.
    apply_changes       Apply tracked changes (inline list).
    apply_changes_from_file  Apply tracked changes from a JSON file.
    validate_document   Structural validation of a .docx file.
    diff_fragments      Compare two .docx files paragraph-by-paragraph.

Resource:
    docx://{document_path}/fragments  Browse document fragments.

Transport:
    stdio (default, for local CLI integration).

Usage::

    # Start the server (stdio)
    docx-mcp-server

    # Or programmatically
    from docx_mcp.server import mcp
    mcp.run(transport="stdio")
"""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Literal

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from pydantic import BaseModel, Field, TypeAdapter, ValidationError

from docx_mcp.converter import document_to_fragments, fragments_to_tagged_text
from docx_mcp.differ import DmpWordDiffer
from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType, DiffOp, RedlineConfig
from docx_mcp.redliner import apply_redlines
from docx_mcp.validator import validate_document as _validate_document

# ---------------------------------------------------------------------------
# MCP server instance
# ---------------------------------------------------------------------------

mcp = FastMCP(
    "docx-mcp",
    instructions=(
        "Legal document redlining server. Use extract_fragments to read a "
        ".docx file, then apply_changes to add tracked changes with comments. "
        "Fragment IDs are 1-based paragraph indices. Text uses pseudo-Markdown: "
        "**bold**, _italic_, __underline__."
    ),
)

# ---------------------------------------------------------------------------
# Pydantic model for tool parameters
# ---------------------------------------------------------------------------


class ChangeParam(BaseModel):
    """A single tracked change to apply to a document paragraph.

    Attributes:
        fragment_id: 1-based paragraph index identifying the target.
        change_type: ``"modify"`` to edit in-place, ``"delete"`` to remove,
            or ``"append_after"`` to insert a new paragraph after this one.
        new_text: Replacement text in pseudo-Markdown (``**bold**``,
            ``_italic_``, ``__underline__``).  Required for modify and
            append_after; omit for delete.
        justification: Human-readable reason for the change.  This becomes
            a Word comment attached to the tracked change.
    """

    fragment_id: int = Field(description="1-based paragraph index")
    change_type: Literal["modify", "delete", "append_after"] = Field(
        description='Type of change: "modify", "delete", or "append_after"',
    )
    new_text: str | None = Field(
        default=None,
        description=(
            "New text in pseudo-Markdown format (**bold**, _italic_, __underline__). "
            "Required for modify and append_after. Omit for delete."
        ),
    )
    justification: str = Field(
        description="Reason for this change. Becomes a Word comment.",
    )


# TypeAdapter for validating changes loaded from a JSON file.
_CHANGES_ADAPTER: TypeAdapter[list[ChangeParam]] = TypeAdapter(list[ChangeParam])

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_document(document_path: str) -> DocxDocument:
    """Load a ``.docx`` file, raising ``ToolError`` on failure."""
    path = Path(document_path)
    if not path.exists():
        msg = f"File not found: {document_path}"
        raise ToolError(msg)
    if not path.is_file():
        msg = f"Not a file: {document_path}"
        raise ToolError(msg)
    try:
        return DocxDocument(path=path)
    except Exception as exc:
        msg = f"Failed to open .docx file: {document_path} ({exc})"
        raise ToolError(msg) from exc


def _resolve_output_path(document_path: str, output_path: str | None) -> Path:
    """Determine the output path, defaulting to ``<stem>_redlined.docx``."""
    if output_path is not None:
        return Path(output_path)
    p = Path(document_path)
    return p.parent / f"{p.stem}_redlined.docx"


def _convert_changes(params: list[ChangeParam]) -> list[Change]:
    """Convert ``ChangeParam`` tool inputs to core ``Change`` models."""
    return [
        Change(
            fragment_id=p.fragment_id,
            change_type=ChangeType(p.change_type),
            new_text=p.new_text,
            justification=p.justification,
        )
        for p in params
    ]


def _format_validation(errors: list[str], warnings: list[str]) -> str:
    """Format validation results as human-readable text."""
    parts: list[str] = []
    if not errors and not warnings:
        parts.append("Validation: passed (0 errors, 0 warnings).")
    elif not errors:
        parts.append(f"Validation: passed (0 errors, {len(warnings)} warning(s)).")
        for i, w in enumerate(warnings, 1):
            parts.append(f"  Warning {i}: {w}")
    else:
        parts.append(f"Validation: FAILED ({len(errors)} error(s), {len(warnings)} warning(s)).")
        for i, e in enumerate(errors, 1):
            parts.append(f"  Error {i}: {e}")
        for i, w in enumerate(warnings, 1):
            parts.append(f"  Warning {i}: {w}")
    return "\n".join(parts)


def _apply_and_save(
    document_path: str,
    change_params: list[ChangeParam],
    output_path: str | None,
    author: str,
    validate: bool,
) -> str:
    """Shared implementation for both apply tools."""
    changes = _convert_changes(change_params)
    config = RedlineConfig(author=author)
    out = _resolve_output_path(document_path, output_path)

    try:
        doc = apply_redlines(document_path, changes, config=config)
    except FileNotFoundError as exc:
        msg = f"File not found: {document_path}"
        raise ToolError(msg) from exc
    except ValueError as exc:
        raise ToolError(str(exc)) from exc

    # Build summary
    counts = Counter(c.change_type for c in change_params)
    summary_parts = []
    for ct in ("modify", "delete", "append_after"):
        if counts[ct]:
            summary_parts.append(f"{counts[ct]} {ct}")
    change_summary = ", ".join(summary_parts)

    lines = [
        f"Applied {len(changes)} change(s) ({change_summary}) to {Path(document_path).name}.",
    ]

    # Validate if requested
    if validate:
        result = _validate_document(doc)
        lines.append(f"Output saved to {out}.")
        lines.append(_format_validation(result.errors, result.warnings))
    else:
        lines.append(f"Output saved to {out}.")

    doc.save(out)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------


@mcp.tool(
    annotations={
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def extract_fragments(
    document_path: str,
    format: Literal["tagged", "json"] = "tagged",
) -> str:
    """Read a .docx file and return its paragraphs as text.

    Each paragraph is identified by a 1-based fragment ID. Use these IDs
    when constructing changes for ``apply_changes``.

    The **tagged** format (default) wraps each paragraph::

        <f=1>The Seller shall deliver the goods.</f=1>
        <f=2>The Buyer shall pay within 30 days.</f=2>

    The **json** format returns a JSON array::

        [{"fragment_id": 1, "text": "The Seller shall deliver the goods."}]

    Text uses pseudo-Markdown: **bold**, _italic_, __underline__.

    Args:
        document_path: Path to the .docx file.
        format: Output format -- "tagged" (default) or "json".

    Returns:
        Fragment text in the requested format.
    """
    doc = _load_document(document_path)
    fragments = document_to_fragments(doc.paragraphs)

    if format == "json":
        data = [{"fragment_id": fid, "text": text} for fid, text in fragments]
        return json.dumps(data, ensure_ascii=False, indent=2)

    return fragments_to_tagged_text(fragments)


@mcp.tool(
    annotations={
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def apply_changes(
    document_path: str,
    changes: list[ChangeParam],
    output_path: str | None = None,
    author: str = "AI Review",
    validate: bool = True,
) -> str:
    """Apply tracked changes to a .docx file and save the result.

    Produces a Word document with professional tracked changes (``w:ins`` /
    ``w:del``) and comments, indistinguishable from a lawyer's redline.

    Each change targets a paragraph by its fragment ID (from
    ``extract_fragments``).  Three change types are supported:

    - **modify**: Replace the paragraph text (word-level diff produces
      fine-grained insertions and deletions).
    - **delete**: Mark the entire paragraph as deleted.
    - **append_after**: Insert a new paragraph after the target.

    The original file is never overwritten -- output defaults to
    ``<name>_redlined.docx`` beside the input.

    Args:
        document_path: Path to the input .docx file.
        changes: List of changes to apply.
        output_path: Where to save.  Defaults to ``<stem>_redlined.docx``.
        author: Author name for tracked changes and comments.
        validate: Run structural validation after applying (default True).

    Returns:
        Summary with change counts, output path, and validation result.
    """
    return _apply_and_save(document_path, changes, output_path, author, validate)


@mcp.tool(
    annotations={
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def apply_changes_from_file(
    document_path: str,
    changes_file: str,
    output_path: str | None = None,
    author: str = "AI Review",
    validate: bool = True,
) -> str:
    """Apply tracked changes from a JSON file to a .docx file.

    Same behaviour as ``apply_changes`` but reads the change list from a
    JSON file on disk.  Useful for large change sets or pre-prepared review
    instructions.

    The JSON file must contain either:

    - A JSON array of change objects, or
    - A JSON object with a ``"changes"`` key containing the array.

    Each change object has the fields: ``fragment_id`` (int),
    ``change_type`` (``"modify"`` / ``"delete"`` / ``"append_after"``),
    ``new_text`` (string or null), ``justification`` (string).

    Args:
        document_path: Path to the input .docx file.
        changes_file: Path to the JSON file with changes.
        output_path: Where to save.  Defaults to ``<stem>_redlined.docx``.
        author: Author name for tracked changes and comments.
        validate: Run structural validation after applying (default True).

    Returns:
        Summary with change counts, output path, and validation result.
    """
    changes_path = Path(changes_file)
    if not changes_path.exists():
        msg = f"Changes file not found: {changes_file}"
        raise ToolError(msg)

    try:
        raw_text = changes_path.read_text(encoding="utf-8")
    except Exception as exc:
        msg = f"Failed to read changes file: {changes_file} ({exc})"
        raise ToolError(msg) from exc

    try:
        raw_data = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        msg = f"Invalid JSON in changes file: {changes_file} ({exc})"
        raise ToolError(msg) from exc

    # Support {"changes": [...]} wrapper
    if isinstance(raw_data, dict):
        if "changes" not in raw_data:
            msg = (
                "Changes file contains a JSON object but no 'changes' key. "
                'Expected a JSON array or {"changes": [...]}.'
            )
            raise ToolError(msg)
        raw_data = raw_data["changes"]

    try:
        change_params = _CHANGES_ADAPTER.validate_python(raw_data)
    except ValidationError as exc:
        msg = f"Invalid changes in {changes_file}:\n{exc}"
        raise ToolError(msg) from exc

    return _apply_and_save(document_path, change_params, output_path, author, validate)


@mcp.tool(
    annotations={
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def validate_document_tool(document_path: str) -> str:
    """Check a .docx file for structural issues.

    Runs validation checks on the document's OOXML structure:

    - **Annotation ID isolation** -- tracked-change and comment IDs must
      not collide across groups.
    - **Comment integrity** -- every comment must have matching range
      markers and references.
    - **Tracked-change attributes** -- ``<w:ins>`` / ``<w:del>`` must
      carry required ``w:id``, ``w:author``, ``w:date``.
    - **Package consistency** -- content-type and relationship entries
      must be present when comments.xml exists.

    Args:
        document_path: Path to the .docx file to validate.

    Returns:
        Human-readable validation summary.
    """
    doc = _load_document(document_path)
    result = _validate_document(doc)
    return _format_validation(result.errors, result.warnings)


@mcp.tool(
    annotations={
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def diff_fragments(
    original_path: str,
    modified_path: str,
) -> str:
    """Compare two .docx files and show paragraph-level text differences.

    Extracts the pseudo-Markdown text from each paragraph in both documents,
    then produces a word-level diff for each fragment position.

    Designed for comparing an original document with its edited version
    (same structure, local edits).  Paragraphs are matched by position
    (fragment 1 vs fragment 1, etc.) -- this tool does **not** detect
    paragraph reordering.  For documents with very different structures,
    the output will show extensive changes.

    Args:
        original_path: Path to the original .docx file.
        modified_path: Path to the modified .docx file.

    Returns:
        Human-readable diff showing changes per fragment.
    """
    doc_a = _load_document(original_path)
    doc_b = _load_document(modified_path)

    frags_a = dict(document_to_fragments(doc_a.paragraphs))
    frags_b = dict(document_to_fragments(doc_b.paragraphs))

    all_ids = sorted(set(frags_a) | set(frags_b))
    if not all_ids:
        return "Both documents are empty."

    differ = DmpWordDiffer()
    lines: list[str] = []

    for fid in all_ids:
        text_a = frags_a.get(fid)
        text_b = frags_b.get(fid)

        if text_a is not None and text_b is None:
            lines.append(f"Fragment {fid}: deleted (only in original)")
            if text_a:
                lines.append(f"  - {text_a}")
        elif text_a is None and text_b is not None:
            lines.append(f"Fragment {fid}: added (only in modified)")
            if text_b:
                lines.append(f"  + {text_b}")
        elif text_a == text_b:
            lines.append(f"Fragment {fid}: unchanged")
        else:
            assert text_a is not None and text_b is not None
            chunks = differ.diff(text_a, text_b)
            has_changes = any(c.op != DiffOp.EQUAL for c in chunks)
            if not has_changes:
                lines.append(f"Fragment {fid}: unchanged")
                continue
            lines.append(f"Fragment {fid}: modified")
            for chunk in chunks:
                if chunk.op == DiffOp.DELETE:
                    lines.append(f"  - {chunk.text}")
                elif chunk.op == DiffOp.INSERT:
                    lines.append(f"  + {chunk.text}")
                # EQUAL chunks are omitted for brevity

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Resource
# ---------------------------------------------------------------------------


@mcp.resource("docx-fragments://{document_path}")
def get_fragments(document_path: str) -> str:
    """Browse the paragraph fragments of a .docx file.

    Returns the tagged text representation of the document's paragraphs,
    each identified by a 1-based fragment ID.  No caching -- the file is
    re-read on every access to reflect the latest state on disk.

    The ``document_path`` in the URI must be URL-encoded if it contains
    path separators.  For example::

        docx-fragments://%2Fhome%2Fuser%2Fcontract.docx

    Args:
        document_path: Path to the .docx file (URL-decoded automatically).

    Returns:
        Tagged text: ``<f=1>text</f=1>`` per paragraph.
    """
    doc = _load_document(document_path)
    fragments = document_to_fragments(doc.paragraphs)
    return fragments_to_tagged_text(fragments)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Run the MCP server via stdio transport."""
    mcp.run(transport="stdio")
