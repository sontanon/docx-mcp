"""MCP server for the docx-mcp legal-document redlining engine.

Exposes the redlining pipeline as MCP tools so that LLM clients (Claude Desktop,
Cursor, ChatGPT, etc.) can extract document fragments (paragraphs and tables),
apply tracked changes, validate output, and compare document versions -- all via
filesystem paths.

Tools:
    extract_fragments   Read a .docx and return paragraphs and tables as text.
    apply_changes       Apply tracked changes (inline list) to paragraphs and tables.
    apply_changes_from_file  Apply tracked changes from a JSON file.
    validate_document   Structural validation of a .docx file.
    diff_fragments      Compare two .docx files (paragraphs and tables).

Resource:
    docx://{document_path}/fragments  Browse document fragments (paragraphs and tables).

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
from typing import Annotated, Literal

from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from pydantic import BaseModel, Discriminator, Field, Tag, TypeAdapter, ValidationError

from docx_mcp.converter import (
    fragments_to_json_interleaved,
    fragments_to_tagged_text_interleaved,
    full_to_fragments,
)
from docx_mcp.differ import DmpWordDiffer
from docx_mcp.document import DocxDocument
from docx_mcp.models import (
    Change,
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
from docx_mcp.table_utils import parse_cell_id
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
# Pydantic models for tool parameters
# ---------------------------------------------------------------------------


class ParagraphChangeParam(BaseModel):
    """A single tracked change to apply to a document paragraph.

    Represents one modification, deletion, or insertion operation targeting a
    specific paragraph in the document. Fragment IDs come from ``extract_fragments``
    and are 1-based paragraph indices (top-level ``<w:p>`` elements in ``<w:body>``).

    Font Inheritance (append_after only)
    -------------------------------------

    When you append a new paragraph, the font family, size, and color are
    automatically copied from the reference paragraph's first text-bearing run.
    Bold, italic, and underline formatting from pseudo-Markdown is layered on top
    of this inherited base. For example, if the reference paragraph uses Times New
    Roman 12pt, your appended paragraph will also use Times New Roman 12pt, even
    if you only specify ``**bold**``.

    Blank Line Management
    ---------------------

    Legal documents typically separate clauses with blank paragraphs. Use:

    - ``blank_lines_before`` / ``blank_lines_after`` when appending to maintain spacing
    - ``delete_next_blanks`` when deleting a clause that has a trailing blank separator

    All blank lines are marked as tracked insertions/deletions and appear in the
    redlined document.

    Validation Rules
    ----------------

    - ``new_text`` is required for ``modify`` and ``append_after``; must be
      null/omitted for ``delete``
    - ``blank_lines_before`` / ``blank_lines_after`` only valid with
      ``append_after`` (error if used with modify/delete)
    - ``delete_next_blanks`` only valid with ``delete`` (error if used with
      modify/append_after)
    - Each paragraph targeted by ``delete_next_blanks`` must be blank
      (whitespace-only); error if non-blank

    Examples
    --------

    Modify with pseudo-Markdown formatting::

        {
          "fragment_id": 3,
          "change_type": "modify",
          "new_text": "The Company **shall** provide **written** notice within _thirty (30)_ days.",
          "justification": "Strengthened obligation and clarified timeline."
        }

    Delete with trailing blank removal::

        {
          "fragment_id": 15,
          "change_type": "delete",
          "justification": "Removed obsolete governing law clause.",
          "delete_next_blanks": 1
        }

    Append with blank line spacing::

        {
          "fragment_id": 20,
          "change_type": "append_after",
          "new_text": "**16. Governing Law.** This Agreement shall be governed by Delaware law.",
          "justification": "Added Delaware choice of law provision.",
          "blank_lines_before": 1,
          "blank_lines_after": 1
        }

    Attributes
    ----------

        fragment_id: 1-based paragraph index from ``extract_fragments``.
        change_type: ``"modify"``, ``"delete"``, or ``"append_after"``.
        new_text: Replacement text in pseudo-Markdown (``**bold**``, ``_italic_``,
            ``__underline__``). Required for modify and append_after; must be None for delete.
        justification: Human-readable reason. Becomes a Word comment attached to the change.
        blank_lines_before: Blank paragraphs to insert before appended paragraph
            (append_after only, default 0).
        blank_lines_after: Blank paragraphs to insert after appended paragraph
            (append_after only, default 0).
        delete_next_blanks: Number of trailing blank paragraphs to delete
            (delete only, default 0).
    """

    fragment_id: str = Field(
        description="Fragment ID. Body paragraphs use plain integers (e.g. '5'). "
        "Headers and footers use prefixed IDs (e.g. 'header_1.3', 'footer_2.1')."
    )
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

    # -- Optional spacing controls -------------------------------------------
    blank_lines_before: int = Field(
        default=0,
        ge=0,
        description=(
            "Number of blank paragraphs to insert before the appended paragraph. "
            "Only used with append_after.  Most legal documents separate clauses "
            "with one blank line; set to 1 to match that convention."
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
            "Set to 1 when a clause is followed by a blank separator line "
            "that should be removed together with the clause.  Each targeted "
            "paragraph must be blank (whitespace-only); an error is raised "
            "if a non-blank paragraph is encountered."
        ),
    )


class TableChangeParam(BaseModel):
    """A single tracked change to apply to a table cell.

    Targets a specific cell within a table using dotted cell reference
    notation: ``"table_id.row.col"`` (e.g., ``"2.1.3"`` for table 2,
    row 1, column 3).

    Only simple tables are supported (rectangular grid, no merged cells,
    no nested tables). Cell IDs come from ``extract_fragments`` output.

    Multi-Paragraph Cells
    ---------------------

    Use ``\\n`` to separate paragraphs within a cell in ``new_text``:

        "Header text\\nBody paragraph\\nFooter paragraph"

    The engine aligns old/new paragraphs positionally:
    - Same count: each paragraph is modified in place
    - More new: excess paragraphs are appended
    - More old: excess paragraphs are deleted

    Change Types
    ------------

    - ``modify_cell``: Replace cell text (word-level diff on each paragraph)
    - ``clear_cell``: Mark all cell paragraphs as deleted (cell becomes empty
      but the ``<w:tc>`` element remains)

    Validation Rules
    ----------------

    - ``new_text`` is required for ``modify_cell``; must be null/omitted
      for ``clear_cell``
    - ``cell_id`` must match an extracted cell from a simple table
    - Row and column indices must be in range

    Examples
    --------

    Modify a cell with multiple paragraphs::

        {
          "cell_id": "2.1.1",
          "change_type": "modify_cell",
          "new_text": "**Updated Header**\\nNew body text.",
          "justification": "Clarified header and updated body."
        }

    Clear a cell::

        {
          "cell_id": "3.2.3",
          "change_type": "clear_cell",
          "justification": "Removed obsolete data."
        }

    Attributes
    ----------

        cell_id: Dotted cell reference ``"table_id.row.col"`` from ``extract_fragments``.
        change_type: ``"modify_cell"`` or ``"clear_cell"``.
        new_text: New cell text in pseudo-Markdown. Use ``\\n`` for paragraph
            breaks. Required for modify_cell; must be None for clear_cell.
        justification: Human-readable reason. Becomes a Word comment attached
            to the cell's first paragraph.
    """

    cell_id: str = Field(
        description='Dotted cell reference "table_id.row.col" from extract_fragments'
    )
    change_type: Literal["modify_cell", "clear_cell"] = Field(
        description='Type of cell change: "modify_cell" or "clear_cell"'
    )
    new_text: str | None = Field(
        default=None,
        description=(
            "New cell text in pseudo-Markdown. Use \\n for paragraph breaks. "
            "Required for modify_cell. Omit for clear_cell."
        ),
    )
    justification: str = Field(
        description="Reason for this change. Becomes a Word comment.",
    )


def _discriminate_change_param(v: dict | BaseModel) -> str:
    """Discriminator function for ChangeParam union.

    Auto-discriminates based on change_type value or field presence
    (fragment_id vs cell_id) for LLM-friendly API.
    """
    change_type = v.get("change_type", "") if isinstance(v, dict) else getattr(v, "change_type", "")

    # Discriminate by change_type value
    if change_type in ("modify", "delete", "append_after"):
        return "paragraph"
    if change_type in ("modify_cell", "clear_cell"):
        return "table"

    # Fallback: discriminate by field presence
    if isinstance(v, dict):
        has_fragment_id = "fragment_id" in v
        has_cell_id = "cell_id" in v
    else:
        has_fragment_id = hasattr(v, "fragment_id")
        has_cell_id = hasattr(v, "cell_id")

    if has_fragment_id:
        return "paragraph"
    if has_cell_id:
        return "table"

    msg = f"Cannot discriminate change type from: {v}"
    raise ValueError(msg)


ChangeParam = Annotated[
    Annotated[ParagraphChangeParam, Tag("paragraph")] | Annotated[TableChangeParam, Tag("table")],
    Discriminator(_discriminate_change_param),
]


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


def _convert_change_param_to_change(param: ChangeParam) -> Change:
    """Convert MCP DTO (ChangeParam) to internal domain model (Change).

    This translation layer allows the MCP API to use LLM-friendly compact
    representations (no 'kind' field, compact cell_id strings) while the
    internal code uses type-safe discriminated unions with explicit fields.

    Args:
        param: User-facing ChangeParam (ParagraphChangeParam or TableChangeParam).

    Returns:
        Internal domain model (ParagraphChange or TableChange).
    """
    if isinstance(param, ParagraphChangeParam):
        return ParagraphChange(
            kind="paragraph",
            fragment_id=param.fragment_id,
            change_type=ParagraphChangeType(param.change_type),
            new_text=param.new_text,
            justification=param.justification,
            blank_lines_before=param.blank_lines_before,
            blank_lines_after=param.blank_lines_after,
            delete_next_blanks=param.delete_next_blanks,
        )
    else:  # TableChangeParam
        table_id, row, col = parse_cell_id(param.cell_id)
        return TableChange(
            kind="table",
            table_id=table_id,
            row=row,
            col=col,
            change_type=TableChangeType(param.change_type),
            new_text=param.new_text,
            justification=param.justification,
        )


def _convert_changes(params: list[ChangeParam]) -> list[Change]:
    """Convert a list of ChangeParam DTOs to internal Change models.

    Returns:
        List of internal domain models (mix of ParagraphChange and TableChange).
    """
    return [_convert_change_param_to_change(p) for p in params]


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
) -> str:
    """Shared implementation for both apply tools."""
    # Convert MCP params to internal domain models
    changes = _convert_changes(change_params)
    config = RedlineConfig(author=author)
    out = _resolve_output_path(document_path, output_path)

    try:
        doc = apply_redlines(
            document_path,
            changes,
            config=config,
        )
    except FileNotFoundError as exc:
        msg = f"File not found: {document_path}"
        raise ToolError(msg) from exc
    except ValueError as exc:
        raise ToolError(str(exc)) from exc

    # Build summary (use original params for counting)
    counts = Counter(p.change_type for p in change_params)
    summary_parts = []
    for ct in ("modify", "delete", "append_after", "modify_cell", "clear_cell"):
        if counts[ct]:
            summary_parts.append(f"{counts[ct]} {ct}")
    change_summary = ", ".join(summary_parts)

    result = _validate_document(doc)

    lines = [
        f"Applied {len(change_params)} change(s) ({change_summary}) to {Path(document_path).name}.",
        f"Output saved to {out}.",
        _format_validation(result.errors, result.warnings),
    ]

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
) -> str:
    """Extract text from a .docx file as tagged fragments.

    Call this first to get fragment IDs, then use those IDs in ``apply_changes``.

    Output format::

        <f=1>**Title**</f=1>
        <f=2>Body paragraph.</f=2>
        <table=3 rows=2 cols=3>
        <cell=3.1.1>Header</cell=3.1.1>
        <cell=3.1.2 span="2">Merged cell</cell=3.1.2>
        </table=3>

    Fragment IDs:
      - Body: ``"1"``, ``"2"``, ...
      - Headers: ``"header_1.1"``, ``"header_1.2"``, ...
      - Footers: ``"footer_1.1"``, ``"footer_2.1"``, ...
      - Table cells: ``"table_id.row.col"`` (e.g., ``"3.1.2"``)

    Formatting: ``**bold**``, ``_italic_``, ``__underline__``. ``\n`` for paragraph
    breaks.

    Limitations:
      - Images, nested tables, VML text boxes are skipped.
      - Pre-existing tracked changes are hard-rejected.

    Args:
        document_path: Absolute path to the .docx file.

    Returns:
        Tagged text string with fragment and table markup.
    """
    doc = _load_document(document_path)

    dirty_parts = doc.has_tracked_changes()
    if dirty_parts:
        locations = ", ".join(dirty_parts)
        msg = (
            f"Document contains pre-existing tracked changes in {locations}. "
            "Please accept or reject all changes before redlining."
        )
        raise ToolError(msg)

    result = full_to_fragments(
        doc,
        hyperlink_resolver=doc.resolve_hyperlink_url,
    )

    return fragments_to_tagged_text_interleaved(result.items)


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
) -> str:
    """Apply tracked changes to a .docx file and save a new redlined document.

    Produces Word-compatible tracked changes (``w:ins`` / ``w:del``) with comments.
    Always call ``extract_fragments`` first to get fragment IDs.

    Paragraph changes (use ``fragment_id``):
      - ``modify``: Word-level diff. ``new_text`` required.
      - ``delete``: Marks paragraph deleted. Omit ``new_text``.
      - ``append_after``: Insert after target. ``new_text`` required.

    Table changes (use ``cell_id`` like ``"2.1.3"``):
      - ``modify_cell``: Replace cell text. ``new_text`` required. Use ``\n`` for
        multi-paragraph cells.
      - ``clear_cell``: Delete cell content. Omit ``new_text``.

    Rules:
      - ``delete_next_blanks`` only with ``delete``.
      - ``blank_lines_before`` / ``blank_lines_after`` only with ``append_after``.
      - Header/footer changes work but comments are silently dropped.
      - Documents with pre-existing tracked changes are hard-rejected.

    ``new_text`` uses pseudo-Markdown: ``**bold**``, ``_italic_``, ``__underline__``.

    Example::

        changes = [
            {
                "fragment_id": 1,
                "change_type": "modify",
                "new_text": "**Updated title**",
                "justification": "Fixed typo"
            },
            {
                "cell_id": "2.1.1",
                "change_type": "modify_cell",
                "new_text": "New header",
                "justification": "Clarified"
            }
        ]

    Args:
        document_path: Absolute path to the input .docx file.
        changes: List of change objects (paragraph or table cell changes).
        output_path: Where to save the redlined document. Defaults to
            ``<stem>_redlined.docx`` beside the input file.
        author: Author name for tracked changes and comments. Defaults to
            "AI Review".

    Returns:
        Summary string with change counts, output path, and validation result.
    """
    return _apply_and_save(document_path, changes, output_path, author)


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
) -> str:
    """Same as ``apply_changes`` but reads the change list from a JSON file.

    Useful for large change sets that would exceed token limits in a direct tool
    call, or for reusing a change set across multiple runs.

    The JSON file must contain either a bare array of change objects, or an
    object with a ``"changes"`` key::

         [
           {"fragment_id": 1, "change_type": "modify", "new_text": "..."},
           {"cell_id": "2.1.1", "change_type": "modify_cell", "new_text": "..."}
         ]

    See ``apply_changes`` for change object schema and rules.

    Args:
        document_path: Absolute path to the input .docx file.
        changes_file: Absolute path to the JSON file containing changes.
        output_path: Where to save the redlined document. Defaults to
            ``<stem>_redlined.docx`` beside the input file.
        author: Author name for tracked changes and comments. Defaults to
            "AI Review".

    Returns:
        Summary string with change counts, output path, and validation result.
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

    return _apply_and_save(document_path, change_params, output_path, author)


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

    Runs validation checks on the document's OOXML structure to ensure it will
    open correctly in Microsoft Word and that all tracked changes and comments
    are properly formed.

    Use this tool:

    - After calling ``apply_changes`` to verify the redlined document is valid
      (automatically enabled by default via ``validate=True`` parameter)
    - When debugging a document that won't open correctly in Word
    - When verifying that an existing redlined document has proper structure

    Validation Checks
    -----------------

    - **Annotation ID isolation**: Tracked-change and comment IDs must not collide
      across groups. Each ``<w:ins>``, ``<w:del>``, and ``<w:comment>`` needs a
      globally unique ID within the document.

    - **Comment integrity**: Every ``<w:comment>`` in comments.xml must have
      matching ``<w:commentRangeStart>`` / ``<w:commentRangeEnd>`` markers in
      the document body, and vice versa.

    - **Tracked-change attributes**: Every ``<w:ins>`` and ``<w:del>`` must have
      required attributes: ``w:id`` (unique ID), ``w:author`` (author name), and
      ``w:date`` (timestamp).

    - **Package consistency**: Content-type and relationship entries must be
      present in the .docx ZIP structure when comments.xml exists.

    Example Output
    --------------

    Success case::

        "Validation: passed (0 errors, 0 warnings)."

    Failure case::

        "Validation: FAILED (2 error(s), 1 warning(s)).
          Error 1: Annotation ID collision: ID 5 used by both tracked change and comment
          Error 2: Orphaned comment range: commentRangeStart with id=3 has no matching end
          Warning 1: Comment with id=7 is not referenced by any comment range"

    Args:
        document_path: Absolute path to the .docx file to validate.

    Returns:
        Human-readable validation summary showing pass/fail status, error count,
        and detailed messages for any issues found.
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
def audit_document_tool(
    document_path: str,
    format: Literal["text", "json"] = "text",
) -> str:
    """Audit a .docx file for structural issues and skipped content.

    Reports headers, footers, images, tables, section breaks, tracked changes,
    comments, and unsupported elements (footnotes, endnotes, text boxes).

    Args:
        document_path: Absolute path to the .docx file.
        format: Output format -- ``"text"`` (default) or ``"json"``.

    Returns:
        Audit report in the requested format.
    """
    from docx_mcp.audit import audit_document

    doc = _load_document(document_path)
    report = audit_document(doc)

    if format == "json":
        import json

        return json.dumps(report.to_dict(), indent=2, ensure_ascii=False)

    return report.to_text()


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
    """Compare two .docx files and show paragraph and table-level text differences.

    Extracts the pseudo-Markdown text from each paragraph and table cell in both
    documents, then produces a word-level diff for each fragment position. This is
    useful for understanding what changed between two versions of a document.

    Use this tool:

    - To compare an original document with its redlined version (see what changes
      were applied)
    - To verify that ``apply_changes`` produced the expected modifications
    - To understand differences between two versions of a document

    Important Limitations
    ---------------------

    Fragments are matched **by position** (fragment 1 vs fragment 1, fragment 2
    vs fragment 2, etc.). This tool does **not** detect fragment reordering or
    track moved sections. If the documents have very different structures (different
    fragment counts, major reordering), the output will show extensive changes.

    Best used for comparing documents with the same basic structure where you made
    local edits (word changes, clause deletions, appended sections, table cell
    modifications).

    Output Format
    -------------

    Each fragment is reported with its change status.

    Paragraphs::

        Fragment 1: unchanged
        Fragment 2: modified
          - shall deliver
          + must deliver immediately
        Fragment 3: unchanged
        Fragment 5: deleted (only in original)
          - This clause is removed.
        Fragment 10: added (only in modified)
          + This is a new clause.

    Tables::

        Table 56: modified
          Cell 56.2.2: modified
            - Gwendolyn Mahon, M.Sc., Ph.D
            + John H. Smith, Ph.D.
        Table 57: unchanged
        Table 58: dimensions changed (3x2 → 4x2)

    Lines starting with ``-`` show deleted text, ``+`` shows inserted text.
    Unchanged fragments are listed but their text is omitted for brevity.

    For tables, each modified cell is shown with its cell ID (table_id.row.col)
    followed by the word-level diff of the cell content.

    Difference from extract_fragments with markup=True
    --------------------------------------------------

    - ``extract_fragments`` with ``markup=True`` shows **tracked changes already
      present in a single document** (existing ``<w:ins>`` / ``<w:del>`` markup)

    - ``diff_fragments`` **compares two separate documents** and computes the
      differences between their plain text (ignoring any tracked changes)

    Args:
        original_path: Absolute path to the original .docx file.
        modified_path: Absolute path to the modified .docx file.

    Returns:
        Human-readable diff showing changes per fragment, with ``-`` for deletions
        and ``+`` for insertions. Tables show cell-level diffs.
    """
    doc_a = _load_document(original_path)
    doc_b = _load_document(modified_path)

    # Extract fragments (paragraphs and tables) from both documents
    result_a = full_to_fragments(doc_a)
    result_b = full_to_fragments(doc_b)

    # Build dictionaries mapping fragment_id -> FragmentItem
    frags_a: dict[str, tuple[str, str] | TableInfo | SkippedTableInfo] = {}
    for item in result_a.items:
        if isinstance(item, tuple):
            fid, _text = item
            frags_a[fid] = item
        else:  # TableInfo or SkippedTableInfo
            frags_a[str(item.table_id)] = item

    frags_b: dict[str, tuple[str, str] | TableInfo | SkippedTableInfo] = {}
    for item in result_b.items:
        if isinstance(item, tuple):
            fid, _text = item
            frags_b[fid] = item
        else:  # TableInfo or SkippedTableInfo
            frags_b[str(item.table_id)] = item

    all_ids = sorted(set(frags_a) | set(frags_b))
    if not all_ids:
        return "Both documents are empty."

    differ = DmpWordDiffer()
    lines: list[str] = []

    for fid in all_ids:
        item_a = frags_a.get(fid)
        item_b = frags_b.get(fid)

        # Handle missing fragments
        if item_a is not None and item_b is None:
            if isinstance(item_a, tuple):
                _fid, text_a = item_a
                lines.append(f"Fragment {fid}: deleted (only in original)")
                if text_a:
                    lines.append(f"  - {text_a}")
            else:  # Table
                lines.append(f"Table {fid}: deleted (only in original)")
            continue

        if item_a is None and item_b is not None:
            if isinstance(item_b, tuple):
                _fid, text_b = item_b
                lines.append(f"Fragment {fid}: added (only in modified)")
                if text_b:
                    lines.append(f"  + {text_b}")
            else:  # Table
                lines.append(f"Table {fid}: added (only in modified)")
            continue

        # Both exist - compare based on type
        assert item_a is not None and item_b is not None

        # Case 1: Both are paragraphs
        if isinstance(item_a, tuple) and isinstance(item_b, tuple):
            _fid_a, text_a = item_a
            _fid_b, text_b = item_b
            if text_a == text_b:
                lines.append(f"Fragment {fid}: unchanged")
            else:
                chunks = differ.diff(text_a, text_b)
                has_changes = any(c.op != DiffOp.EQUAL for c in chunks)
                if not has_changes:
                    lines.append(f"Fragment {fid}: unchanged")
                else:
                    lines.append(f"Fragment {fid}: modified")
                    for chunk in chunks:
                        if chunk.op == DiffOp.DELETE:
                            lines.append(f"  - {chunk.text}")
                        elif chunk.op == DiffOp.INSERT:
                            lines.append(f"  + {chunk.text}")

        # Case 3: Both are SkippedTableInfo
        elif isinstance(item_a, SkippedTableInfo) and isinstance(item_b, SkippedTableInfo):
            if item_a.reason == item_b.reason:
                lines.append(f"Table {fid}: unchanged (skipped: {item_a.reason})")
            else:
                lines.append(f"Table {fid}: skip reason changed")
                lines.append(f"  - {item_a.reason}")
                lines.append(f"  + {item_b.reason}")

        # Case 4: One skipped, one not (both must be table types)
        elif (
            not isinstance(item_a, tuple)
            and not isinstance(item_b, tuple)
            and isinstance(item_a, SkippedTableInfo) != isinstance(item_b, SkippedTableInfo)
        ):
            lines.append(f"Table {fid}: skip status changed")

        # Case 5: Both are TableInfo
        elif isinstance(item_a, TableInfo) and isinstance(item_b, TableInfo):
            # Check dimensions
            if item_a.rows != item_b.rows or item_a.cols != item_b.cols:
                lines.append(
                    f"Table {fid}: dimensions changed "
                    f"({item_a.rows}x{item_a.cols} → {item_b.rows}x{item_b.cols})"
                )
                continue

            # Compare cell-by-cell
            table_has_changes = False
            cell_changes: list[str] = []

            for row_idx in range(item_a.rows):
                for col_idx in range(item_a.cols):
                    cell_a = item_a.cells[row_idx][col_idx]
                    cell_b = item_b.cells[row_idx][col_idx]

                    if cell_a.text != cell_b.text:
                        table_has_changes = True
                        chunks = differ.diff(cell_a.text, cell_b.text)
                        has_diff = any(c.op != DiffOp.EQUAL for c in chunks)
                        if has_diff:
                            cell_changes.append(f"  Cell {cell_a.cell_id}: modified")
                            for chunk in chunks:
                                if chunk.op == DiffOp.DELETE:
                                    cell_changes.append(f"    - {chunk.text}")
                                elif chunk.op == DiffOp.INSERT:
                                    cell_changes.append(f"    + {chunk.text}")

            if table_has_changes:
                lines.append(f"Table {fid}: modified")
                lines.extend(cell_changes)
            else:
                lines.append(f"Table {fid}: unchanged")

        # Case 2: Type mismatch (paragraph vs table)
        else:
            lines.append(f"Fragment {fid}: type changed (paragraph ↔ table)")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Resource
# ---------------------------------------------------------------------------


@mcp.resource("docx-fragments://{document_path}")
def get_fragments(document_path: str) -> str:
    """Browse the paragraph and table fragments of a .docx file.

    Returns the tagged text representation of the document's paragraphs and
    tables, each identified by a 1-based fragment ID. No caching -- the file
    is re-read on every access to reflect the latest state on disk.

    The ``document_path`` in the URI must be URL-encoded if it contains
    path separators. For example::

        docx-fragments://%2Fhome%2Fuser%2Fcontract.docx

    Args:
        document_path: Path to the .docx file (URL-decoded automatically).

    Returns:
        Tagged text: ``<f=N>text</f=N>`` for paragraphs, ``<table=N ...>`` with
        ``<cell=id>text</cell=id>`` for table cells.
    """
    doc = _load_document(document_path)
    result = full_to_fragments(doc)
    return fragments_to_tagged_text_interleaved(result.items)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    """Run the MCP server via stdio transport."""
    mcp.run(transport="stdio")
