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
            blank_lines_before=p.blank_lines_before,
            blank_lines_after=p.blank_lines_after,
            delete_next_blanks=p.delete_next_blanks,
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
    markup: bool = False,
) -> str:
    """Read a .docx file and return its paragraphs as text.

    This is typically the **first step** in a redlining workflow. Each paragraph
    is identified by a 1-based fragment ID that you will use to target changes
    in ``apply_changes``.

    Fragment IDs are position-based, corresponding to top-level ``<w:p>`` elements
    in ``<w:body>``. Fragment 1 is the first paragraph, fragment 2 is the second,
    and so on. These IDs remain stable as long as you don't add/remove paragraphs
    before your target location.

    Output Formats
    --------------

    The **tagged** format (default) wraps each paragraph with XML-like tags::

        <f=1>**CONFIDENTIALITY AGREEMENT**</f=1>
        <f=2>This Agreement is entered into as of January 1, 2025.</f=2>
        <f=3></f=3>
        <f=4>**1. Definitions.** The following terms have the meanings set forth below.</f=4>

    The **json** format returns a JSON array of objects::

        [
          {"fragment_id": 1, "text": "**CONFIDENTIALITY AGREEMENT**"},
          {"fragment_id": 2, "text": "This Agreement is entered into as of January 1, 2025."},
          {"fragment_id": 3, "text": ""},
          {
            "fragment_id": 4,
            "text": "**1. Definitions.** The following terms have the meanings set forth below."
          }
        ]

    Text Formatting
    ---------------

    Text uses pseudo-Markdown for inline formatting:

    - ``**bold**``
    - ``_italic_``
    - ``__underline__``

    Unicode characters (smart quotes, em dashes, section symbols, non-breaking
    spaces) are preserved as-is.

    Tracked Changes (markup=True)
    ------------------------------

    When ``markup=True``, existing tracked changes in the document are shown inline:

    - Inserted text is wrapped with ``++…++``
    - Deleted text is wrapped with ``~~…~~``

    This is useful for inspecting or validating redlined documents that already
    contain tracked changes. For clean documents, leave ``markup=False`` (default).

    Typical Workflow
    ----------------

    1. Call ``extract_fragments`` to see document structure and get fragment IDs
    2. Identify which paragraphs need changes (by reading the text)
    3. Construct a list of ``ChangeParam`` objects with appropriate fragment_ids
    4. Call ``apply_changes`` with the change list
    5. Open the output file in Microsoft Word to review tracked changes

    Args:
        document_path: Absolute path to the .docx file.
        format: Output format -- ``"tagged"`` (default) or ``"json"``.
        markup: Show existing tracked changes inline (default False).

    Returns:
        Fragment text in the requested format (string).
    """
    doc = _load_document(document_path)
    fragments = document_to_fragments(doc.paragraphs, markup=markup)

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

    Produces a Word document with professional tracked changes (``w:ins`` / ``w:del``)
    and comments, indistinguishable from a lawyer's redline. Opens cleanly in
    Microsoft Word with the Review tab showing all changes and comments.

    Typical Workflow
    ----------------

    1. Call ``extract_fragments`` to see document structure and get fragment IDs
    2. Identify which paragraphs need changes (by reading the text)
    3. Construct a list of ``ChangeParam`` objects with appropriate fragment_ids
    4. Call this tool (``apply_changes``) with the changes list
    5. Open the output file in Microsoft Word to review tracked changes

    Change Types
    ------------

    - **modify**: Word-level diff produces fine-grained tracked insertions and
      deletions. For example, changing "shall deliver" to "must deliver immediately"
      will show "shall" as deleted and "must" as inserted, preserving "deliver".

    - **delete**: Entire paragraph marked as deleted. Use ``delete_next_blanks``
      to remove trailing blank separator lines along with the deleted clause.

    - **append_after**: New paragraph inserted after the target fragment. The font
      family, size, and color are automatically inherited from the reference
      paragraph. Use ``blank_lines_before`` / ``blank_lines_after`` to maintain
      legal document spacing conventions.

    Font Inheritance
    ----------------

    When appending new paragraphs (``append_after``), the font family, size, and
    color are automatically copied from the reference paragraph's first text-bearing
    run. Pseudo-Markdown formatting (``**bold**``, ``_italic_``, ``__underline__``)
    is applied on top of the inherited base. For example, if the reference paragraph
    uses Times New Roman 12pt, your appended paragraph will also use Times New Roman
    12pt, even if you only specify ``**bold**`` in the ``new_text``.

    Blank Line Management
    ---------------------

    Legal documents typically separate clauses with blank paragraphs (one blank line
    between sections). When appending new clauses, set ``blank_lines_before=1`` and/or
    ``blank_lines_after=1`` to maintain this spacing. When deleting a clause that has
    a trailing blank separator, set ``delete_next_blanks=1`` to remove it along with
    the clause itself.

    All blank lines are marked as tracked insertions/deletions and appear in the
    redlined document.

    Error Handling
    --------------

    Common errors that will cause this tool to fail:

    - ``fragment_id`` out of range (must be 1..N where N is total paragraph count)
    - ``new_text`` missing when required (modify/append_after)
    - ``new_text`` provided for delete (must be null/omitted)
    - ``delete_next_blanks`` targets a non-blank paragraph (only whitespace-only
      paragraphs can be deleted this way)
    - ``blank_lines_before`` / ``blank_lines_after`` used with modify or delete
      (only valid with append_after)
    - ``delete_next_blanks`` used with modify or append_after (only valid with
      delete)

    Example
    -------

    Applying multiple changes to an NDA::

        changes = [
            {
                "fragment_id": 1,
                "change_type": "modify",
                "new_text": "**MUTUAL NON-DISCLOSURE AGREEMENT** (Revised 2025)",
                "justification": "Updated title and year"
            },
            {
                "fragment_id": 35,
                "change_type": "delete",
                "justification": "Removed Section 7 (proprietary rights legends)",
                "delete_next_blanks": 1
            },
            {
                "fragment_id": 39,
                "change_type": "append_after",
                "new_text": (
                    "**18. Amendments.** No amendment shall be effective unless "
                    "in writing and signed by both parties."
                ),
                "justification": "Added amendments clause per legal review",
                "blank_lines_before": 1,
                "blank_lines_after": 1
            }
        ]

        result = apply_changes(
            document_path="/path/to/contract.docx",
            changes=changes,
            author="Legal AI",
            validate=True
        )

        # Result:
        # "Applied 3 change(s) (1 modify, 1 delete, 1 append_after) to contract.docx.
        #  Output saved to /path/to/contract_redlined.docx.
        #  Validation: passed (0 errors, 0 warnings)."

    Args:
        document_path: Absolute path to the input .docx file.
        changes: List of ``ChangeParam`` objects describing the changes to apply.
        output_path: Where to save the redlined document. Defaults to
            ``<stem>_redlined.docx`` beside the input file.
        author: Author name for tracked changes and comments (appears in Word's
            Review tab). Defaults to "AI Review".
        validate: Run structural validation after applying changes (default True,
            recommended). Checks for annotation ID collisions, comment integrity,
            and tracked-change attributes.

    Returns:
        Summary string with change counts, output path, and validation result.
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

    Same behavior as ``apply_changes`` but reads the change list from a JSON file
    on disk. This is useful when:

    - You have a large change set that would exceed token limits in a direct tool call
    - You want to save and reuse a change set across multiple runs
    - You're working with pre-prepared review instructions from another system

    JSON File Format
    ----------------

    The JSON file must contain either:

    1. A JSON array of change objects (bare array)::

        [
          {
            "fragment_id": 1,
            "change_type": "modify",
            "new_text": "Updated text",
            "justification": "Reason for change"
          },
          {
            "fragment_id": 5,
            "change_type": "delete",
            "justification": "Removed obsolete clause",
            "delete_next_blanks": 1
          }
        ]

    2. A JSON object with a ``"changes"`` key containing the array::

        {
          "changes": [
            {
              "fragment_id": 1,
              "change_type": "modify",
              "new_text": "Updated text",
              "justification": "Reason for change"
            }
          ]
        }

    Each change object supports all fields from ``ChangeParam``:

    - ``fragment_id`` (int, required): 1-based paragraph index
    - ``change_type`` (string, required): ``"modify"``, ``"delete"``, or ``"append_after"``
    - ``new_text`` (string or null): Required for modify/append_after, omit for delete
    - ``justification`` (string, required): Reason for the change
    - ``blank_lines_before`` (int, optional): Blank paragraphs before (append_after only, default 0)
    - ``blank_lines_after`` (int, optional): Blank paragraphs after (append_after only, default 0)
    - ``delete_next_blanks`` (int, optional): Trailing blanks to delete (delete only, default 0)

    The file must be UTF-8 encoded. Path separators and spaces in paths are supported.

    Args:
        document_path: Absolute path to the input .docx file.
        changes_file: Absolute path to the JSON file containing changes.
        output_path: Where to save the redlined document. Defaults to
            ``<stem>_redlined.docx`` beside the input file.
        author: Author name for tracked changes and comments (appears in Word's
            Review tab). Defaults to "AI Review".
        validate: Run structural validation after applying changes (default True,
            recommended).

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
def diff_fragments(
    original_path: str,
    modified_path: str,
) -> str:
    """Compare two .docx files and show paragraph-level text differences.

    Extracts the pseudo-Markdown text from each paragraph in both documents,
    then produces a word-level diff for each fragment position. This is useful
    for understanding what changed between two versions of a document.

    Use this tool:

    - To compare an original document with its redlined version (see what changes
      were applied)
    - To verify that ``apply_changes`` produced the expected modifications
    - To understand differences between two versions of a document

    Important Limitations
    ---------------------

    Paragraphs are matched **by position** (fragment 1 vs fragment 1, fragment 2
    vs fragment 2, etc.). This tool does **not** detect paragraph reordering or
    track moved sections. If the documents have very different structures (different
    paragraph counts, major reordering), the output will show extensive changes.

    Best used for comparing documents with the same basic structure where you made
    local edits (word changes, clause deletions, appended sections).

    Output Format
    -------------

    Each fragment is reported with its change status::

        Fragment 1: unchanged
        Fragment 2: modified
          - shall deliver
          + must deliver immediately
        Fragment 3: unchanged
        Fragment 5: deleted (only in original)
          - This clause is removed.
        Fragment 10: added (only in modified)
          + This is a new clause.

    Lines starting with ``-`` show deleted text, ``+`` shows inserted text.
    Unchanged fragments are listed but their text is omitted for brevity.

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
        and ``+`` for insertions.
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
