# docx-mcp

Legal document redlining engine. Takes AI-generated changes (structured JSON)
and applies them as professional tracked changes with comments inside `.docx`
files. The output is indistinguishable from what a lawyer would produce in
Microsoft Word -- proper `w:ins`/`w:del` markup, comment annotations with
justification text, and preserved formatting.

## Installation

Requires Python 3.14+.

```bash
uv sync
```

## Quick start

### Python API

```python
from docx_mcp import (
    ParagraphChange, ParagraphChangeType,
    TableChange, TableChangeType,
    RedlineConfig, apply_redlines,
)

changes = [
    # Modify a body paragraph
    ParagraphChange(
        kind="paragraph",
        fragment_id="3",              # ← str (was int in v0.1.0)
        change_type=ParagraphChangeType.MODIFY,
        new_text="The Company **shall** provide written notice.",
        justification="Strengthened obligation language.",
    ),
    # Delete a paragraph
    ParagraphChange(
        kind="paragraph",
        fragment_id="5",
        change_type=ParagraphChangeType.DELETE,
        justification="Removed redundant clause.",
    ),
    # Append a new paragraph
    ParagraphChange(
        kind="paragraph",
        fragment_id="7",
        change_type=ParagraphChangeType.APPEND_AFTER,
        new_text="The foregoing shall survive termination.",
        justification="Added survival provision.",
    ),
    # Modify a header paragraph
    ParagraphChange(
        kind="paragraph",
        fragment_id="header_1.1",
        change_type=ParagraphChangeType.MODIFY,
        new_text="CONFIDENTIAL",
        justification="Updated header text.",
    ),
    # Modify a table cell
    TableChange(
        kind="table",
        table_id=2,
        row=1,
        col=1,
        change_type=TableChangeType.MODIFY_CELL,
        new_text="Updated **cell** content",
        justification="Corrected table entry.",
    ),
    # Clear a table cell
    TableChange(
        kind="table",
        table_id=2,
        row=3,
        col=2,
        change_type=TableChangeType.CLEAR_CELL,
        justification="Removed obsolete data.",
    ),
]

doc = apply_redlines("contract.docx", changes)
doc.save("contract_redlined.docx")
```

### CLI

```bash
# Extract fragment text from a document
docx-mcp convert input.docx
docx-mcp convert input.docx --format json

# Apply changes
docx-mcp apply input.docx changes.json -o output.docx

# Validate a redlined document
docx-mcp validate output.docx

# Audit a document for structural issues
docx-mcp audit input.docx
docx-mcp audit input.docx --format json
```

> **Note:** The CLI `convert` command extracts body content only (no headers,
> footers, or tables). For full-document extraction, use the MCP
> `extract_fragments` tool or the Python `full_to_fragments()` function.

### MCP server

The library includes an MCP server so that LLM clients (Claude Desktop, Cursor,
etc.) can redline `.docx` files directly.

```bash
# Start the server (stdio transport)
docx-mcp-server
```

**Configure in Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/docx-mcp", "docx-mcp-server"]
    }
  }
}
```

**Configure in Cursor** (`.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/docx-mcp", "docx-mcp-server"]
    }
  }
}
```

#### Tools

| Tool | Description |
|------|-------------|
| `extract_fragments` | Read a `.docx` and return paragraphs, tables, headers, and footers as tagged text |
| `apply_changes` | Apply tracked changes from an inline list and save |
| `apply_changes_from_file` | Apply tracked changes from a JSON file on disk |
| `validate_document_tool` | Run structural validation checks |
| `diff_fragments` | Compare two `.docx` files paragraph-by-paragraph (full document) |
| `audit_document_tool` | Audit a `.docx` for headers, images, tables, section breaks, and more |

#### Resource

| URI | Description |
|-----|-------------|
| `docx-fragments://{document_path}` | Browse paragraph fragments (URL-encode the path) |

#### Example workflow

An LLM client would typically:

1. Call `extract_fragments` to read the document and get fragment IDs.
2. Reason about the content and construct a list of changes.
3. Call `apply_changes` with the change list to produce a redlined document.
4. Optionally call `diff_fragments` to compare original vs. redlined output.

## Concepts

### Fragments

Documents are decomposed into **fragments**: paragraphs, tables, headers, and
footers, all indexed in document order. Each fragment has a **string ID**.

Fragment IDs:

| Pattern | Meaning | Example |
|---------|---------|---------|
| `"1"`, `"2"`, … | Body paragraphs / tables | `<f=1>Introduction.</f=1>` |
| `"header_P.I"` | Header part P, paragraph I | `<f=header_1.3>Confidential</f=header_1.3>` |
| `"footer_P.I"` | Footer part P, paragraph I | `<f=footer_2.1>Page 1 of 10</f=footer_2.1>` |

Tables and body paragraphs share the same ID space (they interleave in document
order). Fragment `"3"` might be a table and fragment `"4"` a paragraph.

Use `extract_fragments` (MCP) or `full_to_fragments()` (Python) to see the
fragment map for any document:

```
<f=1>Introduction paragraph.</f=1>
<f=2>**Definitions.** The following terms shall apply.</f=2>
<table=3 rows=2 cols=3>
<cell=3.1.1 span="2">Merged Header</cell=3.1.1>
<cell=3.1.3>Header C</cell=3.1.3>
<cell=3.2.1>Data 1</cell=3.2.1>
<cell=3.2.2>Data 2</cell=3.2.2>
<cell=3.2.3>Data 3</cell=3.2.3>
</table=3>
<f=4>Closing paragraph. See [Section 2](https://example.com).</f=4>
<f=header_1.1>Confidential</f=header_1.1>
<f=footer_1.1>Page 1 of 10</f=footer_1.1>
```

### Tables

#### Simple tables

Simple (rectangular) tables are extracted as `<table=N>` blocks. Each cell has
a `cell_id` in `"table_id.row.col"` format (e.g., `"3.1.2"`).

#### Merged-cell tables

Tables with horizontally or vertically merged cells (`gridSpan` / `vMerge`) are
now supported. Merge spans are shown as attributes:

- `span="2"` — cell spans 2 columns (horizontal merge)
- `vspan="3"` — cell spans 3 rows (vertical merge)

Spanned-over cells (positions covered by a merge) are **omitted from output**.
For example, if `cell=3.1.1` has `span="2"`, then `cell=3.1.2` does not appear.

When targeting merged cells with changes, always target the **originating cell**
(the one with the `span`/`vspan` attribute). Targeting a spanned-over position
raises a `ValueError`.

#### Skipped tables

Tables that cannot be processed (nested tables, malformed merges, tables inside
headers/footers) appear as:

```
<table=5 skipped reason="table 5, cell 2.3 contains nested table"/>
```

### Headers and footers

Header and footer paragraphs are extracted with prefixed fragment IDs:
`header_1.1`, `footer_2.1`, etc. The first number is the 1-based part index
(usually `1` for the default header/footer), the second is the 1-based
paragraph index within that part.

Header/footer paragraphs can be modified, deleted, and appended to just like
body paragraphs. **Tables inside headers/footers are not editable** and are
reported as skipped elements.

> **Limitation:** Comments on header/footer changes are not attached to the
> output (Word and LibreOffice do not support comment ranges in those parts).
> They trigger a `UserWarning` and are dropped.

### Hyperlinks

Hyperlinks are extracted as `[link text](url)` inline within paragraph text.
Formatting inside links is preserved: `[**bold link**](url)`.

When modifying an existing paragraph, `[text]` without `(url)` **preserves**
the original hyperlink URL. `[text](new_url)` creates a new link.

When appending new text, `[text](url)` creates a hyperlink. `[text]` without
`(url)` produces plain text — always specify `(url)` on append if you want a
hyperlink.

### Tracked changes policy

Documents with pre-existing tracked changes (`<w:ins>`, `<w:del>`,
`<w:moveFrom>`, `<w:moveTo>`) are **hard-rejected** in both `extract_fragments`
and `apply_redlines`. Accept or reject all changes in Word before processing.

### `collapse_empty` mode

Optional mode that suppresses empty paragraphs from extraction and redlining.
Produces cleaner output for LLM consumption. When enabled, it **must** be
used consistently across extraction and redlining — mismatched values cause
fragment ID misalignment.

### Change types

#### Paragraph changes

| Type           | Description                                         | Requires `new_text` |
|----------------|-----------------------------------------------------|---------------------|
| `modify`       | Word-level diff applied as tracked changes          | Yes                 |
| `delete`       | Entire paragraph marked as deleted                  | No                  |
| `append_after` | New paragraph inserted after the referenced fragment | Yes                |

#### Table cell changes

| Type           | Description                                         | Requires `new_text` |
|----------------|-----------------------------------------------------|---------------------|
| `modify_cell`  | Modify cell content (single or multi-paragraph)     | Yes                 |
| `clear_cell`   | Delete all content in a cell (preserves structure)  | No                  |

**Cell modification** uses positional alignment: if the cell has multiple
paragraphs, the new text is split on newlines (`\n`) and each line is applied
to the corresponding paragraph in order. Cell content is marked with tracked
changes and comments just like paragraph modifications.

#### Blank line management

When appending new paragraphs, you can control surrounding blank lines:

```python
Change(
    fragment_id=10,
    change_type=ChangeType.APPEND_AFTER,
    new_text="New clause text here.",
    justification="Added new provision.",
    blank_lines_before=1,  # Insert 1 blank line before the new paragraph
    blank_lines_after=1,   # Insert 1 blank line after the new paragraph
)
```

When deleting paragraphs, you can remove trailing blank lines automatically:

```python
Change(
    fragment_id=15,
    change_type=ChangeType.DELETE,
    justification="Removed obsolete clause.",
    delete_next_blanks=1,  # Also delete the next blank paragraph
)
```

All blank lines are marked as tracked insertions/deletions and will appear in
the redlined document.

### Pseudo-Markdown

Text content uses a simplified Markdown-like format for inline formatting:

- `**bold**`
- `_italic_`
- `__underline__`

Unicode characters (smart quotes, em dashes, section symbols, non-breaking
spaces) are preserved as-is.

**Font inheritance**: When appending new paragraphs, the font family, size, and
color are automatically copied from the reference paragraph's first text-bearing
run. Bold, italic, and underline formatting from the pseudo-Markdown is layered
on top of the inherited base formatting.

### Changes JSON

The CLI accepts a JSON file containing either a bare array or a
`{"changes": [...]}` wrapper.

#### Paragraph changes example

```json
[
  {
    "fragment_id": "1",
    "change_type": "modify",
    "new_text": "The Seller agrees to deliver within **sixty** days.",
    "justification": "Extended delivery window."
  },
  {
    "fragment_id": "3",
    "change_type": "delete",
    "justification": "Removed governing law clause.",
    "delete_next_blanks": 1
  },
  {
    "fragment_id": "5",
    "change_type": "append_after",
    "new_text": "This Agreement shall be governed by Delaware law.",
    "justification": "Added Delaware governing law.",
    "blank_lines_before": 1,
    "blank_lines_after": 0
  },
  {
    "fragment_id": "header_1.1",
    "change_type": "modify",
    "new_text": "CONFIDENTIAL",
    "justification": "Updated header marking."
  }
]
```

#### Table cell changes example

```json
[
  {
    "cell_id": "2.1.1",
    "change_type": "modify_cell",
    "new_text": "Updated **cell** content",
    "justification": "Corrected cell value."
  },
  {
    "cell_id": "2.3.2",
    "change_type": "clear_cell",
    "justification": "Cleared obsolete data."
  }
]
```

Cell IDs use the format `"table_id.row.col"` where rows and columns are 1-based.

## Validation

The `validate_document()` function (and `docx-mcp validate` CLI) checks:

- **Annotation ID isolation** -- tracked-change and comment IDs don't collide
  across groups
- **Comment integrity** -- every `<w:comment>` has matching range markers in
  the document body, and vice versa
- **Tracked-change attributes** -- every `<w:ins>` and `<w:del>` has required
  `w:id`, `w:author`, and `w:date`
- **Package consistency** -- content-type and relationship entries exist for
  `comments.xml`

```python
from docx_mcp import validate_document

result = validate_document(doc)
if not result.ok:
    for error in result.errors:
        print(error)
```

## Architecture

The library manipulates OOXML directly via `lxml` (not `python-docx`) because
`python-docx` has no tracked-change support. Key design decisions:

- **Word-level diffing** via `diff-match-patch` with a word-to-char mapping
  for high-quality diffs
- **Conservative mutation** -- only changed paragraphs are touched; everything
  else passes through byte-identical
- **Globally unique annotation IDs** via a monotonic `IdManager` seeded from
  the document's existing max ID
- **`python-docx`** is used only for test fixture generation, not in the
  library itself

### Module map

```
src/docx_mcp/
  __init__.py        Public API
  cli.py             CLI entry point (apply, convert, validate)
  models.py          Pydantic data models (Change, ChangeType, RedlineConfig, ...)
  document.py        DocxDocument: ZIP parsing, XML tree access, serialization
  converter.py       Paragraph & table XML -> pseudo-Markdown conversion
  table_utils.py     Table inspection utilities (cell access, simplicity checks)
  tokenizer.py       Word-level tokenization
  differ.py          Word-level diff engine (diff-match-patch wrapper)
  run_ops.py         Diff-to-XML-run mapping, run splitting, element building
  id_manager.py      Monotonic annotation ID allocator
  comments.py        Comment creation and range marker insertion
  redliner.py        Main orchestrator: apply_redlines()
  table_redliner.py  Table cell change application
  audit.py           Document structural audit (headers, images, tables, etc.)
  validator.py       Structural validation checks
  server.py          MCP server (FastMCP 3.x, stdio transport)
  handlers/
    modify.py        Word-level tracked changes on existing paragraphs
    delete.py        Full paragraph deletion markup
    append.py        New paragraph insertion markup
```

## Development

```bash
# Run tests
uv run pytest tests/ -v

# Lint
uvx ruff check src/ tests/

# Auto-fix lint issues
uvx ruff check src/ tests/ --fix

# Type check
uvx ty check src/ tests/
```

431 tests covering all modules, handlers, table operations, headers/footers, hyperlinks, tracked-change rejection, merged-cell tables, section breaks, CLI, validation, and MCP server.

## License

MIT
