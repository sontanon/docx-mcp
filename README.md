# docx-mcp

Legal document redlining engine. Takes AI-generated changes (structured JSON)
and applies them as professional tracked changes with comments inside `.docx`
files. The output is indistinguishable from what a lawyer would produce in
Microsoft Word -- proper `w:ins`/`w:del` markup, comment annotations with
justification text, and preserved formatting.

## Installation

Requires Python 3.13+.

```bash
uv sync
```

## Quick start

### Python API

```python
from docx_mcp import apply_redlines, Change, ChangeType, RedlineConfig

changes = [
    Change(
        fragment_id=3,
        change_type=ChangeType.MODIFY,
        new_text="The Company **shall** provide written notice.",
        justification="Strengthened obligation language.",
    ),
    Change(
        fragment_id=5,
        change_type=ChangeType.DELETE,
        justification="Removed redundant clause.",
    ),
    Change(
        fragment_id=7,
        change_type=ChangeType.APPEND_AFTER,
        new_text="The foregoing shall survive termination.",
        justification="Added survival provision.",
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
```

## Concepts

### Fragments

Paragraphs in the document are numbered 1..N (top-level `<w:p>` elements in
`<w:body>`). Each paragraph is a **fragment**, identified by its 1-based index.
Use `docx-mcp convert` to see the fragment map for any document.

### Change types

| Type           | Description                                         | Requires `new_text` |
|----------------|-----------------------------------------------------|---------------------|
| `modify`       | Word-level diff applied as tracked changes          | Yes                 |
| `delete`       | Entire paragraph marked as deleted                  | No                  |
| `append_after` | New paragraph inserted after the referenced fragment | Yes                |

### Pseudo-Markdown

Text content uses a simplified Markdown-like format for inline formatting:

- `**bold**`
- `_italic_`
- `__underline__`

Unicode characters (smart quotes, em dashes, section symbols, non-breaking
spaces) are preserved as-is.

### Changes JSON

The CLI accepts a JSON file containing either a bare array or a
`{"changes": [...]}` wrapper:

```json
[
  {
    "fragment_id": 1,
    "change_type": "modify",
    "new_text": "The Seller agrees to deliver within **sixty** days.",
    "justification": "Extended delivery window."
  },
  {
    "fragment_id": 3,
    "change_type": "delete",
    "justification": "Removed governing law clause."
  },
  {
    "fragment_id": 5,
    "change_type": "append_after",
    "new_text": "This Agreement shall be governed by Delaware law.",
    "justification": "Added Delaware governing law."
  }
]
```

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
  converter.py       Paragraph XML -> pseudo-Markdown conversion
  tokenizer.py       Word-level tokenization
  differ.py          Word-level diff engine (diff-match-patch wrapper)
  run_ops.py         Diff-to-XML-run mapping, run splitting, element building
  id_manager.py      Monotonic annotation ID allocator
  comments.py        Comment creation and range marker insertion
  redliner.py        Main orchestrator: apply_redlines()
  validator.py       Structural validation checks
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
uv run ruff check src/ tests/

# Auto-fix lint issues
uv run ruff check src/ tests/ --fix
```

206 tests covering all modules, handlers, CLI, and validation.

## License

TBD
