# Migration Guide: v0.1.0 → v0.2.0

## Overview

**v0.2.0** expands docx-mcp from a body-only, simple-table redlining engine to a
**full-document redlining engine** covering headers, footers, merged-cell tables,
hyperlinks, and section breaks.

This guide is written for consumers who **embed the library** and have **LLM
agents** that read the tagged text output and generate tracked changes. If you
use the MCP server tools exclusively, focus on sections 2–5 and 8. If you use
the Python API directly, also read sections 6–7.

### What changed at a glance

| Area | v0.1.0 | v0.2.0 |
|------|--------|--------|
| Fragment IDs | `int` (1, 2, 3…) | `str` (`"1"`, `"header_1.3"`, `"footer_2.1"`) |
| Document scope | Body only | Body + headers + footers |
| Tables | Simple rectangular only; merged = skipped | Merged cells supported with `span`/`vspan` markers |
| Hyperlinks | Invisible (silently dropped) | Rendered as `[text](url)` |
| Tracked changes | Ignored (risked corruption) | Hard-rejected with clear error |
| `extract_fragments` params | `format`, `markup` | Removed (always tagged) |
| `apply_changes` params | `validate` | Removed (always validates) |
| Python | 3.x | **3.14+** |
| FastMCP | 2.x | **3.x** |

---

## 1. New Document Layout

The tagged text output from `extract_fragments` is now richer. The LLM's
system prompt **must** be updated to understand the new format.

### 1.1 Old format (v0.1.0) — body only

```
<f=1>Introduction paragraph.</f=1>
<f=2>**Definitions.** The following terms shall apply.</f=2>
<table=3 rows=2 cols=3>
<cell=3.1.1>Header A</cell=3.1.1>
<cell=3.1.2>Header B</cell=3.1.2>
<cell=3.1.3>Header C</cell=3.1.3>
<cell=3.2.1>Data 1</cell=3.2.1>
<cell=3.2.2>Data 2</cell=3.2.2>
<cell=3.2.3>Data 3</cell=3.2.3>
</table=3>
<f=4>Closing paragraph.</f=4>
```

### 1.2 New format (v0.2.0) — full document

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
<f=header_1.2>Execution Version</f=header_1.2>
<f=footer_1.1>Page 1 of 10</f=footer_1.1>
```

Key differences:
- **Fragment IDs are strings**: body `"1"`, `"2"`; headers `"header_1.1"`;
  footers `"footer_1.1"`.
- **Tables show merge markers**: `span="2"` on cell `3.1.1` means it spans
  columns 1–2. Cell `3.1.2` is spanned-over and **omitted from output**.
- **Hyperlinks appear inline**: `[Section 2](https://example.com)`.
- **Headers and footers are included**: prefixed with `header_P.` or `footer_P.`.
- **Unprocessable tables** are shown as `<table=N skipped reason="..."/>`.

### 1.3 Understanding `span` and `vspan`

| Attribute | Meaning | Example |
|-----------|---------|---------|
| `span="2"` | Cell spans 2 columns (horizontal merge) | `<cell=3.1.1 span="2">` |
| `vspan="3"` | Cell spans 3 rows (vertical merge) | `<cell=4.2.1 vspan="3">` |
| No attribute | Standard single cell | `<cell=3.2.3>` |

**Spanned-over cells are omitted from the output.** If `cell=3.1.1` has
`span="2"`, then `cell=3.1.2` does not appear at all. This reduces noise by
~90% in documents with heavily merged tables.

### 1.4 Skipped tables

Tables that cannot be processed (nested tables, malformed merges) appear as:

```
<table=5 skipped reason="table 5, cell 2.3 contains nested table"/>
```

Tables inside headers or footers are always skipped and reported in
`skipped_elements`.

---

## 2. Fragment ID Reference

| ID Pattern | Where | Example Output | Target With |
|------------|-------|----------------|-------------|
| `"1"`, `"2"`, … | Body paragraphs | `<f=1>text</f=1>` | `fragment_id: "1"` |
| `"header_P.I"` | Header part P, paragraph I | `<f=header_1.3>text</f=header_1.3>` | `fragment_id: "header_1.3"` |
| `"footer_P.I"` | Footer part P, paragraph I | `<f=footer_2.1>text</f=footer_2.1>` | `fragment_id: "footer_2.1"` |
| `"N.r.c"` | Table N, row r, col c | `<cell=3.1.2>text</cell=3.1.2>` | `cell_id: "3.1.2"` |

**Important notes:**

- `P` (part index) is 1-based and corresponds to the document's header/footer
  relationships (usually header1.xml, footer1.xml).
- `I` (paragraph index) is 1-based within that header/footer part.
- Table IDs share the same ID space as body paragraphs (they interleave in
  document order). Fragment 3 might be a table, fragment 4 a paragraph.
- **Never target a spanned-over cell position.** If `cell=3.1.1` has `span="2"`,
  targeting `cell_id: "3.1.2"` will raise a `ValueError`.

---

## 3. Hyperlink Format

Hyperlinks are now extracted and preserved end-to-end.

### 3.1 Reading hyperlinks

In extracted text, hyperlinks appear inline with Markdown-like syntax:

```
See [the official website](https://example.com) for details.
```

Formatting inside links is preserved:

```
[**bold link text**](https://example.com)
```

### 3.2 Writing hyperlinks in changes

**On `modify` (existing paragraph):**

| `new_text` contains | Behavior |
|---------------------|----------|
| `[text](new_url)` | Hyperlink created with new URL |
| `[text]` (no URL) | **Original URL preserved** — existing relationship reused |
| Plain text | No hyperlink |

**On `append_after` (new paragraph):**

| `new_text` contains | Behavior |
|---------------------|----------|
| `[text](url)` | Hyperlink created with specified URL |
| `[text]` (no URL) | **Plain text** — no hyperlink created |

> **Asymmetry warning:** `[text]` without `(url)` has different semantics on
> modify vs append. On modify it preserves the existing link; on append it
> produces plain text. Always provide `(url)` on append if you want a hyperlink.

### 3.3 Deletion

Hyperlinks inside deleted paragraphs are correctly wrapped in `<w:del>`
tracked-change markup. No special handling needed.

---

## 4. MCP Tool API Changes

### 4.1 `extract_fragments`

**Before:**
```json
{
  "document_path": "/path/to/doc.docx",
  "format": "tagged",
  "markup": false
}
```

**After:**
```json
{
  "document_path": "/path/to/doc.docx"
}
```

- `format` removed — always returns tagged text.
- `markup` removed — tracked changes cause hard rejection, never extracted.
- Output now includes headers, footers, and tables interleaved with body paragraphs.
- Pre-existing tracked changes in the document cause a `ToolError`.

### 4.2 `apply_changes`

**Before:**
```json
{
  "document_path": "/path/to/doc.docx",
  "changes": [{"fragment_id": 3, "change_type": "modify", ...}],
  "validate": true
}
```

**After:**
```json
{
  "document_path": "/path/to/doc.docx",
  "changes": [{"fragment_id": "3", "change_type": "modify", ...}]
}
```

- `fragment_id` is now a `str` — JSON integers are auto-coerced (`3` → `"3"`).
- `validate` removed — always validates internally.
- Header/footer fragment IDs (`"header_1.1"`) are accepted.

### 4.3 `apply_changes_from_file`

Same changes as `apply_changes`: `validate` removed, `fragment_id` is now `str`.

### 4.4 `diff_fragments`

Now compares headers and footers in addition to body content. Fragment IDs
in diff output use the new string format.

### 4.5 New: `audit_document_tool`

```json
{
  "document_path": "/path/to/doc.docx",
  "format": "text"
}
```

Returns a structural audit: header/footer counts, image counts, table
classifications, section breaks, tracked changes, comments, and unsupported
elements (footnotes, endnotes, text boxes). Use this as a pre-flight check
before extraction to understand what the document contains.

---

## 5. Change Object Formats

The library has two change representations. Which one you use depends on
whether you're going through the MCP tools or the Python API directly.

### 5.1 MCP / JSON format (`ChangeParam`)

This is the format used in MCP tool calls and JSON files. It uses **field-based
discrimination** — the engine auto-detects the change type from which fields
are present.

**Paragraph changes:** use `fragment_id` (required), `change_type`, `new_text`,
`justification`, plus optional spacing controls.

```json
{
  "fragment_id": "3",
  "change_type": "modify",
  "new_text": "The Company **shall** provide written notice.",
  "justification": "Strengthened obligation language."
}
```

```json
{
  "fragment_id": "15",
  "change_type": "delete",
  "justification": "Removed obsolete clause.",
  "delete_next_blanks": 1
}
```

```json
{
  "fragment_id": "20",
  "change_type": "append_after",
  "new_text": "**16. Governing Law.** This Agreement shall be governed by Delaware law.",
  "justification": "Added Delaware governing law.",
  "blank_lines_before": 1,
  "blank_lines_after": 1
}
```

```json
{
  "fragment_id": "header_1.1",
  "change_type": "modify",
  "new_text": "CONFIDENTIAL",
  "justification": "Updated header text."
}
```

**Table cell changes:** use `cell_id` (required), `change_type`, `new_text`,
`justification`.

```json
{
  "cell_id": "2.1.1",
  "change_type": "modify_cell",
  "new_text": "Updated **header** text",
  "justification": "Corrected table header."
}
```

```json
{
  "cell_id": "3.2.3",
  "change_type": "clear_cell",
  "justification": "Removed obsolete data."
}
```

### 5.2 Python API format (`Change`)

The Python API uses explicit `kind` discrimination with separate
`ParagraphChange` and `TableChange` models.

```python
from docx_mcp import (
    ParagraphChange,
    ParagraphChangeType,
    TableChange,
    TableChangeType,
    apply_redlines,
)

changes = [
    # Body paragraph modify
    ParagraphChange(
        kind="paragraph",
        fragment_id="3",            # ← now a str
        change_type=ParagraphChangeType.MODIFY,
        new_text="Updated text.",
        justification="Reason.",
    ),
    # Header paragraph modify
    ParagraphChange(
        kind="paragraph",
        fragment_id="header_1.1",   # ← string with prefix
        change_type=ParagraphChangeType.MODIFY,
        new_text="CONFIDENTIAL",
        justification="Updated header.",
    ),
    # Table cell modify
    TableChange(
        kind="table",
        table_id=2,                  # ← separate integer fields
        row=1,
        col=1,
        change_type=TableChangeType.MODIFY_CELL,
        new_text="Updated cell.",
        justification="Reason.",
    ),
]

doc = apply_redlines("input.docx", changes)
doc.save("output.docx")
```

> **Note:** `TableChange` uses `table_id`/`row`/`col` (integers), not a
> `cell_id` string. The `cell_id` property is computed and is not serialized.

### 5.3 For LLM Structured Outputs

If your LLM generates changes via structured output (JSON Schema), the
recommended approach is to use a schema with an **explicit discriminator**:

```json
{
  "type": "object",
  "oneOf": [
    {
      "properties": {
        "kind": {"const": "paragraph"},
        "fragment_id": {"type": "string"},
        "change_type": {"enum": ["modify", "delete", "append_after"]},
        "new_text": {"type": ["string", "null"]},
        "justification": {"type": "string"}
      },
      "required": ["kind", "fragment_id", "change_type", "justification"]
    },
    {
      "properties": {
        "kind": {"const": "table"},
        "table_id": {"type": "integer"},
        "row": {"type": "integer"},
        "col": {"type": "integer"},
        "change_type": {"enum": ["modify_cell", "clear_cell"]},
        "new_text": {"type": ["string", "null"]},
        "justification": {"type": "string"}
      },
      "required": ["kind", "table_id", "row", "col", "change_type", "justification"]
    }
  ]
}
```

The MCP server's `ChangeParam` union also works for structured outputs but
uses implicit discrimination (by `change_type` value or field presence) rather
than an explicit `kind` field. Both are valid; choose based on your LLM's
structured output requirements.

---

## 6. Python API Changes

### 6.1 `fragment_id` type change

`ParagraphChange.fragment_id` is now `str` instead of `int`.

```python
# BEFORE (v0.1.0)
ParagraphChange(fragment_id=3, ...)

# AFTER (v0.2.0)
ParagraphChange(fragment_id="3", ...)
```

Pydantic's `coerce_numbers_to_str` is enabled, so `fragment_id=3` still works
in Python. In JSON, `"fragment_id": 3` is auto-coerced to `"3"`.

### 6.2 New functions

| Function | Replaces | Notes |
|----------|----------|-------|
| `full_to_fragments(doc)` | `body_to_fragments(doc.body_elements)` | Includes headers/footers |
| `full_element_map(doc)` | `doc.fragment_map()` / `doc.interleaved_element_map()` | Returns `dict[str, Element]` |
| `fragments_to_tagged_text_interleaved(items)` | `fragments_to_tagged_text(fragments)` | Handles mixed paragraph/table items |
| `fragments_to_json_interleaved(items)` | manual JSON construction | Handles mixed paragraph/table items |
| `audit_document(doc)` | — | New: structural audit |

### 6.3 `apply_redlines()` new parameter

```python
def apply_redlines(
    source: Path | str | bytes,
    changes: Sequence[Change],
    config: RedlineConfig | None = None,
    *,
    collapse_empty: bool = False,  # ← NEW
) -> DocxDocument:
```

### 6.4 New exports

```python
from docx_mcp import (
    AuditReport,                          # NEW
    CellInfo,                             # NEW
    FragmentResult,                       # NEW
    SkippedTableInfo,                     # NEW
    TableInfo,                            # NEW
    audit_document,                       # NEW
    body_to_fragments,                    # NEW
    fragments_to_json_interleaved,        # NEW
    fragments_to_tagged_text_interleaved, # NEW
    full_to_fragments,                    # NEW
)
```

---

## 7. `collapse_empty` Mode

A new optional mode that suppresses empty paragraphs from both extraction and
redlining. This produces cleaner output for LLM consumption.

**Critical contract:** `collapse_empty` must use the **same value** for
extraction and redlining. If you extract with `collapse_empty=True` but redline
with `collapse_empty=False` (the default), fragment IDs will be misaligned and
changes will silently target the wrong paragraphs.

```python
from docx_mcp import full_to_fragments, apply_redlines

# Extraction
result = full_to_fragments(doc, collapse_empty=True)

# Redlining — MUST match
doc = apply_redlines("input.docx", changes, collapse_empty=True)
```

In the MCP server, `collapse_empty` is not yet exposed as a tool parameter
(it defaults to `False`). If you need it, use the Python API directly.

---

## 8. System Prompt Updates

If your LLM agents have system prompts describing the tagged format, update
them with these changes:

### 8.1 Fragment ID description

```
BEFORE: Fragment IDs are 1-based paragraph indices. Each fragment is a
        top-level paragraph in the document body.

AFTER:  Fragment IDs are strings identifying paragraphs and tables in document
        order. Body paragraphs use plain numbers ("1", "2", ...). Headers use
        "header_<part>.<paragraph>" (e.g., "header_1.1"). Footers use
        "footer_<part>.<paragraph>" (e.g., "footer_2.1"). Tables share the same
        ID space as body paragraphs.
```

### 8.2 Tagged format description

```
BEFORE: <f=N>text</f=N> for paragraphs.
        <table=N rows=R cols=C>...</table=N> for simple tables.
        <cell=N.r.c>text</cell=N.r.c> for table cells.

AFTER:  <f=ID>text</f=ID> for paragraphs (body, header, footer).
        <table=N rows=R cols=C>...</table=N> for tables.
        <cell=N.r.c span="S">text</cell=N.r.c> for table cells.
        Span>1 means horizontal merge; vspan>1 means vertical merge.
        Spanned-over cells are omitted from output.
        Hyperlinks appear as [text](url) inline within paragraph text.
```

### 8.3 Hyperlinks

```
NEW:    Hyperlinks in extracted text use [link text](url) format.
        When modifying text containing a hyperlink, [text] without (url)
        preserves the existing URL. [text](new_url) replaces it.
        When appending new text, [text](url) creates a hyperlink; [text]
        without (url) produces plain text.
```

### 8.4 Known limitations to communicate

```
- Comments on header/footer changes are not attached (Word limitation).
- Documents with pre-existing tracked changes are rejected.
- Nested tables, images, footnotes, and text boxes are not extractable.
- Tables inside headers/footers are not editable.
- Never target spanned-over cells (cells covered by a span/vspan merge).
  Target the originating cell instead.
```

---

## 9. Migration Checklist

- [ ] **1. Update Python to 3.14+**
- [ ] **2. Update fragment_id handling** — change `int` to `str` in all code
      that constructs `ParagraphChange`. JSON integers are auto-coerced.
- [ ] **3. Remove `format` and `markup` params** from `extract_fragments` calls.
- [ ] **4. Remove `validate` param** from `apply_changes` /
      `apply_changes_from_file` calls.
- [ ] **5. Update system prompts** — new fragment ID patterns, hyperlink format,
      table span/vspan markers, header/footer awareness (see Section 8).
- [ ] **6. Handle header/footer fragments** — LLM may now see
      `header_*` / `footer_*` IDs and may target them with changes.
- [ ] **7. Handle merged-cell tables** — LLM must understand `span`/`vspan`
      markers and never target spanned-over positions.
- [ ] **8. Update table cell citation format** — `cell_id` strings remain
      `"N.r.c"` but must reference originating cells only.
- [ ] **9. Optionally adopt `collapse_empty`** — use consistently across
      extraction and redlining.
- [ ] **10. Optionally adopt `audit_document`** — pre-flight check before
       extraction to understand document structure.
- [ ] **11. Update imports** — if using new functions (`full_to_fragments`,
       `AuditReport`, etc.), update import paths.
- [ ] **12. Update pyproject.toml / dependencies** — FastMCP 3.x.

---

## 10. Known Limitations & Gotchas

### Comments on header/footer changes

Comments on header/footer changes trigger a Python `UserWarning` and are
**not attached** to the output document. Word and LibreOffice do not support
`commentRangeStart`/`commentRangeEnd` in header/footer XML parts.

### Tracked changes hard-reject

Documents containing any pre-existing `<w:ins>`, `<w:del>`, `<w:moveFrom>`,
or `<w:moveTo>` elements (in body, headers, footers, or comments) are
**rejected in both `extract_fragments` and `apply_redlines`**. Accept or
reject all tracked changes in Word before processing.

### CLI `convert` is body-only

The `docx-mcp convert` CLI command still uses the old body-only extraction
path. It does not include headers, footers, or tables, and produces integer
fragment IDs. For full-document extraction, use the MCP `extract_fragments`
tool or the Python `full_to_fragments()` function.

### `collapse_empty` consistency

If you use `collapse_empty=True` during extraction, you **must** pass
`collapse_empty=True` to `apply_redlines()`. Mismatched values cause
fragment ID misalignment — changes silently target the wrong paragraphs.

### Spanned-over cell targeting

Targeting a cell position that is covered by a merge (`span` or `vspan`)
raises a `ValueError` with guidance to target the originating cell. The
LLM must understand that if `cell=3.1.1` has `span="2"`, cell position
`3.1.2` does not exist in the output and cannot be targeted.

### Tables in headers/footers

Tables inside headers or footers are reported as `skipped_elements` and
are not editable. Only `<w:p>` elements in headers/footers are extractable.

### Formatting-only changes produce no diff

The engine converts pseudo-Markdown to raw text before diffing. Formatting
markers (`**`, `_`, `__`) are stripped, not diffed. To make text bold, the
`new_text` must differ in word content from the original, or the diff will
produce no tracked change. In practice this means: to make existing text bold
via `modify`, you must rephrase the text.

### Nested tables not supported

Tables containing nested `<w:tbl>` elements are skipped with a reason.

### Images, footnotes, endnotes, text boxes

These are not extractable. Images and tables-in-headers/footers are reported
in `skipped_elements`. Footnotes, endnotes, and text boxes are listed in the
`audit_document` report.

### `diff_fragments` matches by position

Fragment comparison is positional (fragment 1 vs fragment 1), not
content-based. Documents with different structures will show extensive diffs.

### `TableChange.cell_id` is computed

In the Python API, `TableChange.cell_id` is a `@property` and does not appear
in serialized JSON output. Use `table_id`/`row`/`col` fields for serialization.

### `__init__.py` imports

All new public types are now exported from `docx_mcp`:
`AuditReport`, `CellInfo`, `FragmentResult`, `SkippedTableInfo`, `TableInfo`,
`audit_document`, `body_to_fragments`, `fragments_to_json_interleaved`,
`fragments_to_tagged_text_interleaved`, `full_to_fragments`.

---

## Appendix A: Full Tagged Format Specification

```
<document>    ::= <fragment>*

<fragment>    ::= <paragraph>
                | <table>
                | <skipped_table>

<paragraph>   ::= "<f=" <fragment_id> ">" <text> "</f=" <fragment_id> ">"

<table>       ::= "<table=" <table_id> " rows=" <rows> " cols=" <cols> ">"
                  <cell>*
                  "</table=" <table_id> ">"

<skipped_table> ::= "<table=" <table_id> ' skipped reason="' <reason> '"/>'

<cell>        ::= "<cell=" <cell_id> [<cell_attrs>] ">" <text> "</cell=" <cell_id> ">"

<cell_attrs>  ::= ' span="' <span> '"'
                | ' vspan="' <vspan> '"'
                | ' span="' <span> '" vspan="' <vspan> '"'

<fragment_id> ::= <body_id> | <header_id> | <footer_id>
<body_id>     ::= "1" | "2" | "3" | ...
<header_id>   ::= "header_" <part_num> "." <para_num>
<footer_id>   ::= "footer_" <part_num> "." <para_num>
<cell_id>     ::= <table_id> "." <row_num> "." <col_num>

<table_id>    ::= <body_id>
<part_num>    ::= "1" | "2" | ...
<para_num>    ::= "1" | "2" | ...
<row_num>     ::= "1" | "2" | ...
<col_num>     ::= "1" | "2" | ...
<rows>        ::= positive integer
<cols>        ::= positive integer
<span>        ::= positive integer (horizontal merge count)
<vspan>       ::= positive integer (vertical merge count)

<text>        ::= pseudo-Markdown text with:
                  - **bold**, _italic_, __underline__
                  - [link text](url) for hyperlinks
                  - [link text] for hyperlinks with preserved URLs
                  - \n for paragraph breaks within table cells
                  - \\*, \\_, \\\\ for literal *, _, \

<reason>      ::= human-readable explanation (e.g., "contains nested table")
```

## Appendix B: JSON Output Format

If you use `fragments_to_json_interleaved()`, the JSON structure mirrors the
tagged format:

```json
[
  {"type": "paragraph", "fragment_id": "1", "text": "Introduction."},
  {"type": "table", "table_id": 2, "rows": 2, "cols": 3,
   "cells": [[
     {"cell_id": "2.1.1", "row": 1, "col": 1, "span": 2, "text": "Merged Header"},
     {"cell_id": "2.1.3", "row": 1, "col": 3, "text": "Header C"}
   ], [
     {"cell_id": "2.2.1", "row": 2, "col": 1, "text": "Data 1"},
     {"cell_id": "2.2.2", "row": 2, "col": 2, "text": "Data 2"},
     {"cell_id": "2.2.3", "row": 2, "col": 3, "text": "Data 3"}
   ]]},
  {"type": "paragraph", "fragment_id": "header_1.1", "text": "Confidential"},
  {"type": "table", "table_id": 4, "skipped": true,
   "reason": "table 4, cell 2.3 contains nested table"}
]
```

Note: `span` and `vspan` are only present when ≠ 1. Spanned-over cells
(`span=0`, `vspan=0`) are omitted from the JSON output.
