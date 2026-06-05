# Changelog

## v0.2.0 (2026-05-13)

### Breaking Changes

- **`fragment_id` is now `str`** instead of `int`. Body paragraphs use `"1"`,
  `"2"`, etc. Headers and footers use prefixed IDs (`"header_1.3"`,
  `"footer_2.1"`). JSON integers are auto-coerced for backward compatibility.
- **`extract_fragments` removed `format` and `markup` parameters.** Always
  returns tagged text. Tracked changes now cause hard rejection rather than
  being extracted as markup.
- **`apply_changes` and `apply_changes_from_file` removed `validate` parameter.**
  Validation now always runs internally.
- **Python 3.14+ required.** FastMCP upgraded to 3.x.

### New Features

- **Header/footer extraction and redlining** — Headers and footers are now
  extracted with prefixed fragment IDs (`header_1.1`, `footer_2.1`) and fully
  support modify, delete, and append_after operations.
- **Merged-cell table support** — Tables with `gridSpan` (horizontal merges)
  and `vMerge` (vertical merges) are now fully supported. Spanned-over cells
  are omitted from extraction output. Span/vspan markers shown in tagged and
  JSON formats.
- **Hyperlink preservation and editing** — Hyperlinks are extracted as
  `[text](url)` pseudo-Markdown, preserved on modify, and creatable on append.
- **Section break preservation** — `append_after` at the end of a section
  correctly restores `<w:sectPr>` so new paragraphs stay in the same section.
- **Tracked changes hard-reject** — Documents with pre-existing tracked changes
  are rejected with a clear error message listing affected parts.
- **`audit_document` tool and CLI command** — Structural diagnostics reporting
  headers, footers, images, tables, section breaks, tracked changes, comments,
  and unsupported elements.
- **`collapse_empty` mode** — Optional suppression of empty paragraphs during
  extraction and redlining (default `False`).
- **Document corpus** — 31 real-world legal, policy, and business documents
  with automated quality analysis.

### Improvements

- **Extraction quality** — Whitespace-only runs no longer produce formatting
  artifacts (`****`, `____`). Spanned-over cells omitted from output (~90%
  noise reduction in merged-cell tables).
- **Table robustness** — Empty cells and wide tables (10+ columns) handled
  correctly. Error messages include table ID for easier debugging.
- **Validation always runs** — Validation is no longer optional; every
  `apply_changes` call validates the output.
- **`diff_fragments` now includes headers and footers** in comparison.
- **New public exports** — `AuditReport`, `CellInfo`, `FragmentResult`,
  `SkippedTableInfo`, `TableInfo`, `audit_document`, `body_to_fragments`,
  `fragments_to_json_interleaved`, `fragments_to_tagged_text_interleaved`,
  `full_to_fragments`.

### Fixes

- **Comments on header/footer changes** — Comments are skipped on header/footer
  changes with a `UserWarning` to prevent LibreOffice corruption.
- **Append order** — Multiple appends to the same fragment now produce the
  correct document order.
- **Annotation ID sharing** — Appended paragraph content and paragraph mark now
  share the same `w:id`.
- **Stale docstrings** — MCP server instructions and tool docstrings updated
  for the new fragment ID format and parameter changes.

### Known Limitations

- Comments on header/footer changes are dropped (Word/LibreOffice limitation).
- CLI `convert` command is body-only (does not include headers, footers, or
  tables). Use MCP `extract_fragments` or Python `full_to_fragments()` for
  full-document extraction.
- Tables inside headers/footers are not editable.
- Nested tables, images, footnotes, endnotes, and text boxes are not
  extractable.
- `collapse_empty` must use the same value for extraction and redlining.
- Formatting-only changes (e.g., making text bold without changing words)
  produce no diff.

---

## v0.1.0 (2025-08)

- Initial release: legal document redlining engine.
- Body-only paragraph and simple-table redlining.
- Word-level diffing via diff-match-patch.
- MCP server with 5 tools (extract, apply, apply_from_file, validate, diff).
- CLI with apply, convert, validate commands.
- Pseudo-Markdown formatting (**bold**, _italic_, __underline__).
- Blank line management and font inheritance.
- Structural validation (annotation IDs, comment integrity, package consistency).
- 389 tests.
