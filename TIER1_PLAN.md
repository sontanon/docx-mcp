# Tier 1 Implementation Plan

Branch: `feat/tier1-foundation`
Parent: `main`

## Order of Implementation

1. **T1.1** Hard-reject pre-existing tracked changes
2. **T1.3** Fix `append_after` bugs
3. **T1.2** Hyperlink preservation & editing
4. **T1.4** Lossiness infrastructure

## Design Decisions (Locked)

| # | Decision | Locked Value |
|---|----------|--------------|
| 1 | Tracked-change rejection scope | `ins`, `del`, `moveFrom`, `moveTo` |
| 2 | Multiple-append granularity | Individual `w:id` per paragraph |
| 3 | Hyperlink URL change strategy | Create new relationship (don't mutate shared `r:id`) |
| 4 | JSON extraction shape | Object `{fragments, skipped_elements}` — breaking change |
| 5 | Audit output format | Both text and JSON with `--format` flag |

## Testing Strategy

- **Fixtures:** Programmatic generation via `tests/generators/generate_fixtures.py`.
- **New fixtures:**
  - `body_tracked_changes.docx` — `<w:ins>` in body paragraph
  - `header_tracked_changes.docx` — `<w:del>` in header paragraph
  - `footer_tracked_changes.docx` — `<w:moveFrom>` in footer paragraph
  - `comments_tracked_changes.docx` — `<w:moveTo>` in comments.xml
  - `hyperlink_paragraph.docx` — plain hyperlink
  - `hyperlink_formatted.docx` — hyperlink with bold/italic inside
  - `multiple_hyperlinks.docx` — paragraph with two hyperlinks
  - `document_with_image.docx` — paragraph + inline image
  - `nested_table.docx` — table cell containing nested table
  - `multi_column.docx` — two-column layout
- **New test files:**
  - `tests/test_tracked_change_rejection.py`
  - `tests/test_append_fixes.py`
  - `tests/test_hyperlinks.py`
  - `tests/test_lossiness.py`
- **Fixture regeneration:** Delete `tests/fixtures/generated/` once to force regeneration of all fixtures including new ones.

## Progress

- [x] ROADMAP.md committed to main
- [x] Branch `feat/tier1-foundation` created
- [x] T1.1 Implementation
- [x] T1.3 Implementation
- [x] T1.2 Implementation
- [x] T1.4 Implementation
- [x] All tests passing (414 passed, 1 skipped)
- [x] Lint/type-check clean
