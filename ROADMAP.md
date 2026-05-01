# docx-mcp Roadmap

A living design document for the next round of improvements to the docx-mcp legal-document redlining engine.

## Mission

Make docx-mcp the most reliable, transparent, and LLM-friendly engine for applying tracked changes to legal .docx files — starting with NDAs, Master Agreements, and general transactional documents.

## Design Principles

1. **Transparency over magic.** The LLM (and the human reviewer) must always know what the engine *cannot* see or change. If we skip content, we say so.
2. **Pristine-input policy.** We only process clean documents. Pre-existing tracked changes are rejected; existing comments are preserved but their IDs are guarded.
3. **Progressive enhancement.** Simple cases must work perfectly before we tackle complex ones. A skipped table with a clear reason is better than a corrupted table.
4. **LLM-native APIs.** Fragment IDs, pseudo-Markdown, and change schemas should be easy for an LLM to generate and for a human to read.

---

## Locked Design Decisions

These were reached via structured Q&A and are not open for re-litigation without a compelling new constraint.

| Topic | Decision | Rationale |
|-------|----------|-----------|
| **Hyperlinks** | Full URL editing via `[text](url)` pseudo-Markdown. | LLMs can rewrite both link text and destinations. |
| **Hyperlink URL preservation (modify)** | `[text]` without `(url)` auto-preserves the original URL. | Prevents accidental link stripping when the LLM only wants to rephrase link text. |
| **Hyperlink URL requirement (append)** | New links in appended text must specify `(url)`. | Explicit is safe; no invisible auto-linking. |
| **Headers / Footers** | Namespace-prefixed fragment IDs: `header_1.3`, `footer_2.1`. | Keeps a flat ID space the LLM already understands; part index is explicit. |
| **Lossiness reporting** | Inline `skipped_elements` summary in `extract_fragments` output **(JSON only)** plus a dedicated `audit_document` tool. | Tagged format stays purely textual; JSON gets structured metadata. |
| **Pre-existing tracked changes** | **Hard reject** in both `extract_fragments` and `apply_changes`. | Prevents annotation ID collisions, confusing diff baselines, and corrupted output. |
| **Existing comments** | **Preserve and avoid collisions.** | `comments.xml` is already loaded; `IdManager` scans it. Existing comments remain untouched. |
| **Images** | Listed in lossiness report as `image` type. No image manipulation support. | Text-only engine; images are expected to be invisible but we warn about them. |
| **Tables (merged cells)** | Supported via logical grid builder (`build_table_grid`). Horizontal (`gridSpan`) and vertical (`vMerge`) merges render with `span` / `vspan` markers. Spanned-over cells are shown as empty with `span="0"`. | Previously deferred to Tier 3; now implemented after corpus analysis showed 40% of documents contained merged-cell tables. |
| **Section breaks / multi-column** | Defer explicit handling until multi-column test fixtures are added and validated. | Current implicit behavior (new paragraphs inherit preceding section properties) is likely correct. |
| **Signatory blocks** | Supported as both paragraph-based and simple-table-based. Merged-cell signature tables skipped with clear reason until Tier 3. | Both formats appear in real documents. |

---

## Tier 1: Fix Critical Bugs & Foundation

These must be completed before any production use on documents containing hyperlinks, multiple appends, or headers/footers.

### T1.1 Hard-reject pre-existing tracked changes

**Problem:** Documents with pending tracked changes (`<w:ins>`, `<w:del>`, `<w:moveFrom>`, `<w:moveTo>`) cause annotation ID collisions and garbled output.

**Decision:** Reject in both `extract_fragments` and `apply_changes` with a clear error message listing where tracked changes were found (body, header, footer, or comments).

**Implementation notes:**
- Scan `document.xml`, all `word/header*.xml`, all `word/footer*.xml`, and `comments.xml`.
- Look for any element whose local tag is `ins`, `del`, `moveFrom`, or `moveTo`.
- Error message: `Document contains pre-existing tracked changes in {location}. Please accept or reject all changes before redlining.`

**Tests needed:**
- Fixture with a body `<w:ins>` → rejected.
- Fixture with a header `<w:del>` → rejected.
- Pristine document → accepted.

---

### T1.2 Hyperlink preservation & editing

**Problem:** Hyperlinks are completely invisible to the pipeline. Text inside `<w:hyperlink>` is silently dropped during extraction, diffing misaligns, and deletion leaves hyperlink text visible.

**Decision:** Support hyperlinks end-to-end: extraction, modification, deletion, and creation via `[text](url)` pseudo-Markdown.

**Implementation notes:**

1. **Extraction (`converter.py`):**
   - `paragraph_to_pseudo_markdown()` must descend into `<w:hyperlink>` children.
   - Render as `[text](url)` where `url` is resolved from the relationship ID (`r:id` → `word/_rels/document.xml.rels` → target URL).
   - Inline formatting inside links (bold, italic) must be preserved: `[**bold text**](url)`.

2. **Run extraction (`run_ops.py`):**
   - `extract_runs()` must find `w:r` inside `w:hyperlink`, not just direct children.
   - Each extracted run should carry a `hyperlink_rel_id` field (or `None` for plain runs).

3. **Modification (`handlers/modify.py`):**
   - During rebuild, runs with the same `hyperlink_rel_id` must be grouped back into a `<w:hyperlink>` wrapper.
   - If the LLM provides `[text]` without `(url)` in `new_text`, preserve the original `r:id`.
   - If the LLM provides `[text](new_url)`, create a new relationship in `word/_rels/document.xml.rels` and assign a new `r:id`.
   - If a hyperlink is split by a diff boundary (some text deleted, some kept), the kept portion retains the wrapper; the inserted portion gets a new wrapper with the resolved/new URL.

4. **Deletion (`handlers/delete.py`):**
   - Must wrap runs inside `<w:hyperlink>` in `<w:del>`, just like plain runs.
   - The `<w:hyperlink>` wrapper itself should remain outside the `<w:del>` (or be deleted entirely if all its runs are deleted — Word accepts either; test both).

5. **Append (`handlers/append.py`):**
   - Extend pseudo-Markdown parser to recognize `[text](url)`.
   - Build `<w:hyperlink r:id="...">` wrappers around the corresponding runs.
   - Require explicit `(url)`; `[text]` without URL becomes plain text (no auto-linking).

6. **Relationship management (`document.py`):**
   - `DocxDocument` must load `word/_rels/document.xml.rels` (already loaded as `_rels_tree`).
   - Add helper to resolve `r:id` → target URL.
   - Add helper to create a new relationship entry and return a new `r:id`.
   - Ensure `to_bytes()` serializes the updated rels tree.

**Tests needed:**
- Extract paragraph with hyperlink → `[text](url)` in output.
- Modify paragraph with hyperlink, keep URL implicit → hyperlink preserved.
- Modify paragraph with hyperlink, change URL → new relationship created.
- Delete paragraph with hyperlink → hyperlink text marked as deleted.
- Append paragraph with `[text](url)` → hyperlink created in output.
- Hyperlink with bold/italic inside → formatting preserved.

---

### T1.3 Fix `append_after` bugs

**Problem A: Multiple appends to the same fragment reverse insertion order.**

Because `handle_append_after` uses `addnext` on the *original* reference paragraph every time, appending to fragment 5 twice results in the second append appearing *before* the first in the document.

**Fix:** In `redliner.py`, when sorting paragraph changes, group `APPEND_AFTER` changes by `fragment_id` and process them in reverse order *within* each group. Alternatively, pass the previously appended paragraph (or last blank) as the new reference for subsequent appends to the same fragment.

**Problem B: Content wrapper uses a different annotation ID from the paragraph mark.**

`_build_inserted_paragraph` receives `ins_id` for the paragraph mark `<w:ins>`, but calls `id_manager.next_id()` for the content wrapper `<w:ins>`. This causes Word to show two separate tracked changes for one paragraph insertion.

**Fix:** Use the same `ins_id` for both the paragraph mark and the content wrapper, consistent with `handle_delete`.

**Problem C: Stale comment in `append.py`.**

The docstring says "we insert blanks_after first … then content, then blanks_before" but the code does the opposite. Fix the comment.

**Tests needed:**
- Two appends to fragment 5 → verify text order in XML.
- Single append → verify paragraph mark and content share the same `w:id`.

---

### T1.4 Lossiness infrastructure

**Problem:** The engine silently skips content (images, merged tables, nested content) and the LLM has no way to know what it missed.

**Decision:** Two-pronged approach.

**A. Inline summary in `extract_fragments` (JSON only):**

When `format="json"`, each top-level element in the returned array gets a `skipped_elements` sibling array at the root:

```json
{
  "fragments": [...],
  "skipped_elements": [
    {"type": "image", "location": "header_1", "description": "logo.png"},
    {"type": "merged_table", "table_id": 4, "reason": "cell 1.2 has horizontal merge (gridSpan)"},
    {"type": "nested_table", "table_id": 5, "reason": "contains nested table"}
  ]
}
```

Tagged format remains purely textual; no skipped items are injected.

**B. New `audit_document` tool / CLI subcommand:**

A standalone diagnostic that reports structural findings without extracting text:

```bash
docx-mcp audit contract.docx
```

Reports:
- Number of headers / footers / header parts / footer parts.
- Number of images (by counting `<w:drawing>` / `<w:pict>` in body, headers, footers).
- Number of simple vs. skipped tables with reasons.
- Presence of section breaks, multi-column layouts.
- Presence of tracked changes or comments.
- Any unsupported elements (footnotes, endnotes, text boxes).

**Tests needed:**
- JSON extraction returns `skipped_elements` for merged table and image.
- `audit_document` on NDA fixture reports 0 skipped; on complex fixture reports multiple.

---

## Tier 2: Expand Coverage

### T2.1 Header/footer extraction & redlining **[DONE]**

**Problem:** Headers and footers are invisible. Any text that needs redlining in a header/footer is unreachable.

**Decision:** Load header/footer XML parts and expose their paragraphs via prefixed fragment IDs. Tables inside headers/footers are extracted as skipped elements, not editable.

**Implementation notes:**

1. **Loading (`document.py`):**
   - Parse `word/_rels/document.xml.rels` to discover header/footer relationships.
   - Load each referenced XML part into `self._header_trees` / `self._footer_trees`.
   - Store ZIP entry paths (`_header_paths` / `_footer_paths`) for re-serialization.

2. **Fragment indexing (`document.py`):**
   - `full_element_map()` returns a unified `dict[str, etree._Element]`:
     - Body: `"1"`, `"2"`, ...
     - Headers: `"header_1.1"`, `"header_1.2"`, ...
     - Footers: `"footer_1.1"`, `"footer_1.2"`, ...
   - Only `<w:p>` elements are included for headers/footers; tables are skipped.
   - `resolve_fragment_id()` returns `(element, tree_type)`.

3. **Extraction (`converter.py` + `server.py`):**
   - `full_to_fragments()` processes body + headers + footers.
   - Tagged format: `<f=header_1.1>Confidential</f=header_1.1>`.
   - JSON format: `{"type": "paragraph", "fragment_id": "header_1.3", "text": "..."}`.

4. **Change model (`models.py`):**
   - `ParagraphChange.fragment_id` changed to `str` with `coerce_numbers_to_str=True`.
   - Pydantic auto-coerces `5` → `"5"` for backward-compatible JSON/API usage.

5. **Applying changes (`redliner.py` + handlers):**
   - Uses `full_element_map()` for all paragraph changes.
   - `_sort_paragraph_changes()` uses document-order position index for stable sorting.
   - `handle_modify` / `handle_delete` / `handle_append_after` work unchanged (operate on `<w:p>`).
   - Comments in headers/footers work correctly (global `comments.xml`).

6. **Serialization (`document.py`):**
   - `to_bytes()` re-serializes modified header/footer trees back into the ZIP.

**Tests:**
- `test_header_footer.py`: Extraction, modify, delete, append, comment, serialization roundtrip.

---

### T2.2 Section-break handling **[DONE]**

**Decision:** Defensive fix in `handle_append_after`: if the next sibling of the reference paragraph is `<w:sectPr>`, temporarily remove it, perform all insertions, then restore the `sectPr` at the end of the inserted block. This ensures the new paragraph stays in the same section.

**Tests:**
- `test_section_breaks.py`: Append after last paragraph before section break, multi-section extraction.

---

### T2.3 Table robustness improvements **[DONE]**

**Decision:** Improve test coverage and error messages for simple tables.

**Implementation notes:**
- `is_simple_table()` accepts optional `table_id` parameter; error messages include `table {id}, ...` prefix.
- Empty cells extract as `""`.
- Wide tables (10 columns) are fully addressable.

**Tests:**
- `test_table_robustness.py`: Empty cell extraction/modify, wide table extraction/modify, improved error messages.

---

## Tier 3: Advanced Features

### T3.1 Merged-cell table support **[DONE]**

**Problem:** Tables with `gridSpan` (horizontal merge) or `vMerge` (vertical merge) are entirely skipped, even when only a single header cell spans columns.

**Decision:** Support merged cells incrementally. Horizontal merges (`gridSpan`) first; vertical merges (`vMerge`) second.

**Implementation notes:**
1. **Grid builder (`table_utils.py`):**
   - `build_table_grid()` walks each row's `<w:tc>` elements and tracks `gridSpan` and `vMerge` (`restart` / `continue`).
   - Returns a logical grid of `GridCell` objects with `span`, `vspan`, and `is_spanned_over` flags.
   - Nested tables are still rejected.

2. **Cell addressing (`table_utils.py`):**
   - `get_cell_element(table, row, col)` resolves logical `(row, col)` through the grid map to the correct physical `<w:tc>`.
   - Targeting a spanned-over cell raises a clear error directing the user to the starting cell.

3. **Extraction (`converter.py`):**
   - `_extract_table_info()` populates `CellInfo` with `span` and `vspan` (default 1, omitted when trivial).
   - Tagged format: `<cell=2.1.1 span="2">Header A</cell=2.1.1>` and `<cell=2.1.2 span="0"></cell=2.1.2>`.
   - JSON format includes `span` and `vspan` fields in each cell object.

4. **Redlining (`table_redliner.py`):**
   - Modifying a merged cell modifies the spanning `<w:tc>`; the text replacement applies to the whole cell.
   - Deleting a merged cell's content clears the spanning cell.
   - No support for *splitting* a merged cell via redlining (out of scope).

**Tests:**
- `test_table_extraction.py`: Merged-cell extraction with span markers in both tagged and JSON formats.
- `test_table_redliner.py`: Modify spanning cell, error on spanned-over cell.

---

### T3.2 Nested tables

**Problem:** Some legal documents use nested tables for complex layouts (e.g., a terms table inside a wider schedule table).

**Decision:** Not supported in the near term. The flat fragment ID architecture cannot represent hierarchy without significant refactoring.

**Future design sketch:**
- Hierarchical fragment IDs: `2.1.3` (table 2, row 1, col 3) for a cell containing a nested table; the nested table's cells would be `2.1.3.1.1.1`, etc.
- `body_to_fragments()` becomes recursive.
- `element_map` becomes a tree or a flat dict with string keys.
- Only pursue if real user demand exists.

---

### T3.3 Footnotes / endnotes

**Problem:** Legal scholarship and some transactional documents use footnotes for definitions, citations, or cross-references.

**Decision:** Out of scope until requested. The architecture is similar to headers/footers (separate XML parts, relationship-based), so the pattern from T2.1 would apply.

---

## Open Questions for Future Iteration

1. **Numbered list continuity.** If we delete a numbered paragraph, Word auto-renumbers remaining items. If we mark it as tracked deletion, does the numbering look odd in the redlined view? Should we test this explicitly?

2. **Bookmark preservation.** Legal documents use Word bookmarks for cross-references. Our paragraph rebuilds during `handle_modify` may drop or corrupt bookmark start/end markers. Should we audit bookmark handling?

3. **Revision marks (`w:moveFrom` / `w:moveTo`).** Our hard-reject currently targets `w:ins` and `w:del`. Should `w:moveFrom` / `w:moveTo` also trigger rejection?

4. **Track changes author attribution.** Currently all new tracked changes use the configured author (default "AI Review"). Should we support preserving original authors for *existing* changes (if we ever allow dirty documents)?

5. **Performance on very large documents.** The current pipeline loads the entire ZIP into memory and parses all XML trees. Is this acceptable for 500-page merger agreements? Should we benchmark?

6. **Auto-accept vs. auto-reject for dirty documents.** Instead of hard-rejecting, could we offer a mode that auto-accepts all existing tracked changes into the baseline before applying new ones? (Higher risk; probably not worth it.)

7. **Diffing headers/footers.** Should `diff_fragments` compare headers and footers between two documents, or only body content?

8. **Hyperlink relationship target types.** Should we support hyperlinks to bookmarks (`#bookmark`) or internal document locations, or only external URLs?

---

## How to Use This Roadmap

- Each tier is a **milestone**. Do not start Tier 2 until Tier 1 is complete and all tests pass.
- Each item has an identifier (e.g., **T1.2**). Reference these IDs in commit messages and PR descriptions.
- When an item is implemented, update this file: change the status to `[DONE]` and add a link to the relevant PR or commit.
- When a new design question arises, add it to **Open Questions** and schedule a structured Q&A round.

---

## Changelog

| Date | Author | Change |
|------|--------|--------|
| 2026-04-27 | AI Assistant | Initial roadmap drafted after codebase review and structured Q&A. |
| 2026-04-29 | AI Assistant | Tier 1 completed: tracked-change rejection, hyperlink support, append fixes, lossiness infra. |
| 2026-04-29 | AI Assistant | Tier 2 completed: header/footer extraction & redlining, section-break handling, table robustness. |
| 2026-04-30 | AI Assistant | Merged-cell table support (T3.1): `build_table_grid`, `span`/`vspan` markers in extraction, redlining with spanned-cell error handling. Formatting artifacts (`****`, `____`) fixed. `collapse_empty` mode implemented for cleaner LLM output. |

