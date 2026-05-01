# Improvement Plan: Extraction Quality & Table Robustness

## Context

After analyzing 31 real documents from the docx-corpus, sub-agents identified that **71% of documents** (22/31) rate as "low usefulness" for LLM redlining due to extraction noise. The top issues are:

1. **Formatting artifacts** (`****`, `____`) — ~90% of documents
2. **Excessive empty paragraphs** — ~60% of documents
3. **Skipped tables** (merged cells, nested tables) — ~40% of documents

This plan addresses all three issues.

---

## Item 1: Fix Formatting Artifacts (`****`, `____`)

**Problem:** Empty bold/italic/underline runs produce artifact markers. A run with `<w:tab/>` or `<w:t> </w:t>` and bold formatting produces `**\t**` or `** **`. A sequence of empty runs produces `****`.

**Root cause:** `_format_run` checks `if not text:` which catches empty strings but **not** whitespace-only strings.

**Fix:**
```python
# In _format_run (converter.py, line 179)
text = _extract_run_text(run, include_del_text=include_del_text)
- if not text:
+ if not text or not text.strip():
      return ""
```

**Rationale:** A bold space or bold tab carries no semantic information for text editing. In legal documents, formatting applies to words, not whitespace.

**Risk:** Negligible. No legitimate use case for bold/italic whitespace in legal text.

**Test:** Verify that documents previously showing `****` no longer do. Add unit test with a paragraph containing bold empty run + bold tab run + normal text.

---

## Item 2: Optional Empty Paragraph Suppression

**Problem:** Word documents use empty `<w:p>` elements for vertical spacing. In some documents, 50%+ of paragraphs are empty, inflating fragment counts and adding noise.

**Decision:** Add `collapse_empty: bool = False` parameter to extraction functions. When True, empty paragraphs are suppressed from the output.

**ID Stability Problem:** If extraction suppresses empty paragraphs but redlining still uses original numbering, IDs get out of sync.

**Solution:**

### Design: `FragmentIdMapper`

A lightweight helper class that maps between "display IDs" (clean, collapsed) and "real IDs" (actual XML element indices).

```python
@dataclass
class FragmentIdMapper:
    """Maps between collapsed display IDs and real XML element indices."""

    # display_id (1-based, collapsed) -> real_element_index (0-based into element list)
    display_to_real: dict[int, int]
    # real_element_index -> display_id (or None if suppressed)
    real_to_display: dict[int, int | None]

    @classmethod
    def build(
        cls,
        elements: list[etree._Element],
        *,
        collapse_empty: bool = False,
        is_empty: Callable[[etree._Element], bool] | None = None,
    ) -> FragmentIdMapper:
        """Build mapper from an element list."""
```

### Usage in Extraction

```python
def full_to_fragments(doc, *, collapse_empty: bool = False, ...):
    elements = ...  # all body/header/footer elements
    mapper = FragmentIdMapper.build(elements, collapse_empty=collapse_empty)
    # Use mapper.display_to_real to look up actual elements
    # Output fragment IDs use display IDs
```

### Usage in Redlining

```python
def apply_redlines(source, changes, *, collapse_empty: bool = False, ...):
    doc = DocxDocument(source)
    # Build mapper with same collapse_empty flag
    mapper = FragmentIdMapper.build(all_elements, collapse_empty=collapse_empty)
    # Validate: check that all change fragment_ids exist in display IDs
    # Resolve: display_id -> real_element_index -> actual element
```

### Error Detection

If a user passes changes with IDs that don't match the `collapse_empty` setting:

```python
# In _validate_paragraph_changes:
for change in paragraph_changes:
    display_id = int(change.fragment_id)
    if display_id not in mapper.display_to_real:
        msg = (
            f"Change references fragment_id={change.fragment_id}, "
            f"but with collapse_empty={collapse_empty}, valid IDs are "
            f"{min(mapper.display_to_real)}..{max(mapper.display_to_real)}. "
            f"Hint: extraction and redlining must use the same collapse_empty setting."
        )
        raise ValueError(msg)
```

### API Changes

- `extract_fragments` tool: add `collapse_empty: bool = False`
- `apply_redlines` function: add `collapse_empty: bool = False`
- Server tool params: add `collapse_empty` to both tools

**Default:** `False` (backward compatible). When `True`, output is cleaner but requires consistent usage.

**Test:**
- Document with 5 paragraphs where #2 and #4 are empty
- `collapse_empty=False` → IDs 1,2,3,4,5
- `collapse_empty=True` → IDs 1,2,3 (display), mapper knows 1→0, 2→2, 3→4
- Verify redlining with mismatched `collapse_empty` raises clear error

---

## Item 3: Merged-Cell Table Support

**Problem:** Tables with `gridSpan` (horizontal merge) or `vMerge` (vertical merge) are entirely skipped, even when only a single header cell spans columns.

**Decision:** Support merged cells incrementally. Horizontal merges (`gridSpan`) first; vertical merges (`vMerge`) second.

### Grid Builder

Build a logical grid that accounts for spans:

```python
@dataclass
class GridCell:
    """A cell in the logical table grid."""

    cell_id: str          # "table.row.col" for the starting cell
    text: str
    span: int = 1         # horizontal span (gridSpan)
    vspan: int = 1        # vertical span (vMerge)
    is_spanned_over: bool = False  # True for cells covered by another cell's span
```

### Grid Building Algorithm

```python
def build_table_grid(tbl: etree._Element, table_id: int) -> list[list[GridCell]]:
    """Build a logical grid from a table with merged cells.

    For each row:
    1. Iterate <w:tc> cells in row order
    2. For each cell, check w:tcPr/w:gridSpan@val for horizontal span
    3. Check w:tcPr/w:vMerge for vertical span ("restart" = start, "continue" = covered)
    4. Place cell in grid, marking spanned-over positions with is_spanned_over=True
    5. If a cell would overlap an already-placed cell, that's a complex merge — skip
    """
```

### Output Format (Option B)

For a table with a header spanning 3 columns:

```xml
<table=1 rows=2 cols=4>
<cell=1.1.1 span="3">Header spanning three columns</cell=1.1.1>
<cell=1.1.2 span="0"></cell=1.1.2>   <!-- spanned over -->
<cell=1.1.3 span="0"></cell=1.1.3>   <!-- spanned over -->
<cell=1.1.4>Normal cell</cell=1.1.4>
<cell=1.2.1>C</cell=1.2.1>
<cell=1.2.2>D</cell=1.2.2>
<cell=1.2.3>E</cell=1.2.3>
<cell=1.2.4>F</cell=1.2.4>
</table=1>
```

For vertical merge:
```xml
<cell=1.1.1 vspan="2">Row 1-2 merged</cell=1.1.1>
<cell=1.2.1 vspan="0"></cell=1.2.1>   <!-- spanned over vertically -->
```

### Redlining Merged Cells

For `modify_cell` on a merged cell:
- If targeting the **starting cell** (span > 0, vspan > 0): modify normally
- If targeting a **spanned-over cell** (span = 0 or vspan = 0): raise error — "This cell is part of a merge. Target the starting cell (X.Y.Z) instead."

### Incremental Approach

**Phase 1:** Support horizontal merges only (`gridSpan`)
- Most common case (header rows, signatory tables)
- Simpler grid building

**Phase 2:** Support vertical merges (`vMerge`)
- Less common but important for some forms
- Requires tracking spans across rows

**Phase 3:** Support nested tables
- Most complex; cells contain entire sub-tables
- May require recursive cell ID scheme: `1.1.1.1.1` for nested

### `is_simple_table` Update

Current: rejects any table with gridSpan or vMerge.

New: accept tables with gridSpan/vMerge if the merge is "clean" (no overlapping spans, no nested tables). Only reject genuinely malformed tables.

```python
def is_simple_table(tbl: etree._Element, table_id: int | None = None) -> tuple[bool, str | None]:
    """Check if table is editable.

    Returns:
        (True, None) if table is fully editable (no merges, no nesting).
        (True, "has_merges") if table has clean merges (editable with span markers).
        (False, reason) if table is too complex to edit.
    """
```

**Test:**
- Fixture with 2-column table, header row spans both columns
- Fixture with 3-column table, first column spans 2 rows
- Verify extraction produces span markers
- Verify modify_cell on starting cell works
- Verify modify_cell on spanned-over cell raises clear error

---

## Item 4: JSON Output Format Update

If we add span markers, the JSON `extract_fragments` output needs a new field:

```json
{
  "type": "table",
  "table_id": 1,
  "rows": 2,
  "cols": 4,
  "cells": [
    [
      {"cell_id": "1.1.1", "row": 1, "col": 1, "text": "Header", "span": 3},
      {"cell_id": "1.1.2", "row": 1, "col": 2, "text": "", "span": 0},
      {"cell_id": "1.1.3", "row": 1, "col": 3, "text": "", "span": 0},
      {"cell_id": "1.1.4", "row": 1, "col": 4, "text": "Normal"}
    ]
  ]
}
```

The `span` field is optional (defaults to 1, omitted for brevity). `span: 0` means "spanned over."

---

## Item 5: Tabs

**Decision:** Keep tabs as-is. Document as a known limitation.

Rationale: Tabs may carry structural intent (alignment, indentation). Normalizing them risks losing information. The noise is acceptable compared to the risk of incorrect interpretation.

---

## Execution Order

1. **Item 1** (formatting artifacts) — one-line fix + test
2. **Item 2** (empty paragraph suppression) — new `FragmentIdMapper` + API changes + tests
3. **Item 3** (merged tables) — grid builder + span markers + tests
4. **Update ROADMAP** — mark completed items

---

## Files to Modify

| File | Changes |
|------|---------|
| `src/docx_mcp/converter.py` | Fix `_format_run` whitespace check; add `collapse_empty` to `full_to_fragments`; update table extraction for spans |
| `src/docx_mcp/models.py` | Add `collapse_empty` to `ParagraphChange` or tool params |
| `src/docx_mcp/server.py` | Add `collapse_empty` to tool params |
| `src/docx_mcp/redliner.py` | Add `collapse_empty` param; integrate `FragmentIdMapper` |
| `src/docx_mcp/table_utils.py` | Update `is_simple_table` for merge detection |
| `src/docx_mcp/table_redliner.py` | Handle span markers in `modify_cell` |
| `tests/test_converter.py` | Tests for whitespace fix |
| `tests/test_redliner.py` | Tests for `collapse_empty` mode |
| `tests/test_tables.py` | Tests for merged-cell tables |
| `tests/generators/generate_fixtures.py` | Fixtures for merged-cell tables |
| `ROADMAP.md` | Update status |
