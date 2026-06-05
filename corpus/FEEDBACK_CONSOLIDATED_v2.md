# Corpus Extraction Quality Feedback — Consolidated Report v2

Generated: 2026-04-30
Method: 7 parallel sub-agents analyzed 31 real documents using new extraction (post-improvements)
Extractor version: `feat/tier2-expand-coverage` with formatting-artifact fix, `collapse_empty` mode, and merged-cell table support.

---

## Executive Summary

| Rating | Count | % | v1 Baseline |
|---|---|---|---|
| **High** | 7 | 23% | 10% |
| **Medium** | 12 | 39% | 19% |
| **Low** | 12 | 39% | 71% |

**Key improvement:** Low-quality documents dropped from **71% to 39%**. High-quality documents more than doubled (10% → 23%).

**What got better:**
- ✅ **Skipped tables:** From ~40% of documents (12/31) to **1 skipped table across all 31 docs**. Merged-cell tables are now fully extracted.
- ✅ **Formatting artifacts:** From ~90% of documents affected to ~55% (17/31). Total artifact count dropped from thousands to **281**.
- ✅ **Empty paragraph suppression:** `collapse_empty=True` is available and works correctly (verified: removes exactly the expected count, e.g., 8,607 from `5c0005`).

**What got worse / New issues:**
- ⚠️ ~~**Span marker noise:**~~ **FIXED** — Spanned-over cells are now omitted from output. `5721ca` went from 18,637 span markers to **2,044** (89% reduction). `4edf7d` went from 6,426 to **411** (94% reduction).
- ⚠️ **Residual bold artifacts:** 281 remaining `****` occurrences across 17 documents. These are split-bold-run patterns (e.g., `**S****tate**`) that the whitespace-only run filter does not catch.

---

## Global Metrics

| Metric | Total | Per-doc median | Worst offender |
|---|---|---|---|
| Lines | — | 97 | `5c0005` (26,064) |
| Characters | — | 5,430 | `5c0005` (2.3M) |
| Bold artifacts (`****`) | 281 | 2 | `2d9608` (117) |
| Underline artifacts (`____`) | 15 | 0 | `a586b1` (1) |
| Literal tabs (`\t`) | 8,058 | 0 | `5c0005` (6,269) |
| Empty paragraphs | 2,420 | 0 | `bdd19f` (2,285) |
| Tables extracted | 282 | 0 | `5c0005` (151) |
| Tables skipped | **1** | 0 | `2b838bb` (1 nested table) |
| Span markers (`span=`) | ~3,000 | 0 | `5721ca` (2,044 post-fix) |

---

## Document-by-Document Analysis

| # | Filename | Lines | Chars | Artifacts | Tabs | Empty | Tables | Skipped | Spans | Rating | Key Issues |
|---|----------|-------|-------|-----------|------|-------|--------|---------|-------|--------|------------|
| 1 | `05fbf2a0` | 21 | 1,304 | 1 | 7 | 0 | 0 | 0 | 0 | low | very_short, tabs, formatting_artifacts |
| 2 | `122bbf79` | 12 | 329 | 0 | 0 | 0 | 0 | 0 | 0 | low | very_short |
| 3 | `17f2564d` | 0 | 11 | 0 | 0 | 0 | 0 | 0 | 0 | low | empty_document, very_short |
| 4 | `29a9a0f8` | 22 | 399 | 0 | 0 | 0 | 0 | 0 | 0 | low | mostly_empty, very_short |
| 5 | `2b838bb7` | 54 | 5,059 | 7 | 0 | 0 | 3 | 1 | 0 | medium | formatting_artifacts, skipped_tables |
| 6 | `2d5bf513` | 48 | 5,731 | 0 | 0 | 0 | 0 | 0 | 0 | high | (clean) |
| 7 | `2d96086f` | 11,210 | 1,661,225 | 122 | 240 | 0 | 89 | 0 | 264 | medium | formatting_artifacts, tabs, span_markers |
| 8 | `32a4ce11` | 0 | 11 | 0 | 0 | 0 | 0 | 0 | 0 | low | empty_document, very_short |
| 9 | `38868c65` | 467 | 50,227 | 9 | 0 | 0 | 5 | 0 | 91 | high | formatting_artifacts, span_markers |
| 10 | `40ff3b00` | 18 | 2,805 | 4 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, very_short |
| 11 | `471d4e32` | 1,491 | 50,090 | 1 | 0 | 0 | 1 | 0 | 0 | high | minimal_artifacts |
| 12 | `4edf7d22` | 803 | 341,738 | 17 | 64 | 0 | 1 | 0 | 411 | medium | tabs, formatting_artifacts |
| 13 | `53819804` | 26 | 5,385 | 2 | 0 | 0 | 0 | 0 | 0 | medium | very_short, formatting_artifacts |
| 14 | `545e6d0f` | 29 | 3,491 | 0 | 0 | 0 | 0 | 0 | 0 | high | (clean) |
| 15 | `5721cae2` | 4,526 | 1,161,232 | 107 | 4 | 0 | 28 | 0 | 2,044 | medium | formatting_artifacts |
| 16 | `5c000517` | 26,064 | 2,326,124 | 0 | 6,269 | 0 | 151 | 0 | 953 | high | tabs |
| 17 | `75e9cfde` | 57 | 4,733 | 4 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, very_short |
| 18 | `83beef8e` | 36 | 3,288 | 3 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, very_short |
| 19 | `84553adf` | 88 | 13,710 | 3 | 11 | 0 | 0 | 0 | 0 | high | formatting_artifacts, tabs |
| 20 | `86fafe35` | 42 | 4,639 | 1 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, very_short |
| 21 | `8d7c00f3` | 101 | 4,307 | 3 | 0 | 0 | 1 | 0 | 55 | medium | formatting_artifacts, table_span_artifacts |
| 22 | `91abe386` | 33 | 1,986 | 8 | 0 | 0 | 0 | 0 | 0 | low | formatting_artifacts, very_short |
| 23 | `a509c7a0` | 5 | 246 | 0 | 0 | 0 | 0 | 0 | 0 | low | very_short |
| 24 | `a586b1c9` | 101 | 22,845 | 11 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, missing_spaces |
| 25 | `ab4f6c29` | 129 | 8,625 | 4 | 0 | 0 | 1 | 0 | 4 | medium | formatting_artifacts, table_span_artifacts |
| 26 | `ae7c9554` | 97 | 5,430 | 2 | 0 | 42 | 0 | 0 | 0 | low | formatting_artifacts, empty_paragraphs |
| 27 | `bdd19f65` | 11,299 | 1,844,532 | 28 | 1,456 | 2,285 | 1 | 0 | 0 | medium | tabs, empty_paragraphs, formatting_artifacts |
| 28 | `d7468207` | 28 | 4,628 | 0 | 0 | 3 | 0 | 0 | 0 | high | (clean) |
| 29 | `ec94ef4d` | 165 | 12,997 | 15 | 4 | 90 | 0 | 0 | 0 | low | empty_paragraphs, formatting_artifacts |
| 30 | `ee619d01` | 6,317 | 902,211 | 4 | 0 | 0 | 0 | 0 | 0 | medium | formatting_artifacts, very_long |
| 31 | `f533eb69` | 40 | 1,604 | 1 | 3 | 0 | 1 | 0 | 0 | low | formatting_artifacts, tabs, very_short |

---

## Issue Deep-Dives

### Issue 1: Span Marker Noise (NEW)

**Frequency:** 5 documents affected; 2 are severe (`4edf7d`, `5721ca`)

**What it looks like:**
```xml
<cell=58.1.1 span="2">**Self-Direction ****Checklist**</cell=58.1.1>
<cell=58.1.2 span="0" vspan="0"></cell=58.1.2>
```

In `5721ca`, 28 tables generate **18,637 span markers** — roughly one span marker per 62 characters of output. The document becomes unreadable.

**Root cause:** We chose Option B in the Improvement Plan (emit spanned-over cells as empty with `span="0"`). For documents with many merged cells, this creates exponential noise.

**Recommended fix:** Switch to Option A — **omit spanned-over cells entirely**. Only the starting cell gets a `span="N"` attribute. The grid dimensions (`rows=`, `cols=`) already communicate the logical layout. This would reduce span markers by ~80-90%.

**Risk:** An LLM editing a table might try to address a cell that no longer exists in the output. Error handling in `get_cell_element()` already covers this (it raises an error for spanned-over positions).

---

### Issue 2: Residual Bold Artifacts

**Frequency:** 17/31 documents (55%)

**What it looks like:**
```
**S****tate** Register Volume 25
**P****arish Pump**
**trade****description**
```

**Root cause:** These are NOT whitespace-only runs. They are bold runs that contain a single letter or fragment of a word, adjacent to another bold run. Example:
- Run 1: `<w:b/><w:t>S</w:t>` → `**S**`
- Run 2: `<w:b/><w:t>tate</w:t>` → `**tate**`

Concatenated: `**S****tate**`

This happens when Word splits a single word into multiple runs (e.g., due to spell-check, revision tracking, or copy-paste). Our `paragraph_to_pseudo_markdown` concatenates runs naively.

**Fix complexity:** Medium. We would need to merge adjacent runs with identical formatting before applying pseudo-Markdown markers. This requires a pre-processing pass in `paragraph_to_pseudo_markdown`.

**Decision needed:** Is 281 artifacts across 31 documents worth a medium-complexity fix? The artifacts are visually annoying but do not prevent redlining.

---

### Issue 3: Literal Tabs

**Frequency:** 8/31 documents (26%)

**What it looks like:**
```
	1. Licensed Adjusters or motor vehicle physical damage appraisers...
```

**Decision:** Keep as-is (per Item 5 of Improvement Plan). Tabs may carry structural intent (indentation, alignment). Normalizing them risks losing information.

---

### Issue 4: Empty Paragraphs

**Frequency:** 6/31 documents with >10 empty fragments; 2 documents severely affected (`bdd19f`: 2,285 empty, `ec94ef`: 90 empty)

**Mitigation:** `collapse_empty=True` mode exists and is opt-in. When used, these empty paragraphs are suppressed.

**Decision:** No code change needed. Document the `collapse_empty` option more prominently.

---

## Files

- Raw batch feedback: `corpus/feedback_v2/batch_1.json` through `batch_7.json`
- Consolidated JSON: `corpus/feedback_v2/consolidated.json`
- Previous (v1) report: `corpus/FEEDBACK_CONSOLIDATED.md`

---

## Recommendations

1. ~~**Fix span marker noise (HIGH PRIORITY):**~~ **DONE** — Spanned-over cells omitted from tagged/JSON output. Noise reduced by ~90%.
2. **Decide on residual bold artifacts:** Either accept the 281 occurrences as acceptable noise, or implement adjacent-run merging in `paragraph_to_pseudo_markdown`.
3. **Document `collapse_empty`:** Add a note in server responses or documentation that `collapse_empty=True` is recommended for documents with excessive empty paragraphs.
4. **No action needed for tabs:** Keep literal tabs; document as known limitation.

