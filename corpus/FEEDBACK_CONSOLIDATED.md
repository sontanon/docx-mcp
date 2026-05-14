# Corpus Extraction Quality Feedback — Consolidated Report

Generated: 2026-04-29
Method: 7 parallel sub-agents analyzed 31 real documents using `extract_fragments` (tagged format)

---

## Executive Summary

| Rating | Count | % |
|---|---|---|
| **High** | 3 | 10% |
| **Medium** | 6 | 19% |
| **Low** | 22 | 71% |

Only 3 of 31 documents (10%) produced clean, noise-free tagged text suitable for LLM redlining without issues. The vast majority suffer from one or more problems that degrade readability and editing fidelity.

---

## Issue 1: Empty Formatting Artifacts (`****`)

**Frequency:** ~90% of documents (28/31)

**What it looks like:**
```
<f=1>**\t****\t****\t****Academic Action Plan**</f=1>
<f=18>__R____ECOVERY POINTS:__ Number of credits...</f=18>
```

**Root cause:** Word documents often contain empty bold/italic/underline runs (e.g., `<w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r>`). Our `paragraph_to_pseudo_markdown` concatenates all runs and outputs formatting markers even when the text between them is empty.

**Impact:**
- Clutters output for LLM consumption
- Makes word-level diff noisy (artifacts appear as "changes" when they're not)
- Reduces signal-to-noise ratio

**Affected documents:** Nearly all. Worst in:
- `bdd19f...` (civil code): 1,470 tabs + artifacts across 11K paragraphs
- `ee619d...` (gov correspondence): 222 tabs + artifacts across 6K paragraphs
- `5c0005...` (judicial rules): artifacts across 16K paragraphs

---

## Issue 2: Excessive Empty Paragraphs

**Frequency:** ~60% of documents (19/31)

**What it looks like:**
```
<f=7></f=7>
<f=14></f=14>
<f=15></f=15>
```

**Root cause:** Word documents use empty paragraphs for vertical spacing. Our extractor faithfully outputs every `<w:p>`, including empty ones. In some documents (forms, templates), over 50% of paragraphs are empty.

**Impact:**
- Visual noise for LLMs
- Inflated fragment counts (e.g., 16,430 body paragraphs where 8,607 are empty)
- Word-level diff on empty paragraphs produces meaningless tracked changes

**Worst offenders:**
- `5c0005...`: 8,607/16,430 empty (52%)
- `ec94ef...`: 91/166 empty (55%)
- `91abe3...`: 19/34 empty (56%)
- `5721ca...`: high empty count in 968-fragment document

---

## Issue 3: Literal Tab Characters

**Frequency:** ~50% of documents (15/31)

**What it looks like:**
```
<f=1>**\t****\t****\t****Academic Action Plan**</f=1>
```

**Root cause:** Word uses tabs for alignment (indentation, table-like layouts). Our extractor preserves `\t` literally.

**Impact:**
- Tabs are invisible in many rendering contexts but consume tokens
- Multiple consecutive tabs look like noise
- May represent structural intent (columns, alignment) that's lost without context

**Worst offenders:**
- `bdd19f...`: 1,470 tabs
- `5c0005...`: hundreds of tabs
- `ee619d...`: 222 tabs

---

## Issue 4: Skipped Tables

**Frequency:** ~40% of documents (12/31)

**Root cause:** Tables with merged cells (gridSpan, vMerge) or nested tables are skipped entirely.

**Impact:**
- Lost structured content (forms, data tables, schedules)
- Documents where ALL content is in tables become effectively empty

**Worst offenders:**
- `4edf7d...` (MA DEP permit): 800 paragraphs in tables, all skipped — document reads as empty
- `5721ca...` (nonprofit manual): 18/28 tables skipped
- `5c0005...` (judicial rules): 52/151 tables skipped
- `2d9608...` (gov report): 25/89 tables skipped

---

## Clean Documents (Gold Standard)

These 3 documents produced noise-free output:

1. **`2d5bf513...`** (Fort Drum press release, 828 words)
   - 49 body paragraphs, 0 empty artifacts, 0 tabs, 0 skipped tables
   - Clean prose with headers/footers

2. **`d7468207...`** (Texas Health and Safety Code, 680 words)
   - 35 body paragraphs, 0 empty artifacts, 0 tabs, 0 skipped tables
   - Clean statute text

3. **`545e6d0f...`** (Finance report, 457 words)
   - 26 body paragraphs, 0 empty artifacts, 0 tabs, 0 skipped tables
   - Clean prose with footer

---

## Other Observations

### VML Text Boxes
- 2 documents (`32a4ce...` SWOT form, `17f256...` tableft) store all text in VML shapes
- Our parser extracts 0 body words from these
- **Decision:** Document as unsupported; not a priority for legal document pipeline

### Pre-existing Tracked Changes
- 6 documents in corpus contain tracked changes
- `apply_redlines()` correctly rejects them
- No action needed — this is correct behavior

### Hyperlinks
- Preserved well as Markdown `[text](url)`
- No issues reported

### Headers/Footers
- Successfully extracted with prefixed IDs
- No issues reported

---

## Improvement Areas (for Planning)

1. **Collapse empty formatting runs** — Don't output `****` when no text exists between markers
2. **Optional empty paragraph suppression** — Configurable mode to skip `<w:p>` with no text
3. **Tab normalization** — Replace tabs with spaces or collapse consecutive tabs
4. **Merged-cell table support** — Handle gridSpan/vMerge so tables aren't skipped
5. **Nested table support** — Handle tables inside table cells
6. **Better empty paragraph detection** — Distinguish "intentional spacing" from "noise"
7. **Content controls / SDT** — Some forms use structured document tags for fillable fields

---

## Files

- Raw batch feedback: `corpus/feedback/batch_1.json` through `batch_7.json`
- Forensic deep-dive: `corpus/forensic_report.json`
- Quality report: `corpus/quality_report.json`
- Smoke test results: `corpus/smoke_test_report.json`
