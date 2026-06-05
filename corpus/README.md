# docx-mcp Real Document Corpus

This directory contains scripts and metadata for fetching a diverse corpus of real `.docx` documents from the public web, plus their extracted tagged-text representations.

## Directory Layout

```
corpus/
  explore.py              # Explore docx-corpus dataset statistics
  download.py             # Download a subset of documents (legacy)
  smoke_test.py           # Smoke-test downloaded docs against our pipeline
  fetch_corpus.py         # Master script: fetch 30+ diverse docs + extract text
  metadata/               # JSON manifests and combined metadata
    legal_manifest.json
    forms_manifest.json
    policies_manifest.json
    reports_manifest.json
    correspondence_manifest.json
    all_downloaded.json
  downloaded/             # Actual .docx files (gitignored, regenerable)
  extracted_text/         # Tagged text output (gitignored, regenerable)
```

## Quick Start

Fetch the full corpus (~30 documents, ~13 MB):

```bash
uv run python corpus/fetch_corpus.py
```

This will:
1. Load metadata from the `superdoc-dev/docx-corpus` HuggingFace dataset
2. Select diverse documents across 5 types (`legal`, `forms`, `policies`, `reports`, `correspondence`)
3. Download them to `corpus/downloaded/`
4. Extract tagged (LLM-friendly) text to `corpus/extracted_text/`
5. Save metadata to `corpus/metadata/all_downloaded.json`

## Tagged Text Format

Extracted files use our standard tagged format:

```
<f=1>First body paragraph text.</f=1>
<f=header_1.1>Header text</f=header_1.1>
<f=footer_1.1>Page 1</f=footer_1.1>
<table=2 rows=3 cols=4>
<cell=2.1.1>Header A</cell=2.1.1>
<cell=2.1.2>Cell content</cell=2.1.2>
...
</table=2>
<table=5 skipped reason="contains merged cells"/>
```

This is the same format returned by the MCP `extract_fragments` tool.

## Corpus Statistics (Final)

- **31 documents** across 5 types and 8 topics
- **~1.3M words** total
- **~43K fragments** total
- Size: ~13 MB .docx files + ~7 MB extracted text

| Type | Count | Avg Words | Notes |
|------|-------|-----------|-------|
| legal | 10 | 30,190 | 2 massive judicial docs (~280K words each) |
| forms | 6 | 38,536 | Includes table-heavy environmental permit form |
| policies | 5 | 23,870 | Mix of education and healthcare policies |
| reports | 5 | 45,757 | Large government reports with tables |
| correspondence | 5 | 25,921 | Government letters and announcements |

| Topic | Count |
|-------|-------|
| government | 6 |
| healthcare | 5 |
| education | 5 |
| legal_judicial | 5 |
| general | 4 |
| finance | 3 |
| nonprofit | 2 |
| environment | 1 |

### Known Edge Cases in Corpus

- **VML text boxes**: 2 documents use shape-based text that our parser does not extract (SWOT template, registration sign). These are kept as known edge cases.
- **Pre-existing tracked changes**: 6 documents contain tracked changes and are correctly rejected by `apply_redlines()`.
- **Table-heavy forms**: 2 large government forms store all content in tables. Our parser extracts them as `<table=N>` fragments, not body paragraphs.
- **Images in headers**: Several documents have images in header/footer areas.

## Smoke Testing

Run the pipeline against all downloaded documents:

```bash
uv run python corpus/smoke_test.py
```

This validates that every document can be:
- Loaded by `DocxDocument`
- Extracted by `full_to_fragments()`
- Processed by `apply_redlines()` (no-op test)

## Data Source

All documents come from [docx-corpus](https://huggingface.co/datasets/superdoc-dev/docx-corpus) — an open dataset of 736K+ real `.docx` files scraped from the public web, classified by type and topic.

License: ODC-By (Open Data Commons Attribution)
