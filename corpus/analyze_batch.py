"""Analyze tagged extraction quality for a batch of corpus documents.

Run via:
    uv run python corpus/analyze_batch.py <batch_num> <id1> <id2> ...

Writes feedback to corpus/feedback/batch_<batch_num>.json
"""

import json
import sys
from pathlib import Path
from typing import Any

from docx_mcp.converter import fragments_to_tagged_text_interleaved, full_to_fragments
from docx_mcp.document import DocxDocument

CORPUS_DIR = Path(__file__).parent
DOWNLOAD_DIR = CORPUS_DIR / "downloaded"
FEEDBACK_DIR = CORPUS_DIR / "feedback"
FEEDBACK_DIR.mkdir(exist_ok=True)


def analyze_document(doc_id: str) -> dict[str, Any]:
    """Extract and analyze one document."""
    path = DOWNLOAD_DIR / f"{doc_id}.docx"
    result: dict[str, Any] = {
        "id": doc_id,
        "path": str(path),
        "load_ok": False,
        "extract_ok": False,
        "errors": [],
        "feedback": {
            "overall_usefulness": None,  # high / medium / low
            "strengths": [],
            "weaknesses": [],
            "improvements": [],
        },
        "stats": {},
    }

    if not path.exists():
        result["errors"].append("File not found")
        return result

    try:
        doc = DocxDocument(path)
        result["load_ok"] = True
    except Exception as e:
        result["errors"].append(f"Load failed: {e}")
        return result

    try:
        frag_result = full_to_fragments(doc)
        tagged = fragments_to_tagged_text_interleaved(frag_result.items)
        result["extract_ok"] = True
    except Exception as e:
        result["errors"].append(f"Extract failed: {e}")
        return result

    # Stats
    body_paras = [
        item for item in frag_result.items
        if isinstance(item, tuple) and not item[0].startswith(("header_", "footer_"))
    ]
    header_paras = [
        item for item in frag_result.items
        if isinstance(item, tuple) and item[0].startswith("header_")
    ]
    footer_paras = [
        item for item in frag_result.items
        if isinstance(item, tuple) and item[0].startswith("footer_")
    ]
    tables: list[Any] = [
        item for item in frag_result.items
        if hasattr(item, "table_id") and not hasattr(item, "reason")
    ]
    skipped_tables: list[Any] = [
        item for item in frag_result.items
        if hasattr(item, "reason")
    ]

    body_text = "\n".join(text for _fid, text in body_paras)
    body_words = len(body_text.split())

    result["stats"] = {
        "body_paragraphs": len(body_paras),
        "header_paragraphs": len(header_paras),
        "footer_paragraphs": len(footer_paras),
        "tables": len(tables),
        "skipped_tables": len(skipped_tables),
        "skipped_elements": len(frag_result.skipped_elements),
        "body_words_approx": body_words,
        "tagged_lines": len(tagged.splitlines()),
    }

    # Heuristic usefulness assessment
    feedback: dict[str, Any] = result["feedback"]

    # Check for empty/whitespace-only body paragraphs
    empty_paras = sum(1 for _fid, text in body_paras if not text.strip())
    if empty_paras > len(body_paras) * 0.5 and len(body_paras) > 5:
        msg = (
            f"{empty_paras}/{len(body_paras)} body paragraphs are empty — "
            "document may be sparse or use excessive whitespace"
        )
        feedback["weaknesses"].append(msg)

    # Check table representation
    if tables:
        feedback["strengths"].append(
            f"{len(tables)} table(s) extracted with cell coordinates and content"
        )
        for tbl in tables:
            if tbl.rows > 10 or tbl.cols > 5:
                feedback["strengths"].append(
                    f"Large table {tbl.table_id} ({tbl.rows}x{tbl.cols}) "
                    "is fully addressable by cell ID"
                )
    if skipped_tables:
        reasons: set[str] = set(t.reason for t in skipped_tables)
        feedback["weaknesses"].append(
            f"{len(skipped_tables)} table(s) skipped: {', '.join(reasons)}"
        )

    # Check header/footer visibility
    if header_paras:
        feedback["strengths"].append(
            f"{len(header_paras)} header paragraph(s) visible with prefixed IDs"
        )
    if footer_paras:
        feedback["strengths"].append(
            f"{len(footer_paras)} footer paragraph(s) visible with prefixed IDs"
        )

    # Check for formatting artifacts
    if "****" in body_text or "**\n**" in body_text:
        feedback["weaknesses"].append(
            "Empty bold/italic formatting artifacts visible (e.g., '****') "
            "— noisy for LLM reading"
        )

    # Check for hyperlink clarity
    if "](" in body_text:
        feedback["strengths"].append(
            "Hyperlinks preserved as Markdown [text](url)"
        )

    # Check for tab/whitespace noise
    tabs = body_text.count("\t")
    if tabs > 10:
        feedback["weaknesses"].append(
            f"Many literal tabs ({tabs}) in text — may represent tables "
            "or alignment but look odd as plain text"
        )

    # Overall usefulness
    problem_indicators = len(feedback["weaknesses"])
    if problem_indicators == 0 and body_words > 100:
        feedback["overall_usefulness"] = "high"
    elif problem_indicators <= 1 and body_words > 50:
        feedback["overall_usefulness"] = "medium"
    else:
        feedback["overall_usefulness"] = "low"

    # Improvement suggestions
    if skipped_tables:
        feedback["improvements"].append(
            "Consider supporting merged-cell or nested tables to reduce skipped content"
        )
    if empty_paras > 5:
        feedback["improvements"].append(
            "Empty paragraphs could be optionally collapsed to reduce visual noise"
        )
    if "****" in body_text:
        feedback["improvements"].append(
            "Filter out empty formatting runs to reduce artifact noise"
        )

    # Store first 20 lines of tagged output for inspection
    lines = tagged.splitlines()
    result["tagged_preview"] = lines[:30]

    return result


def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: analyze_batch.py <batch_num> <id1> <id2> ...")
        sys.exit(1)

    batch_num = sys.argv[1]
    doc_ids = sys.argv[2:]

    print(f"Batch {batch_num}: analyzing {len(doc_ids)} documents...")
    results: list[dict[str, Any]] = []
    for doc_id in doc_ids:
        print(f"  {doc_id[:16]}...", end=" ", flush=True)
        r = analyze_document(doc_id)
        results.append(r)
        if r["errors"]:
            print(f"ERRORS: {r['errors']}")
        else:
            usefulness = r["feedback"]["overall_usefulness"]
            words = r["stats"]["body_words_approx"]
            print(f"usefulness={usefulness} words={words}")

    out_path = FEEDBACK_DIR / f"batch_{batch_num}.json"
    with open(out_path, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\nWrote feedback to {out_path}")


if __name__ == "__main__":
    main()
