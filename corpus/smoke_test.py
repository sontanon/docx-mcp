"""Smoke-test our pipeline against downloaded real documents."""

from __future__ import annotations

import json
from pathlib import Path

from docx_mcp.audit import audit_document
from docx_mcp.converter import full_to_fragments
from docx_mcp.document import DocxDocument
from docx_mcp.redliner import apply_redlines

CORPUS_DIR = Path(__file__).parent
DOWNLOAD_DIR = CORPUS_DIR / "downloaded"


def smoke_test_document(path: Path) -> dict:
    """Run extraction and a no-op redline on a document."""
    result = {
        "path": str(path.name),
        "load_ok": False,
        "extract_ok": False,
        "extract_error": None,
        "fragment_count": 0,
        "table_count": 0,
        "audit": {},
        "redline_ok": False,
        "redline_error": None,
    }

    try:
        doc = DocxDocument(path)
        result["load_ok"] = True
    except Exception as e:
        result["load_error"] = str(e)
        return result

    # Audit
    try:
        audit = audit_document(doc)
        result["audit"] = audit.to_dict()
    except Exception as e:
        result["audit_error"] = str(e)

    # Extract fragments
    try:
        frag_result = full_to_fragments(doc)
        result["extract_ok"] = True
        result["fragment_count"] = len(frag_result.items)
        result["table_count"] = len(frag_result.skipped_elements)
    except Exception as e:
        result["extract_error"] = str(e)
        return result

    # Try a no-op redline (empty changes list)
    try:
        apply_redlines(path, [])
        result["redline_ok"] = True
    except Exception as e:
        result["redline_error"] = str(e)

    return result


def main() -> None:
    """Run smoke tests on all downloaded documents."""
    paths = sorted(DOWNLOAD_DIR.glob("*.docx"))
    print(f"Smoke-testing {len(paths)} documents...\n")

    results = []
    for path in paths:
        print(f"Testing: {path.name[:40]:40s} ... ", end="", flush=True)
        r = smoke_test_document(path)
        results.append(r)

        status = []
        if r["load_ok"]:
            status.append("load")
        if r["extract_ok"]:
            status.append(f"extract({r['fragment_count']})")
        if r["redline_ok"]:
            status.append("redline")
        if r["extract_error"]:
            status.append(f"EXTRACT_ERR: {r['extract_error'][:60]}")
        if r.get("redline_error"):
            status.append(f"REDLINE_ERR: {r['redline_error'][:60]}")

        print(" | ".join(status))

    # Summary
    total = len(results)
    load_ok = sum(1 for r in results if r["load_ok"])
    extract_ok = sum(1 for r in results if r["extract_ok"])
    redline_ok = sum(1 for r in results if r["redline_ok"])
    extract_errs = [r for r in results if r["extract_error"]]
    redline_errs = [r for r in results if r.get("redline_error")]

    print(f"\n{'='*50}")
    print(f"Total:      {total}")
    print(f"Load OK:    {load_ok}/{total}")
    print(f"Extract OK: {extract_ok}/{total}")
    print(f"Redline OK: {redline_ok}/{total}")

    if extract_errs:
        print(f"\nExtract errors ({len(extract_errs)}):")
        for r in extract_errs:
            print(f"  {r['path']}: {r['extract_error']}")

    if redline_errs:
        print(f"\nRedline errors ({len(redline_errs)}):")
        for r in redline_errs:
            print(f"  {r['path']}: {r['redline_error']}")

    # Save report
    report_path = CORPUS_DIR / "smoke_test_report.json"
    with open(report_path, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\nReport saved to {report_path}")


if __name__ == "__main__":
    main()
