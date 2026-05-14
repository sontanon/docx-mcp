#!/usr/bin/env python3
"""Analyze corpus quality and flag low-quality documents."""

import json
import re
from pathlib import Path

CORPUS_DIR = Path("/home/santiago/Code/docx-mcp/corpus")
METADATA_PATH = CORPUS_DIR / "metadata" / "all_downloaded.json"
EXTRACTED_DIR = CORPUS_DIR / "extracted_text"
OUTPUT_PATH = CORPUS_DIR / "quality_report.json"

# Template placeholder patterns
PLACEHOLDER_RE = re.compile(
    r"\[\s*(company name|date|your name|address|insert|place holder|placeholder)\s*\]"
    r"|<insert\s|lorem ipsum|type here|\[docjuris|\[enter",
    re.IGNORECASE,
)


def analyze_document(doc: dict) -> dict:
    """Analyze a single document's extracted text."""
    doc_id = doc["id"]
    text_path = EXTRACTED_DIR / f"{doc_id}.txt"
    text = text_path.read_text(encoding="utf-8")

    # Extract all body fragments <f=N>...</f=N> (plain integer N)
    body_fragments = re.findall(r"<f=(\d+)>(.*?)</f=\1>", text, re.DOTALL)
    body_texts = [content for _fid, content in body_fragments]
    body_fragment_count = len(body_texts)

    empty_body_fragments = sum(1 for t in body_texts if not t.strip())
    non_empty_texts = [t.strip() for t in body_texts if t.strip()]
    non_empty_count = len(non_empty_texts)
    body_words = sum(len(t.split()) for t in non_empty_texts)

    all_body_text = " ".join(non_empty_texts).lower()
    filename_lower = doc["filename"].lower()

    # Determine flag reason
    reason = None

    if body_fragment_count == 0:
        reason = "No body text: document consists only of headers/footers/tables"
    elif empty_body_fragments == body_fragment_count:
        reason = "Effectively empty: all body fragments are empty or whitespace-only"
    elif body_fragment_count <= 5 and body_words <= 100:
        reason = (
            f"Too small: only {body_fragment_count} fragments, "
            f"{body_words} words, no substantive content"
        )
    elif non_empty_count <= 3 and body_words <= 30:
        reason = (
            f"Too small: only {non_empty_count} non-empty fragments, "
            f"{body_words} words, no substantive content"
        )
    elif "template" in filename_lower and body_words < 100:
        reason = "Template shell: filename indicates template and content is minimal"
    elif PLACEHOLDER_RE.search(all_body_text) and body_words < 150:
        reason = "Template shell: mostly placeholder text with no substantive content"
    elif body_words < 20 and non_empty_count <= 2:
        reason = f"Too small: only {non_empty_count} non-empty fragments, {body_words} words"

    # Garbage content heuristic
    if reason is None:
        path_like = len(re.findall(r"\b[a-zA-Z]:\\|/\w+/\w+\.\w+\b", all_body_text))
        if path_like > 3 and body_words < 50:
            reason = "Garbage content: appears to be a file listing or path dump"
        elif body_words > 0:
            alpha_chars = sum(1 for c in all_body_text if c.isalpha())
            if alpha_chars > 0:
                alpha_ratio = alpha_chars / len(all_body_text)
                if alpha_ratio < 0.5 and body_words < 30:
                    reason = "Garbage content: high non-alphanumeric ratio, likely random data"

    return {
        "id": doc_id,
        "filename": doc["filename"],
        "type": doc["type"],
        "topic": doc["topic"],
        "word_count": body_words,
        "fragment_count": doc.get("fragment_count", body_fragment_count),
        "body_fragments": body_fragment_count,
        "empty_body_fragments": empty_body_fragments,
        "non_empty_body_fragments": non_empty_count,
        "reason": reason,
        "notes": None,
    }


def main() -> None:
    with open(METADATA_PATH, encoding="utf-8") as f:
        metadata = json.load(f)

    flagged = []
    keep = []

    for doc in metadata:
        result = analyze_document(doc)
        if result["reason"]:
            flagged.append(result)
        else:
            result["notes"] = "Good document with substantial text"
            keep.append(result)

    # Add more specific notes for kept documents
    for k in keep:
        if k["body_fragments"] < 10:
            k["notes"] = "Small but usable document"
        elif k["body_fragments"] > 100:
            k["notes"] = "Large document with substantial text"
        else:
            k["notes"] = "Good document with substantial text"

    report = {
        "flagged_for_removal": [
            {
                "id": f["id"],
                "filename": f["filename"],
                "type": f["type"],
                "topic": f["topic"],
                "word_count": f["word_count"],
                "fragment_count": f["fragment_count"],
                "reason": f["reason"],
            }
            for f in flagged
        ],
        "keep": [
            {
                "id": k["id"],
                "filename": k["filename"],
                "type": k["type"],
                "topic": k["topic"],
                "word_count": k["word_count"],
                "fragment_count": k["fragment_count"],
                "body_fragments": k["body_fragments"],
                "notes": k["notes"],
            }
            for k in keep
        ],
        "stats": {
            "total": len(metadata),
            "flagged": len(flagged),
            "kept": len(keep),
        },
    }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)

    # Print summary
    print(f"Total documents analyzed: {len(metadata)}")
    print(f"Flagged for removal: {len(flagged)}")
    print(f"Kept: {len(keep)}")
    print()

    if flagged:
        print("Flagged documents:")
        for f in flagged:
            print(f"  - {f['filename']} ({f['type']}/{f['topic']}): {f['reason']}")
        print()

    # Recommendations
    need_counts: dict[str, int] = {}
    for f in flagged:
        key = f"{f['type']} docs"
        need_counts[key] = need_counts.get(key, 0) + 1

    if need_counts:
        print("Recommendations for replacement:")
        for doc_type, count in sorted(need_counts.items()):
            print(f"  - Need {count} more {doc_type} to replace flagged ones")
    else:
        print("Recommendations: No replacements needed — corpus is fully usable.")


if __name__ == "__main__":
    main()
