"""Fetch a diverse corpus of real .docx documents and extract tagged text.

This script:
1. Loads docx-corpus metadata from HuggingFace
2. Selects diverse documents across types and topics
3. Downloads them
4. Extracts tagged (LLM-friendly) text representations
5. Saves everything with clear metadata
"""

import json
import zipfile
from pathlib import Path

import requests
from datasets import load_dataset  # type: ignore

from docx_mcp.audit import audit_document
from docx_mcp.converter import fragments_to_tagged_text_interleaved, full_to_fragments
from docx_mcp.document import DocxDocument

CORPUS_DIR = Path(__file__).parent
DOWNLOAD_DIR = CORPUS_DIR / "downloaded"
DOWNLOAD_DIR.mkdir(exist_ok=True)
TEXT_DIR = CORPUS_DIR / "extracted_text"
TEXT_DIR.mkdir(exist_ok=True)
META_DIR = CORPUS_DIR / "metadata"
META_DIR.mkdir(exist_ok=True)

# Target: at least 30 documents across these categories
CATEGORIES = [
    ("legal", "en", 0.8, 10),
    ("forms", "en", 0.8, 6),
    ("policies", "en", 0.8, 5),
    ("reports", "en", 0.8, 5),
    ("correspondence", "en", 0.8, 5),
]

# Quality floor: skip tiny documents that are likely shells or junk
MIN_WORD_COUNT = 100


def is_valid_docx(path: Path) -> bool:
    """Check if file is a valid ZIP with word/document.xml."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            return "word/document.xml" in zf.namelist()
    except zipfile.BadZipFile:
        return False


def download_file(url: str, path: Path) -> bool:
    """Download a file, return True on success."""
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        path.write_bytes(resp.content)
        return True
    except Exception as e:
        print(f"    Download failed: {e}")
        return False


def pick_diverse(items: list[dict], n: int) -> list[dict]:
    """Pick n items with topic diversity and word count spread."""
    sorted_items = sorted(items, key=lambda x: x["word_count"])
    total = len(sorted_items)
    if total == 0:
        return []

    selected = []
    topics_seen = set()

    # Always include smallest and largest for range
    if n >= 2 and total >= 2:
        selected.append(sorted_items[0])
        selected.append(sorted_items[-1])

    # Pick from quartiles
    for pct in [0.25, 0.5, 0.75]:
        idx = int(total * pct)
        candidate = sorted_items[idx]
        if candidate not in selected:
            selected.append(candidate)

    # Fill remaining with topic-diverse picks
    for item in sorted_items:
        if len(selected) >= n:
            break
        topic = item["topic"]
        if topic not in topics_seen and item not in selected:
            selected.append(item)
            topics_seen.add(topic)

    # If still not enough, just fill
    for item in sorted_items:
        if len(selected) >= n:
            break
        if item not in selected:
            selected.append(item)

    return selected[:n]


def load_existing_manifest(doc_type: str) -> list[dict]:
    """Load existing manifest if present."""
    path = META_DIR / f"{doc_type}_manifest.json"
    if path.exists():
        with open(path) as f:
            return json.load(f)
    return []


def generate_manifests() -> dict[str, list[dict]]:
    """Generate filtered manifests for each category, filling gaps."""
    print("Loading docx-corpus dataset...")
    ds = load_dataset("superdoc-dev/docx-corpus", split="train")
    print(f"Total rows: {len(ds):,}")

    manifests: dict[str, list[dict]] = {}
    for doc_type, lang, min_conf, target in CATEGORIES:
        existing = load_existing_manifest(doc_type)
        existing_ids = {item["id"] for item in existing}
        needed = target - len(existing)

        print(f"\nCategory: {doc_type} — have {len(existing)}, need {needed} more")
        if needed <= 0:
            print("  -> Already at target, skipping")
            manifests[doc_type] = existing
            continue

        print(
            f"  Filtering: type={doc_type}, lang={lang}, "
            f"conf>={min_conf}, words>={MIN_WORD_COUNT}..."
        )
        filtered = ds.filter(
            lambda x, dt=doc_type, ln=lang, mc=min_conf: (
                x["type"] == dt
                and x["language"] == ln
                and x["confidence"] >= mc
                and x["word_count"] >= MIN_WORD_COUNT
            ),
        )
        items = []
        for row in filtered:
            items.append(
                {
                    "id": row["id"],
                    "filename": row["filename"],
                    "type": row["type"],
                    "topic": row["topic"],
                    "language": row["language"],
                    "word_count": row["word_count"],
                    "confidence": row["confidence"],
                    "url": row["url"],
                }
            )
        print(f"  Found {len(items):,} matching rows")

        # Exclude already-downloaded IDs
        new_items = [item for item in items if item["id"] not in existing_ids]
        print(f"  After excluding existing: {len(new_items):,}")

        picked = pick_diverse(new_items, needed)
        print(f"  Selected {len(picked)} new documents")

        merged = existing + picked
        manifests[doc_type] = merged

        # Save merged manifest
        manifest_path = META_DIR / f"{doc_type}_manifest.json"
        with open(manifest_path, "w") as f:
            json.dump(merged, f, indent=2)
        print(f"  Saved manifest to {manifest_path}")

    return manifests


def download_manifests(manifests: dict[str, list[dict]]) -> list[dict]:
    """Download all documents from manifests."""
    all_items = []
    for _doc_type, items in manifests.items():
        all_items.extend(items)

    print(f"\n\n=== DOWNLOADING {len(all_items)} DOCUMENTS ===")
    downloaded = []

    for item in all_items:
        doc_id = item["id"]
        filename = f"{doc_id}.docx"
        path = DOWNLOAD_DIR / filename

        if path.exists():
            print(f"  {doc_id[:16]}... already exists")
            valid = is_valid_docx(path)
        else:
            print(
                f"  {doc_id[:16]}... downloading "
                f"({item['word_count']:,} words, {item['type']}, {item['topic']})"
            )
            success = download_file(item["url"], path)
            valid = success and is_valid_docx(path)

        if valid:
            size_kb = path.stat().st_size / 1024
            print(f"    -> Valid .docx ({size_kb:.1f} KB)")
            downloaded.append({"item": item, "path": path, "valid": True})
        else:
            print("    -> INVALID")
            if path.exists():
                path.unlink()
            downloaded.append({"item": item, "path": path, "valid": False})

    return downloaded


def extract_and_save(downloaded: list[dict]) -> list[dict]:
    """Extract tagged text and audit info for each downloaded document."""
    # Load existing metadata to merge
    meta_path = META_DIR / "all_downloaded.json"
    existing_results: list[dict] = []
    if meta_path.exists():
        with open(meta_path) as f:
            existing_results = json.load(f)
    existing_ids = {r["id"] for r in existing_results}

    valid_items = [d for d in downloaded if d["valid"]]
    new_items = [d for d in valid_items if d["item"]["id"] not in existing_ids]
    print(f"\n\n=== EXTRACTING TEXT FROM {len(new_items)} NEW DOCUMENTS ===")

    results = list(existing_results)
    for d in new_items:
        item = d["item"]
        path = d["path"]
        doc_id = item["id"]

        print(f"\n  {doc_id[:16]}... ({item['type']}: {item['topic']})")

        result = {
            "id": doc_id,
            "type": item["type"],
            "topic": item["topic"],
            "word_count": item["word_count"],
            "confidence": item["confidence"],
            "filename": item["filename"],
        }

        try:
            doc = DocxDocument(path)
        except Exception as e:
            result["error"] = f"Load failed: {e}"
            results.append(result)
            print(f"    ERROR loading: {e}")
            continue

        # Audit
        try:
            audit = audit_document(doc)
            result["audit"] = audit.to_dict()
            print(f"    Audit: {audit.header_parts}H/{audit.footer_parts}F "
                  f"{audit.simple_tables}T/{audit.skipped_tables}S")
        except Exception as e:
            result["audit_error"] = str(e)
            print(f"    Audit ERROR: {e}")

        # Extract tagged text
        try:
            frag_result = full_to_fragments(doc)
            tagged_text = fragments_to_tagged_text_interleaved(frag_result.items)

            text_path = TEXT_DIR / f"{doc_id}.txt"
            text_path.write_text(tagged_text, encoding="utf-8")

            result["fragment_count"] = len(frag_result.items)
            result["skipped_count"] = len(frag_result.skipped_elements)
            result["tagged_text_path"] = str(text_path.relative_to(CORPUS_DIR))
            print(
                f"    Extracted: {result['fragment_count']} fragments, "
                f"{result['skipped_count']} skipped"
            )
            print(f"    Saved tagged text to {text_path.name}")
        except Exception as e:
            result["extract_error"] = str(e)
            print(f"    Extract ERROR: {e}")

        results.append(result)

    # Save combined metadata
    with open(meta_path, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\nSaved combined metadata to {meta_path}")

    return results


def print_summary(results: list[dict]) -> None:
    """Print summary of corpus."""
    print("\n\n" + "=" * 60)
    print("CORPUS SUMMARY")
    print("=" * 60)

    total = len(results)
    successful = sum(1 for r in results if "error" not in r)
    by_type: dict[str, int] = {}
    by_topic: dict[str, int] = {}

    for r in results:
        by_type[r["type"]] = by_type.get(r["type"], 0) + 1
        by_topic[r["topic"]] = by_topic.get(r["topic"], 0) + 1

    print(f"\nTotal documents:     {total}")
    print(f"Successfully parsed: {successful}")
    print(f"Failed:              {total - successful}")

    print("\nBy type:")
    for t, c in sorted(by_type.items(), key=lambda x: -x[1]):
        print(f"  {t:20s}: {c}")

    print("\nBy topic:")
    for t, c in sorted(by_topic.items(), key=lambda x: -x[1]):
        print(f"  {t:20s}: {c}")

    print("\nOutput directories:")
    print(f"  Documents:  {DOWNLOAD_DIR}")
    print(f"  Tagged text:{TEXT_DIR}")
    print(f"  Metadata:   {META_DIR}")

    # Size stats
    total_words = sum(r.get("word_count", 0) for r in results)
    total_frags = sum(r.get("fragment_count", 0) for r in results if "fragment_count" in r)
    print(f"\nTotal words:      {total_words:,}")
    print(f"Total fragments:  {total_frags:,}")


if __name__ == "__main__":
    print("=" * 60)
    print("DOCX-MCP CORPUS FETCHER")
    print("=" * 60)

    manifests = generate_manifests()
    downloaded = download_manifests(manifests)
    results = extract_and_save(downloaded)
    print_summary(results)
