"""Download and validate a diverse subset of real .docx documents."""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from urllib.request import urlopen

import requests

from docx_mcp.audit import audit_document
from docx_mcp.document import DocxDocument

CORPUS_DIR = Path(__file__).parent
DOWNLOAD_DIR = CORPUS_DIR / "downloaded"
DOWNLOAD_DIR.mkdir(exist_ok=True)


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
        print(f"  Download failed: {e}")
        return False


def pick_diverse_subset(items: list[dict], n: int) -> list[dict]:
    """Pick n items distributed across word count ranges and topics."""
    # Sort by word count
    sorted_items = sorted(items, key=lambda x: x["word_count"])
    total = len(sorted_items)

    # Pick from different percentiles and ensure topic diversity
    selected = []
    topics_seen = set()

    # Always include smallest and largest
    if n >= 2:
        selected.append(sorted_items[0])
        selected.append(sorted_items[-1])

    # Pick from quartiles
    indices = [total // 4, total // 2, total * 3 // 4]
    for idx in indices:
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


def main() -> None:
    """Download and validate a subset of documents."""
    manifest_path = CORPUS_DIR / "legal_en_manifest.json"
    with open(manifest_path) as f:
        items = json.load(f)

    print(f"Manifest has {len(items)} items. Picking diverse subset...")
    subset = pick_diverse_subset(items, n=15)

    print(f"\nSelected {len(subset)} documents:")
    for i, item in enumerate(subset, 1):
        print(f"  {i}. {item['filename']:40s} | {item['topic']:15s} | {item['word_count']:5,} words | conf={item['confidence']:.2f}")

    print("\n--- Downloading ---")
    results = []
    for item in subset:
        doc_id = item["id"]
        filename = f"{doc_id}.docx"
        path = DOWNLOAD_DIR / filename

        if path.exists():
            print(f"  {doc_id[:16]}... already exists")
            valid = is_valid_docx(path)
        else:
            print(f"  {doc_id[:16]}... downloading ({item['word_count']:,} words)")
            success = download_file(item["url"], path)
            valid = success and is_valid_docx(path)

        if valid:
            size_kb = path.stat().st_size / 1024
            print(f"    -> Valid .docx ({size_kb:.1f} KB)")
            results.append({"item": item, "path": path, "valid": True})
        else:
            print(f"    -> INVALID")
            if path.exists():
                path.unlink()
            results.append({"item": item, "path": path, "valid": False})

    # Audit valid documents
    valid_results = [r for r in results if r["valid"]]
    print(f"\n--- Auditing {len(valid_results)} valid documents ---")

    for r in valid_results:
        item = r["item"]
        path = r["path"]
        print(f"\n  {item['filename'][:40]:40s} ({item['word_count']:,} words)")
        try:
            doc = DocxDocument(path)
            audit = audit_document(doc)
            print(f"    Headers:     {audit.header_parts} parts, {audit.header_paragraphs} paragraphs")
            print(f"    Footers:     {audit.footer_parts} parts, {audit.footer_paragraphs} paragraphs")
            print(f"    Images:      {audit.images_body} body, {audit.images_header} header, {audit.images_footer} footer")
            print(f"    Tables:      {audit.simple_tables} simple, {audit.skipped_tables} skipped")
            print(f"    Sections:    {audit.section_breaks}")
            print(f"    Multi-col:   {'yes' if audit.multi_column else 'no'}")
            print(f"    Tracked chg: {'yes' if audit.has_tracked_changes else 'no'}")
            print(f"    Comments:    {'yes' if audit.has_comments else 'no'}")
            if audit.unsupported_elements:
                print(f"    Unsupported: {', '.join(audit.unsupported_elements)}")
            # Quick smoke test: can we extract fragments?
            try:
                frags = doc.full_element_map()
                print(f"    Fragments:   {len(frags)} total")
            except Exception as e2:
                print(f"    Fragment extract FAILED: {e2}")
        except Exception as e:
            print(f"    ERROR during audit: {e}")

    # Summary
    print(f"\n\n=== SUMMARY ===")
    print(f"Total attempted: {len(subset)}")
    print(f"Valid .docx:     {len(valid_results)}")
    print(f"Failed/invalid:  {len(subset) - len(valid_results)}")


if __name__ == "__main__":
    main()
