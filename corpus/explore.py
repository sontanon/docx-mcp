"""Explore the docx-corpus dataset from HuggingFace."""

import json
from collections import Counter
from pathlib import Path

from datasets import load_dataset  # type: ignore

CORPUS_DIR = Path(__file__).parent


def explore_dataset() -> None:
    """Load docx-corpus metadata and print statistics."""
    print("Loading docx-corpus dataset...")
    ds = load_dataset("superdoc-dev/docx-corpus", split="train")

    total = len(ds)
    print(f"Total documents: {total:,}")

    # Overall distribution
    print("\n--- Document type distribution ---")
    type_counts = Counter(row["type"] for row in ds)
    for doc_type, count in type_counts.most_common():
        pct = count / total * 100
        print(f"  {doc_type:20s}: {count:6,} ({pct:5.1f}%)")

    print("\n--- Topic distribution ---")
    topic_counts = Counter(row["topic"] for row in ds)
    for topic, count in topic_counts.most_common():
        pct = count / total * 100
        print(f"  {topic:20s}: {count:6,} ({pct:5.1f}%)")

    print("\n--- Language distribution (top 20) ---")
    lang_counts = Counter(row["language"] for row in ds)
    for lang, count in lang_counts.most_common(20):
        pct = count / total * 100
        print(f"  {lang:5s}: {count:6,} ({pct:5.1f}%)")

    # Focus: legal documents
    print("\n\n=== LEGAL DOCUMENTS DEEP DIVE ===")
    legal = ds.filter(lambda x: x["type"] == "legal")
    print(f"Total legal documents: {len(legal):,}")

    print("\n--- Legal by topic ---")
    legal_topic_counts = Counter(row["topic"] for row in legal)
    for topic, count in legal_topic_counts.most_common():
        pct = count / len(legal) * 100
        print(f"  {topic:20s}: {count:6,} ({pct:5.1f}%)")

    print("\n--- Legal by language (top 20) ---")
    legal_lang_counts = Counter(row["language"] for row in legal)
    for lang, count in legal_lang_counts.most_common(20):
        pct = count / len(legal) * 100
        print(f"  {lang:5s}: {count:6,} ({pct:5.1f}%)")

    print("\n--- Legal word count distribution ---")
    word_counts = sorted(row["word_count"] for row in legal)
    print(f"  Min: {min(word_counts):,}")
    print(f"  Max: {max(word_counts):,}")
    print(f"  Median: {word_counts[len(word_counts) // 2]:,}")
    print(f"  Mean: {sum(word_counts) / len(word_counts):,.0f}")
    print(f"  25th percentile: {word_counts[len(word_counts) // 4]:,}")
    print(f"  75th percentile: {word_counts[len(word_counts) * 3 // 4]:,}")

    print("\n--- Legal confidence distribution ---")
    confidences = [row["confidence"] for row in legal]
    print(f"  Min: {min(confidences):.3f}")
    print(f"  Max: {max(confidences):.3f}")
    print(f"  Mean: {sum(confidences) / len(confidences):.3f}")

    # Filter: English legal documents with high confidence
    print("\n\n=== ENGLISH LEGAL HIGH-CONFIDENCE SUBSET ===")
    legal_en_high = ds.filter(
        lambda x: x["type"] == "legal"
        and x["language"] == "en"
        and x["confidence"] >= 0.8,
    )
    print(f"Count: {len(legal_en_high):,}")

    print("\n--- Word count distribution (EN legal, conf>=0.8) ---")
    wc = sorted(row["word_count"] for row in legal_en_high)
    print(f"  Min: {min(wc):,}")
    print(f"  Max: {max(wc):,}")
    print(f"  Median: {wc[len(wc) // 2]:,}")
    print(f"  Mean: {sum(wc) / len(wc):,.0f}")

    # Save sample metadata
    sample_size = min(50, len(legal_en_high))
    print(f"\n--- Saving top {sample_size} sample metadata ---")
    sample = legal_en_high.select(range(sample_size))
    sample_data = []
    for row in sample:
        sample_data.append(
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

    manifest_path = CORPUS_DIR / "legal_en_manifest.json"
    with open(manifest_path, "w") as f:
        json.dump(sample_data, f, indent=2)
    print(f"Saved to {manifest_path}")

    # Also look at other types relevant to legal work
    print("\n\n=== OTHER RELEVANT TYPES ===")
    relevant_types = ["forms", "policies", "reports", "correspondence"]
    for rtype in relevant_types:
        subset = ds.filter(
            lambda x, rtype=rtype: (
                x["type"] == rtype
                and x["language"] == "en"
                and x["confidence"] >= 0.8
            ),
        )
        print(f"\n--- {rtype.upper()} (en, conf>=0.8) ---")
        print(f"  Count: {len(subset):,}")
        if len(subset) > 0:
            topics = Counter(row["topic"] for row in subset)
            print(f"  Top topics: {dict(topics.most_common(5))}")
            wc = [row["word_count"] for row in subset]
            print(f"  Avg words: {sum(wc) / len(wc):,.0f}")


if __name__ == "__main__":
    explore_dataset()
