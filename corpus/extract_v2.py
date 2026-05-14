import sys
from pathlib import Path

# Ensure src is on path
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from docx_mcp.converter import body_to_fragments, fragments_to_tagged_text_interleaved
from docx_mcp.document import DocxDocument

DOWNLOADED = Path("corpus/downloaded")
OUTPUT = Path("corpus/extracted_text_v2")
OUTPUT.mkdir(parents=True, exist_ok=True)

docx_files = sorted(DOWNLOADED.glob("*.docx"))
print(f"Extracting {len(docx_files)} documents...")

for docx_path in docx_files:
    try:
        doc = DocxDocument(path=docx_path)
        result = body_to_fragments(doc.body_elements, collapse_empty=False)
        tagged = fragments_to_tagged_text_interleaved(result.items)
        out_path = OUTPUT / f"{docx_path.stem}.txt"
        out_path.write_text(tagged, encoding="utf-8")
        print(f"  OK: {docx_path.name} -> {len(tagged)} chars")
    except Exception as exc:
        print(f"  ERR: {docx_path.name} -> {exc}")

print("Done.")
