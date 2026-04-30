#!/usr/bin/env python3
"""Forensic document analysis script for flagged .docx files."""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

CORPUS_DIR = Path("/home/santiago/Code/docx-mcp/corpus")
DOWNLOADED_DIR = CORPUS_DIR / "downloaded"
REPORT_PATH = CORPUS_DIR / "forensic_report.json"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}


def get_text_from_element(element: ET.Element | None) -> str:
    """Extract text from w:t elements within an XML element."""
    if element is None:
        return ""
    texts = [t.text for t in element.findall(".//w:t", NS) if t.text]
    return "".join(texts).strip()


def count_words(text: str) -> int:
    """Count words using whitespace splitting."""
    return len(text.split()) if text else 0


def analyze_document(doc_path: Path) -> dict:
    """Perform deep forensic analysis on a single .docx file."""
    findings = {
        "zip_valid": False,
        "word_dir_contents": [],
        "document_xml_paragraphs": 0,
        "document_xml_text_runs": 0,
        "direct_body_paragraphs": 0,
        "table_paragraphs": 0,
        "headers": 0,
        "footers": 0,
        "footnotes": 0,
        "endnotes": 0,
        "comments": 0,
        "altChunk": False,
        "embedded_objects": False,
        "content_controls": False,
        "images_with_text": False,
        "drawingml_textboxes": False,
        "vml_textboxes": False,
        "core_title": "",
        "largest_text_block": "",
        "total_body_text": "",
        "total_table_text": "",
        "unusual_elements": [],
    }

    try:
        with zipfile.ZipFile(doc_path, "r") as zf:
            findings["zip_valid"] = True
            word_files = [
                {"name": info.filename, "size": info.file_size}
                for info in zf.infolist()
                if info.filename.startswith("word/")
            ]
            findings["word_dir_contents"] = word_files

            # Parse document.xml
            doc_xml = None
            if "word/document.xml" in zf.namelist():
                with zf.open("word/document.xml") as f:
                    doc_xml = ET.parse(f).getroot()

                body = doc_xml.find(".//w:body", NS)
                if body is not None:
                    # Direct body paragraphs vs table paragraphs
                    direct_p = body.findall("w:p", NS)
                    tbl_p = body.findall(".//w:tbl//w:p", NS)
                    findings["direct_body_paragraphs"] = len(direct_p)
                    findings["table_paragraphs"] = len(tbl_p)
                    findings["document_xml_paragraphs"] = len(direct_p) + len(tbl_p)
                    findings["document_xml_text_runs"] = len(body.findall(".//w:t", NS))

                    # Extract text from tables separately
                    tbl_text_elem = ET.Element("dummy")
                    for tbl in body.findall("w:tbl", NS):
                        tbl_text_elem.append(tbl)
                    table_text = get_text_from_element(tbl_text_elem)

                    # Direct body text (after removing tables)
                    direct_text = get_text_from_element(body)

                    findings["total_body_text"] = direct_text
                    findings["total_table_text"] = table_text

                    # Largest text block
                    max_text = ""
                    for p in body.findall(".//w:p", NS):
                        p_text = get_text_from_element(p)
                        if len(p_text) > len(max_text):
                            max_text = p_text
                    findings["largest_text_block"] = max_text[:300]

                # Unusual elements
                if doc_xml.find(".//w:altChunk", NS) is not None:
                    findings["altChunk"] = True
                    findings["unusual_elements"].append("altChunk")
                if doc_xml.find(".//w:customXml", NS) is not None:
                    findings["unusual_elements"].append("customXml")
                if doc_xml.find(".//mc:AlternateContent", NS) is not None:
                    findings["unusual_elements"].append("mc:AlternateContent")
                if doc_xml.find(".//w:object", NS) is not None:
                    findings["embedded_objects"] = True
                    findings["unusual_elements"].append("embedded_objects")
                if doc_xml.find(".//o:OLEObject", NS) is not None:
                    findings["embedded_objects"] = True
                    findings["unusual_elements"].append("embedded_objects")
                if doc_xml.find(".//w:sdt", NS) is not None:
                    findings["content_controls"] = True
                    findings["unusual_elements"].append("content_controls")
                if doc_xml.find(".//a:p", NS) is not None:
                    findings["images_with_text"] = True
                    findings["drawingml_textboxes"] = True
                    findings["unusual_elements"].append("DrawingML_textbox")
                if doc_xml.find(".//v:textbox", NS) is not None:
                    findings["vml_textboxes"] = True
                    findings["unusual_elements"].append("VML_textbox")
                if doc_xml.find(".//wp:docPr", NS) is not None:
                    for dp in doc_xml.findall(".//wp:docPr", NS):
                        if dp.get("descr") or dp.get("title"):
                            findings["images_with_text"] = True
                            findings["unusual_elements"].append("image_description")
                            break

            # Headers and footers
            header_words = 0
            footer_words = 0
            for name in zf.namelist():
                if re.match(r"word/header\d+\.xml", name):
                    with zf.open(name) as f:
                        hroot = ET.parse(f).getroot()
                        htext = get_text_from_element(hroot)
                        if htext:
                            header_words += count_words(htext)
                elif re.match(r"word/footer\d+\.xml", name):
                    with zf.open(name) as f:
                        froot = ET.parse(f).getroot()
                        ftext = get_text_from_element(froot)
                        if ftext:
                            footer_words += count_words(ftext)

            findings["headers"] = header_words
            findings["footers"] = footer_words

            # Footnotes, endnotes, comments
            for part, key in [
                ("word/footnotes.xml", "footnotes"),
                ("word/endnotes.xml", "endnotes"),
                ("word/comments.xml", "comments"),
            ]:
                if part in zf.namelist():
                    with zf.open(part) as f:
                        proot = ET.parse(f).getroot()
                        ptext = get_text_from_element(proot)
                        findings[key] = count_words(ptext)

            # Core properties
            if "docProps/core.xml" in zf.namelist():
                with zf.open("docProps/core.xml") as f:
                    croot = ET.parse(f).getroot()
                    title_elem = croot.find("dc:title", NS)
                    if title_elem is not None and title_elem.text:
                        findings["core_title"] = title_elem.text.strip()

            # Settings
            if "word/settings.xml" in zf.namelist():
                with zf.open("word/settings.xml") as f:
                    sroot = ET.parse(f).getroot()
                    if sroot.find(".//w:documentProtection", NS) is not None:
                        findings["unusual_elements"].append("document_protection")

            # Glossary
            if "word/glossaryDocument.xml" in zf.namelist():
                findings["unusual_elements"].append("glossaryDocument")

    except zipfile.BadZipFile:
        findings["zip_valid"] = False
    except ET.ParseError as e:
        findings["unusual_elements"].append(f"XML_parse_error: {e}")
    except Exception as e:
        findings["unusual_elements"].append(f"error: {e}")

    return findings


def classify(findings: dict, info: dict) -> tuple[str, str, str, str]:
    """Return (classification, confidence, explanation, recommendation)."""
    direct_words = count_words(findings["total_body_text"])
    table_words = count_words(findings["total_table_text"])
    total_words = direct_words + table_words
    alt_text_words = findings["headers"] + findings["footers"] + findings["footnotes"] + findings["endnotes"] + findings["comments"]

    # Parsing issue: content in tables that parser missed
    if findings["table_paragraphs"] > 10 and (table_words > 50 or findings["table_paragraphs"] > 100):
        return (
            "parsing_issue",
            "high",
            f"Document has {findings['table_paragraphs']} paragraphs inside tables ({table_words} words) but only {findings['direct_body_paragraphs']} direct body paragraphs ({direct_words} words). Our parser only counts top-level body paragraphs and misses table content entirely.",
            "fix_parser_table_support",
        )

    # Parsing issue: text in DrawingML/VML text boxes
    if (findings["drawingml_textboxes"] or findings["vml_textboxes"]) and total_words > 0 and total_words < 50:
        return (
            "parsing_issue",
            "high",
            f"Document text resides in DrawingML/VML text boxes ({total_words} words) rather than direct body paragraphs. Our parser does not descend into shape text boxes.",
            "fix_parser_textbox_support",
        )

    # Parsing issue: embedded objects
    if findings["embedded_objects"]:
        return (
            "parsing_issue",
            "high",
            "Document contains embedded OLE objects which may hold substantial text not present in body paragraphs.",
            "investigate_pipeline",
        )

    # Parsing issue: altChunk
    if findings["altChunk"]:
        return (
            "parsing_issue",
            "high",
            "Document contains altChunk elements referencing external content that our parser does not resolve.",
            "investigate_pipeline",
        )

    # Parsing issue: substantial text in headers/footers/notes/comments
    if alt_text_words > 20 and total_words < 50:
        return (
            "parsing_issue",
            "medium",
            f"Document body is minimal ({total_words} words) but headers/footers/notes/comments contain {alt_text_words} words of text that may not be extracted.",
            "investigate_pipeline",
        )

    # Edge case: explicitly a template
    if "template" in info["filename"].lower():
        return (
            "edge_case",
            "high",
            "File is explicitly a template shell with minimal boilerplate text. Intentionally blank for user fill-in.",
            "keep_for_testing",
        )

    # Edge case: form with content controls
    if info.get("type") == "forms" and findings["content_controls"] and total_words < 50:
        return (
            "edge_case",
            "medium",
            "Form document with content controls. Minimal visible text is expected; fields are meant to be filled.",
            "keep_for_testing",
        )

    # Truly bad: minimal everything
    if total_words < 30 and alt_text_words == 0 and not findings["unusual_elements"]:
        return (
            "truly_bad",
            "high",
            f"Document has only {total_words} words in {findings['document_xml_paragraphs']} paragraphs, no headers/footers/notes, and no alternative content sources. Truly minimal.",
            "remove",
        )

    if total_words == 0 and findings["document_xml_paragraphs"] > 0:
        return (
            "truly_bad",
            "high",
            "Document has paragraphs but zero extractable text (all empty or whitespace-only).",
            "remove",
        )

    # Truly bad: very minimal even with minor header/footer text
    if total_words < 15 and alt_text_words < 25 and not findings["unusual_elements"]:
        return (
            "truly_bad",
            "high",
            f"Document has only {total_words} words in {findings['document_xml_paragraphs']} paragraphs with negligible header/footer text ({alt_text_words} words). No substantive content.",
            "remove",
        )

    # Fallback
    return (
        "investigate_manually",
        "medium",
        f"Document has {total_words} words, {findings['document_xml_paragraphs']} paragraphs, and unusual elements: {findings['unusual_elements']}. Needs human review.",
        "investigate_manually",
    )


def main() -> None:
    with open(CORPUS_DIR / "quality_report.json") as f:
        quality_report = json.load(f)

    flagged = quality_report["flagged_for_removal"]
    results = []
    summary = {"truly_bad": 0, "parsing_issue": 0, "edge_case": 0, "investigate_manually": 0}

    for item in flagged:
        doc_id = item["id"]
        doc_path = DOWNLOADED_DIR / f"{doc_id}.docx"
        print(f"Analyzing {doc_id} ({item['filename']}) ...")

        findings = analyze_document(doc_path)
        classification, confidence, explanation, recommendation = classify(findings, item)

        results.append({
            "id": doc_id,
            "path": str(doc_path.relative_to(CORPUS_DIR)),
            "classification": classification,
            "confidence": confidence,
            "findings": findings,
            "explanation": explanation,
            "recommendation": recommendation,
        })
        summary[classification] += 1

    report = {
        "flagged_documents": results,
        "summary": summary,
    }

    with open(REPORT_PATH, "w") as f:
        json.dump(report, f, indent=2)

    print(f"\nReport written to {REPORT_PATH}")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
