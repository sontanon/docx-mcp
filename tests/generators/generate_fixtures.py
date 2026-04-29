"""Generate .docx test fixtures programmatically using python-docx.

Run directly:
    python -m tests.generators.generate_fixtures

Or invoke via the pytest conftest session fixture (automatic).
"""

from __future__ import annotations

import io
import zipfile
from collections.abc import Callable
from pathlib import Path

from docx import Document
from docx.enum.text import WD_UNDERLINE
from docx.oxml.ns import qn
from lxml import etree

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "generated"


def _ensure_dir() -> Path:
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    return FIXTURES_DIR


def generate_simple_5para() -> Path:
    """5 plain-text paragraphs with no formatting."""
    doc = Document()
    paragraphs = [
        "The Seller agrees to deliver the goods within thirty days of execution.",
        "The Buyer shall make payment in full upon receipt of the goods.",
        "This Agreement shall be governed by the laws of the State of New York.",
        "Neither party shall assign this Agreement without prior written consent.",
        "This Agreement constitutes the entire understanding between the parties.",
    ]
    for text in paragraphs:
        doc.add_paragraph(text)
    path = _ensure_dir() / "simple_5para.docx"
    doc.save(str(path))
    return path


def generate_formatted_runs() -> Path:
    """Paragraphs with mixed bold, italic, and underline runs."""
    doc = Document()

    # Paragraph 1: mixed formatting within a sentence
    p1 = doc.add_paragraph()
    p1.add_run("The ")
    bold_run = p1.add_run("Seller")
    bold_run.bold = True
    p1.add_run(" agrees to deliver the ")
    italic_run = p1.add_run("goods")
    italic_run.italic = True
    p1.add_run(" within ")
    underline_run = p1.add_run("thirty days")
    underline_run.underline = WD_UNDERLINE.SINGLE
    p1.add_run(" of execution.")

    # Paragraph 2: bold + italic combination
    p2 = doc.add_paragraph()
    p2.add_run("Payment shall be made in ")
    bi_run = p2.add_run("United States Dollars")
    bi_run.bold = True
    bi_run.italic = True
    p2.add_run(" upon receipt.")

    # Paragraph 3: all plain text (contrast case)
    doc.add_paragraph("This paragraph has no special formatting and serves as a control case.")

    # Paragraph 4: underline + bold
    p4 = doc.add_paragraph()
    p4.add_run("The ")
    bu_run = p4.add_run("indemnification obligations")
    bu_run.bold = True
    bu_run.underline = WD_UNDERLINE.SINGLE
    p4.add_run(" shall survive termination of this Agreement.")

    # Paragraph 5: multiple formatting switches in one sentence
    p5 = doc.add_paragraph()
    p5.add_run("Notwithstanding the ")
    b = p5.add_run("foregoing")
    b.bold = True
    p5.add_run(", the parties ")
    i = p5.add_run("acknowledge")
    i.italic = True
    p5.add_run(" and ")
    u = p5.add_run("agree")
    u.underline = WD_UNDERLINE.SINGLE
    p5.add_run(" to the following terms.")

    path = _ensure_dir() / "formatted_runs.docx"
    doc.save(str(path))
    return path


def generate_nda_skeleton() -> Path:
    """A realistic NDA skeleton with headings, recitals, numbered sections."""
    doc = Document()

    doc.add_heading("NON-DISCLOSURE AGREEMENT", level=0)

    doc.add_paragraph(
        "This Non-Disclosure Agreement (this \u201cAgreement\u201d) is entered into "
        "as of January 1, 2026 (the \u201cEffective Date\u201d), by and between:"
    )

    # Parties
    doc.add_paragraph("ABC Corporation, a Delaware corporation (\u201cDisclosing Party\u201d); and")
    doc.add_paragraph("XYZ Industries, a California corporation (\u201cReceiving Party\u201d).")

    # Recitals
    doc.add_heading("RECITALS", level=1)
    doc.add_paragraph(
        "WHEREAS, the Disclosing Party possesses certain confidential and proprietary "
        "information relating to its business operations, technology, and trade secrets;"
    )
    doc.add_paragraph(
        "WHEREAS, the Receiving Party desires to receive such confidential information "
        "for the purpose of evaluating a potential business relationship;"
    )
    doc.add_paragraph(
        "NOW, THEREFORE, in consideration of the mutual covenants and agreements "
        "contained herein, the parties agree as follows:"
    )

    # Sections
    doc.add_heading("1. DEFINITION OF CONFIDENTIAL INFORMATION", level=2)
    doc.add_paragraph(
        "\u201cConfidential Information\u201d means any and all non-public information, "
        "whether written, oral, electronic, or visual, disclosed by the Disclosing Party "
        "to the Receiving Party, including but not limited to: trade secrets, business "
        "plans, financial data, customer lists, technical specifications, and software "
        "source code."
    )

    doc.add_heading("2. OBLIGATIONS OF RECEIVING PARTY", level=2)
    doc.add_paragraph(
        "The Receiving Party shall: (a)\u00a0hold all Confidential Information in strict "
        "confidence; (b)\u00a0not disclose any Confidential Information to any third party "
        "without the prior written consent of the Disclosing Party; and (c)\u00a0use the "
        "Confidential Information solely for the purpose of evaluating the potential "
        "business relationship."
    )

    doc.add_heading("3. TERM AND TERMINATION", level=2)
    doc.add_paragraph(
        "This Agreement shall remain in effect for a period of two (2) years from "
        "the Effective Date. Either party may terminate this Agreement upon thirty (30) "
        "days\u2019 prior written notice to the other party."
    )

    doc.add_heading("4. REMEDIES", level=2)
    doc.add_paragraph(
        "The Receiving Party acknowledges that any breach of this Agreement may cause "
        "irreparable harm to the Disclosing Party, and that monetary damages may be "
        "inadequate. Accordingly, the Disclosing Party shall be entitled to seek "
        "equitable relief, including injunction and specific performance, in addition "
        "to all other remedies available at law or in equity."
    )

    doc.add_heading("5. GOVERNING LAW", level=2)
    doc.add_paragraph(
        "This Agreement shall be governed by and construed in accordance with the laws "
        "of the State of Delaware, without regard to its conflict of laws principles."
    )

    # Signature block
    doc.add_paragraph("")  # spacer
    doc.add_paragraph(
        "IN WITNESS WHEREOF, the parties have executed this Agreement "
        "as of the date first written above."
    )

    path = _ensure_dir() / "nda_skeleton.docx"
    doc.save(str(path))
    return path


def generate_styled_headings() -> Path:
    """Document with Heading 1, Heading 2, and body text in different styles."""
    doc = Document()

    doc.add_heading("Master Services Agreement", level=1)
    doc.add_paragraph("This Master Services Agreement governs the provision of services.")

    doc.add_heading("Scope of Services", level=2)
    doc.add_paragraph(
        "The Provider shall perform the services described in each Statement of Work."
    )

    doc.add_heading("Payment Terms", level=2)
    doc.add_paragraph("All invoices shall be paid within forty-five (45) days of receipt.")

    doc.add_heading("Intellectual Property", level=1)
    doc.add_paragraph(
        "All intellectual property developed during the engagement shall be owned "
        "by the Client, subject to the Provider\u2019s pre-existing intellectual property rights."
    )

    path = _ensure_dir() / "styled_headings.docx"
    doc.save(str(path))
    return path


def generate_numbered_list() -> Path:
    """Document with numbered and bulleted list items."""
    doc = Document()

    doc.add_paragraph("The following obligations apply:")

    # Numbered list items using List Number style
    for i, text in enumerate(
        [
            "Maintain confidentiality of all proprietary information.",
            "Return all materials upon termination of the Agreement.",
            "Notify the Disclosing Party of any unauthorized disclosure.",
            "Cooperate fully in any investigation of a breach.",
        ],
        start=1,
    ):
        p = doc.add_paragraph(f"{i}. {text}", style="List Number")  # noqa: F841

    doc.add_paragraph("Additional considerations include:")

    # Bulleted items
    for text in [
        "Applicable regulatory requirements",
        "Industry standard security practices",
        "Third-party audit requirements",
    ]:
        doc.add_paragraph(text, style="List Bullet")

    path = _ensure_dir() / "numbered_list.docx"
    doc.save(str(path))
    return path


def generate_existing_comments() -> Path:
    """Document with 3 pre-existing comments (manually crafted XML)."""
    doc = Document()

    doc.add_paragraph("This clause requires further review by outside counsel.")
    doc.add_paragraph("The indemnification cap should be discussed with the client.")
    doc.add_paragraph("Standard governing law provision for Delaware entities.")

    # Add comments via raw XML manipulation

    # Create the comments part
    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:comment w:id="100" w:author="Senior Partner" w:initials="SP" '
        'w:date="2026-01-15T09:00:00Z">'
        "<w:p><w:r><w:t>Please verify this against the latest template.</w:t></w:r></w:p>"
        "</w:comment>"
        '<w:comment w:id="101" w:author="Associate" w:initials="AS" '
        'w:date="2026-01-16T14:30:00Z">'
        "<w:p><w:r><w:t>Cap amount needs to be confirmed.</w:t></w:r></w:p>"
        "</w:comment>"
        '<w:comment w:id="102" w:author="Senior Partner" w:initials="SP" '
        'w:date="2026-01-17T11:00:00Z">'
        "<w:p><w:r><w:t>Standard provision, no changes needed.</w:t></w:r></w:p>"
        "</w:comment>"
        "</w:comments>"
    )

    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    comments_part = Part(
        PackURI("/word/comments.xml"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        comments_xml.encode("utf-8"),
        doc.part.package,
    )
    doc.part.relate_to(
        comments_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    )

    # Add comment range markers to each paragraph
    body = doc.element.body
    paragraphs = body.findall(qn("w:p"))
    for para, comment_id in zip(paragraphs, [100, 101, 102], strict=False):
        # Insert commentRangeStart before first run
        range_start = para.makeelement(qn("w:commentRangeStart"), {qn("w:id"): str(comment_id)})
        runs = para.findall(qn("w:r"))
        if runs:
            runs[0].addprevious(range_start)

            # Insert commentRangeEnd after last run
            range_end = para.makeelement(qn("w:commentRangeEnd"), {qn("w:id"): str(comment_id)})
            runs[-1].addnext(range_end)

            # Insert comment reference run
            ref_run = para.makeelement(qn("w:r"), {})
            ref_rpr = ref_run.makeelement(qn("w:rPr"), {})
            ref_style = ref_rpr.makeelement(qn("w:rStyle"), {qn("w:val"): "CommentReference"})
            ref_rpr.append(ref_style)
            ref_run.append(ref_rpr)
            comment_ref = ref_run.makeelement(
                qn("w:commentReference"), {qn("w:id"): str(comment_id)}
            )
            ref_run.append(comment_ref)
            range_end.addnext(ref_run)

    path = _ensure_dir() / "existing_comments.docx"
    doc.save(str(path))
    return path


def generate_special_chars() -> Path:
    """Paragraphs with smart quotes, em-dashes, section symbols, non-breaking spaces."""
    doc = Document()

    # Smart quotes and apostrophes
    doc.add_paragraph(
        "The parties to this \u201cAgreement\u201d (as defined herein) "
        "acknowledge the Seller\u2019s obligations."
    )

    # Em-dash and en-dash
    doc.add_paragraph(
        "The confidentiality period\u2014which begins on the Effective Date\u2014"
        "shall last for two years. See Sections\u00a01\u20135 for details."
    )

    # Section symbol and non-breaking spaces
    doc.add_paragraph(
        "Pursuant to \u00a7\u00a05.1 of the Agreement, the Receiving Party shall "
        "comply with all applicable laws."
    )

    # Ellipsis and other unicode
    doc.add_paragraph(
        "The obligations include, without limitation\u2026 maintaining records, "
        "reporting breaches, and cooperating with audits."
    )

    # Mixed: smart quotes with formatting
    p = doc.add_paragraph()
    p.add_run("The term \u201c")
    b = p.add_run("Confidential Information")
    b.bold = True
    p.add_run("\u201d shall have the meaning set forth in \u00a7\u00a01.")

    path = _ensure_dir() / "special_chars.docx"
    doc.save(str(path))
    return path


def generate_long_paragraph() -> Path:
    """Single paragraph with ~300 words and mixed formatting."""
    doc = Document()

    p = doc.add_paragraph()

    p.add_run(
        "Notwithstanding anything to the contrary contained in this Agreement, "
        "the Receiving Party acknowledges and agrees that the Confidential Information "
        "disclosed by the Disclosing Party constitutes valuable trade secrets and "
        "proprietary information of the Disclosing Party. "
    )
    b1 = p.add_run(
        "The Receiving Party shall exercise at least the same degree of care to "
        "protect the Confidential Information as it uses to protect its own confidential "
        "information of a similar nature, but in no event less than reasonable care. "
    )
    b1.bold = True
    p.add_run(
        "The Receiving Party shall limit access to the Confidential Information to those "
        "of its employees, agents, and representatives who have a legitimate need to know "
        "such information for the purposes contemplated by this Agreement, and who have "
        "been informed of the confidential nature of such information and have agreed to "
        "be bound by obligations of confidentiality no less restrictive than those "
        "contained herein. The Receiving Party shall be responsible for any breach of "
        "this Agreement by any of its employees, agents, or representatives. "
    )
    i1 = p.add_run(
        "In the event that the Receiving Party becomes aware of any unauthorized use "
        "or disclosure of Confidential Information, it shall promptly notify the "
        "Disclosing Party in writing and shall cooperate fully with the Disclosing Party "
        "in investigating and remedying such unauthorized use or disclosure. "
    )
    i1.italic = True
    p.add_run(
        "The obligations set forth in this Section shall survive the termination or "
        "expiration of this Agreement for a period of five (5) years from the date "
        "of disclosure of the relevant Confidential Information, or for so long as "
        "such Confidential Information remains a trade secret under applicable law, "
        "whichever period is longer."
    )

    path = _ensure_dir() / "long_paragraph.docx"
    doc.save(str(path))
    return path


def generate_blank_separated() -> Path:
    """Clauses separated by blank paragraphs — for delete_next_blanks tests.

    Layout (fragment IDs):
      1: Clause A text
      2: (blank)
      3: Clause B text
      4: (blank)
      5: Clause C text
      6: (blank)
      7: Clause D text
    """
    doc = Document()
    doc.add_paragraph("The Seller shall deliver the goods within thirty days.")
    doc.add_paragraph("")  # blank separator
    doc.add_paragraph("The Buyer shall pay upon receipt of the goods.")
    doc.add_paragraph("")  # blank separator
    doc.add_paragraph("This Agreement is governed by the laws of New York.")
    doc.add_paragraph("")  # blank separator
    doc.add_paragraph("Neither party may assign without written consent.")
    path = _ensure_dir() / "blank_separated.docx"
    doc.save(str(path))
    return path


def generate_simple_table() -> Path:
    """3x3 table with plain text, paragraphs before and after."""
    doc = Document()

    doc.add_paragraph("Introduction paragraph before the table.")

    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"

    headers = ["Header A", "Header B", "Header C"]
    for col_idx, header in enumerate(headers):
        table.rows[0].cells[col_idx].text = header

    for row_idx in range(1, 3):
        for col_idx in range(3):
            table.rows[row_idx].cells[col_idx].text = f"Cell {row_idx}.{col_idx + 1}"

    doc.add_paragraph("Conclusion paragraph after the table.")

    path = _ensure_dir() / "simple_table.docx"
    doc.save(str(path))
    return path


def generate_formatted_table() -> Path:
    """Table with bold/italic text in cells."""
    doc = Document()

    doc.add_paragraph("Document with formatted table.")

    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"

    cell = table.rows[0].cells[0]
    p = cell.paragraphs[0]
    run = p.add_run("Bold Header")
    run.bold = True

    cell = table.rows[0].cells[1]
    p = cell.paragraphs[0]
    run = p.add_run("Italic Header")
    run.italic = True

    table.rows[1].cells[0].text = "Normal cell"
    table.rows[1].cells[1].text = "Plain text"

    path = _ensure_dir() / "formatted_table.docx"
    doc.save(str(path))
    return path


def generate_table_multi_para() -> Path:
    """Table with cells containing multiple paragraphs."""
    doc = Document()

    doc.add_paragraph("Document with multi-paragraph cells.")

    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"

    cell = table.rows[0].cells[0]
    cell.text = "Header A"

    cell = table.rows[0].cells[1]
    cell.text = "Header B"

    cell = table.rows[1].cells[0]
    p1 = cell.paragraphs[0]
    p1.text = "First paragraph in cell."
    cell.add_paragraph("Second paragraph in cell.")

    cell = table.rows[1].cells[1]
    p1 = cell.paragraphs[0]
    p1.text = "Another cell."
    cell.add_paragraph("With two paragraphs.")
    cell.add_paragraph("And a third one.")

    path = _ensure_dir() / "table_multi_para.docx"
    doc.save(str(path))
    return path


def generate_merged_cell_table() -> Path:
    """Table with horizontal merged cells (should be rejected as non-simple)."""
    doc = Document()

    doc.add_paragraph("Document with merged cells.")

    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"

    table.rows[0].cells[0].text = "A"
    table.rows[0].cells[1].text = "B"

    a = table.rows[0].cells[0]
    b = table.rows[0].cells[1]
    a.merge(b)

    table.rows[0].cells[2].text = "C"
    table.rows[1].cells[0].text = "D"
    table.rows[1].cells[1].text = "E"
    table.rows[1].cells[2].text = "F"

    path = _ensure_dir() / "merged_cell_table.docx"
    doc.save(str(path))
    return path


def generate_mixed_content() -> Path:
    """Multiple paragraphs interspersed with multiple tables."""
    doc = Document()

    doc.add_paragraph("First paragraph before any table.")

    table1 = doc.add_table(rows=2, cols=2)
    table1.style = "Table Grid"
    table1.rows[0].cells[0].text = "T1-A"
    table1.rows[0].cells[1].text = "T1-B"
    table1.rows[1].cells[0].text = "T1-C"
    table1.rows[1].cells[1].text = "T1-D"

    doc.add_paragraph("Second paragraph between tables.")

    table2 = doc.add_table(rows=1, cols=3)
    table2.style = "Table Grid"
    table2.rows[0].cells[0].text = "X"
    table2.rows[0].cells[1].text = "Y"
    table2.rows[0].cells[2].text = "Z"

    doc.add_paragraph("Third paragraph after all tables.")

    path = _ensure_dir() / "mixed_content.docx"
    doc.save(str(path))
    return path


# ---------------------------------------------------------------------------
# Tracked-change fixtures (for T1.1 rejection tests)
# ---------------------------------------------------------------------------


def _inject_tracked_change_into_paragraph(paragraph: etree._Element, tag: str) -> None:
    """Inject a tracked-change wrapper around the first run in a paragraph."""
    runs = paragraph.findall(qn("w:r"))
    if not runs:
        # Create a dummy run if none exists
        run = etree.SubElement(paragraph, qn("w:r"))
        t = etree.SubElement(run, qn("w:t"))
        t.text = "tracked"
        runs = [run]

    # Build wrapper: <w:ins w:id="1" w:author="Test" w:date="2026-01-01T00:00:00Z">
    wrapper = etree.Element(qn(f"w:{tag}"))
    wrapper.set(qn("w:id"), "1")
    wrapper.set(qn("w:author"), "Test")
    wrapper.set(qn("w:date"), "2026-01-01T00:00:00Z")

    # Move all runs inside the wrapper
    for run in runs:
        wrapper.append(run)

    # Insert wrapper where first_run was
    paragraph.insert(0, wrapper)


def generate_body_tracked_changes() -> Path:
    """Document with <w:ins> inside a body paragraph."""
    doc = Document()
    doc.add_paragraph("This paragraph has tracked changes.")
    doc.add_paragraph("This one is clean.")

    # Post-process XML
    tree = doc.element
    body = tree.body
    paragraphs = body.findall(qn("w:p"))
    _inject_tracked_change_into_paragraph(paragraphs[0], "ins")

    path = _ensure_dir() / "body_tracked_changes.docx"
    doc.save(str(path))
    return path


def generate_header_tracked_changes() -> Path:
    """Document with <w:del> inside a header paragraph."""
    doc = Document()
    doc.add_paragraph("Body paragraph.")

    # Add a header
    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_para.text = "Header with tracked change."

    doc.save(str(path := _ensure_dir() / "header_tracked_changes_tmp.docx"))

    # Post-process: inject <w:del> into header XML
    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    # Find header file
    header_entry = None
    for name in entries:
        if name.startswith("word/header") and name.endswith(".xml"):
            header_entry = name
            break

    if header_entry:
        header_tree = etree.fromstring(entries[header_entry])
        paragraphs = header_tree.findall(f".//{{{etree.QName(qn('w:p')).namespace}}}p")
        if paragraphs:
            _inject_tracked_change_into_paragraph(paragraphs[0], "del")
        entries[header_entry] = etree.tostring(
            header_tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    final_path = _ensure_dir() / "header_tracked_changes.docx"
    final_path.write_bytes(buf.getvalue())
    path.unlink()
    return final_path


def generate_footer_tracked_changes() -> Path:
    """Document with <w:moveFrom> inside a footer paragraph."""
    doc = Document()
    doc.add_paragraph("Body paragraph.")

    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.text = "Footer with tracked change."

    doc.save(str(path := _ensure_dir() / "footer_tracked_changes_tmp.docx"))

    with zipfile.ZipFile(path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    footer_entry = None
    for name in entries:
        if name.startswith("word/footer") and name.endswith(".xml"):
            footer_entry = name
            break

    if footer_entry:
        footer_tree = etree.fromstring(entries[footer_entry])
        paragraphs = footer_tree.findall(f".//{{{etree.QName(qn('w:p')).namespace}}}p")
        if paragraphs:
            _inject_tracked_change_into_paragraph(paragraphs[0], "moveFrom")
        entries[footer_entry] = etree.tostring(
            footer_tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    final_path = _ensure_dir() / "footer_tracked_changes.docx"
    final_path.write_bytes(buf.getvalue())
    path.unlink()
    return final_path


def generate_comments_tracked_changes() -> Path:
    """Document with <w:moveTo> inside comments.xml."""
    doc = Document()
    doc.add_paragraph("Body paragraph.")

    # Add comments part with a tracked change inside
    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:comment w:id="1" w:author="Test" w:initials="T" '
        'w:date="2026-01-01T00:00:00Z">'
        "<w:p>"
        "<w:r><w:t>Clean text.</w:t></w:r>"
        '<w:moveTo w:id="2" w:author="Test" w:date="2026-01-01T00:00:00Z">'
        "<w:r><w:t>Moved text.</w:t></w:r>"
        "</w:moveTo>"
        "</w:p>"
        "</w:comment>"
        "</w:comments>"
    )

    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    comments_part = Part(
        PackURI("/word/comments.xml"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        comments_xml.encode("utf-8"),
        doc.part.package,
    )
    doc.part.relate_to(
        comments_part,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    )

    path = _ensure_dir() / "comments_tracked_changes.docx"
    doc.save(str(path))
    return path


def _add_hyperlink_to_run(
    paragraph,
    run_index: int,
    url: str,
    doc,
) -> None:
    """Post-process a python-docx paragraph to wrap a run in a hyperlink.

    Inserts a ``<w:hyperlink>`` wrapper around the run at *run_index*
    and creates the relationship in the document's rels.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    # Get the run element (only direct children, not descendants inside hyperlinks)
    p_element = paragraph._element
    runs = [child for child in p_element if etree.QName(child.tag).localname == "r"]
    target_run = runs[run_index]

    # Create relationship
    part = doc.part
    r_id = part.relate_to(
        url,
        RT.HYPERLINK,
        is_external=True,
    )

    # Build hyperlink wrapper
    hyperlink = etree.Element(qn("w:hyperlink"))
    hyperlink.set(qn("r:id"), r_id)

    # Move run into hyperlink
    target_run.addprevious(hyperlink)
    hyperlink.append(target_run)


def generate_hyperlink_paragraph() -> Path:
    """Paragraph with a plain hyperlink."""
    doc = Document()
    p = doc.add_paragraph("Visit ")
    p.add_run("our website")
    p.add_run(" for more info.")

    _add_hyperlink_to_run(p, 1, "https://example.com", doc)

    path = _ensure_dir() / "hyperlink_paragraph.docx"
    doc.save(str(path))
    return path


def generate_hyperlink_formatted() -> Path:
    """Paragraph with hyperlink containing bold and italic text."""
    doc = Document()
    p = doc.add_paragraph("See ")
    run = p.add_run("important terms")
    run.bold = True
    run.italic = True
    p.add_run(" for details.")

    _add_hyperlink_to_run(p, 1, "https://example.com/terms", doc)

    path = _ensure_dir() / "hyperlink_formatted.docx"
    doc.save(str(path))
    return path


def generate_multiple_hyperlinks() -> Path:
    """Paragraph with two separate hyperlinks."""
    doc = Document()
    p = doc.add_paragraph("Contact ")
    p.add_run("sales")
    p.add_run(" or ")
    p.add_run("support")
    p.add_run(".")

    # Add in reverse order so indices don't shift after wrapping
    _add_hyperlink_to_run(p, 3, "https://example.com/support", doc)
    _add_hyperlink_to_run(p, 1, "https://example.com/sales", doc)

    path = _ensure_dir() / "multiple_hyperlinks.docx"
    doc.save(str(path))
    return path


def generate_header_footer_text() -> Path:
    """Document with plain text in header and footer."""
    doc = Document()
    doc.add_paragraph("First body paragraph.")
    doc.add_paragraph("Second body paragraph.")

    section = doc.sections[0]
    header = section.header
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_para.text = "Header paragraph text."

    footer = section.footer
    footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_para.text = "Footer paragraph text."

    path = _ensure_dir() / "header_footer_text.docx"
    doc.save(str(path))
    return path


def generate_two_section() -> Path:
    """Document with two sections separated by a section break."""
    doc = Document()
    doc.add_paragraph("Paragraph in first section.")

    # Add a section break with different page size
    new_section = doc.add_section()
    new_section.page_width = 15840000  # landscape-ish (in EMUs)
    new_section.page_height = 10080000
    doc.add_paragraph("Paragraph in second section.")

    path = _ensure_dir() / "two_section.docx"
    doc.save(str(path))
    return path


def generate_table_empty_cell() -> Path:
    """Simple table with one empty cell."""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = "Header A"
    table.rows[0].cells[1].text = "Header B"
    table.rows[1].cells[0].text = "Data 1"
    # Cell 1.2 is intentionally empty

    path = _ensure_dir() / "table_empty_cell.docx"
    doc.save(str(path))
    return path


def generate_wide_table() -> Path:
    """10-column table for stress testing."""
    doc = Document()
    table = doc.add_table(rows=2, cols=10)
    headers = [f"Col {i}" for i in range(1, 11)]
    for col_idx, header in enumerate(headers):
        table.rows[0].cells[col_idx].text = header
    for col_idx in range(10):
        table.rows[1].cells[col_idx].text = f"Data {col_idx + 1}"

    path = _ensure_dir() / "wide_table.docx"
    doc.save(str(path))
    return path


ALL_GENERATORS: list[tuple[str, Callable[[], Path]]] = [
    ("simple_5para", generate_simple_5para),
    ("formatted_runs", generate_formatted_runs),
    ("nda_skeleton", generate_nda_skeleton),
    ("styled_headings", generate_styled_headings),
    ("numbered_list", generate_numbered_list),
    ("existing_comments", generate_existing_comments),
    ("special_chars", generate_special_chars),
    ("long_paragraph", generate_long_paragraph),
    ("blank_separated", generate_blank_separated),
    ("simple_table", generate_simple_table),
    ("formatted_table", generate_formatted_table),
    ("table_multi_para", generate_table_multi_para),
    ("merged_cell_table", generate_merged_cell_table),
    ("mixed_content", generate_mixed_content),
    ("body_tracked_changes", generate_body_tracked_changes),
    ("header_tracked_changes", generate_header_tracked_changes),
    ("footer_tracked_changes", generate_footer_tracked_changes),
    ("comments_tracked_changes", generate_comments_tracked_changes),
    ("hyperlink_paragraph", generate_hyperlink_paragraph),
    ("hyperlink_formatted", generate_hyperlink_formatted),
    ("multiple_hyperlinks", generate_multiple_hyperlinks),
    ("header_footer_text", generate_header_footer_text),
    ("two_section", generate_two_section),
    ("table_empty_cell", generate_table_empty_cell),
    ("wide_table", generate_wide_table),
]


def generate_all() -> dict[str, Path]:
    """Generate all test fixtures. Returns a mapping of name -> path."""
    results = {}
    for name, gen_fn in ALL_GENERATORS:
        path = gen_fn()
        results[name] = path
        print(f"  Generated: {path}")
    return results


if __name__ == "__main__":
    print("Generating test fixtures...")
    paths = generate_all()
    print(f"\nGenerated {len(paths)} fixtures in {FIXTURES_DIR}")
