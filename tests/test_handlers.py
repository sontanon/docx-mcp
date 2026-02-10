"""Tests for the change handlers (delete, append, modify)."""

from __future__ import annotations

from lxml import etree

from docx_mcp.converter import paragraph_to_pseudo_markdown
from docx_mcp.document import DocxDocument
from docx_mcp.handlers.append import handle_append_after
from docx_mcp.handlers.delete import handle_delete
from docx_mcp.handlers.modify import handle_modify
from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig
from docx_mcp.namespaces import XML, qn, xpath

# ===================================================================
# Helpers
# ===================================================================


def _make_paragraph(*run_specs: tuple[str, dict[str, bool] | None]) -> etree._Element:
    """Build a ``<w:p>`` element with runs (same helper as test_run_ops)."""
    p = etree.Element(qn("w", "p"))
    for text, fmt in run_specs:
        r = etree.SubElement(p, qn("w", "r"))
        if fmt:
            rpr = etree.SubElement(r, qn("w", "rPr"))
            if fmt.get("bold"):
                etree.SubElement(rpr, qn("w", "b"))
            if fmt.get("italic"):
                etree.SubElement(rpr, qn("w", "i"))
            if fmt.get("underline"):
                u = etree.SubElement(rpr, qn("w", "u"))
                u.set(qn("w", "val"), "single")
        t = etree.SubElement(r, qn("w", "t"))
        t.text = text
        if text and (text[0] == " " or text[-1] == " "):
            t.set(f"{{{XML}}}space", "preserve")
    return p


def _make_body_with_paragraphs(
    *para_specs: list[tuple[str, dict[str, bool] | None]],
) -> etree._Element:
    """Build a ``<w:body>`` with multiple paragraphs."""
    body = etree.Element(qn("w", "body"))
    for spec in para_specs:
        p = _make_paragraph(*spec)
        body.append(p)
    return body


def _default_config() -> RedlineConfig:
    return RedlineConfig(author="Test Author")


# ===================================================================
# Delete handler tests
# ===================================================================


class TestHandleDelete:
    def test_wraps_runs_in_del(self):
        p = _make_paragraph(("Hello world", None))
        mgr = IdManager()
        handle_delete(p, id_manager=mgr, config=_default_config())

        # Should have a <w:del> element
        del_els = xpath(p, "w:del")
        assert len(del_els) == 1

        # The <w:del> should contain runs
        runs = xpath(del_els[0], "w:r")
        assert len(runs) >= 1

    def test_converts_t_to_del_text(self):
        p = _make_paragraph(("Hello world", None))
        mgr = IdManager()
        handle_delete(p, id_manager=mgr, config=_default_config())

        # All <w:t> should be converted to <w:delText>
        del_el = xpath(p, "w:del")[0]
        t_elements = xpath(del_el, "w:r/w:t")
        assert len(t_elements) == 0

        del_text_elements = xpath(del_el, "w:r/w:delText")
        assert len(del_text_elements) >= 1
        assert del_text_elements[0].text == "Hello world"

    def test_marks_paragraph_mark_deleted(self):
        p = _make_paragraph(("Hello", None))
        mgr = IdManager()
        handle_delete(p, id_manager=mgr, config=_default_config())

        # <w:pPr><w:rPr><w:del> should exist
        ppr_del = xpath(p, "w:pPr/w:rPr/w:del")
        assert len(ppr_del) == 1

    def test_del_has_correct_attributes(self):
        config = RedlineConfig(author="Lawyer X")
        p = _make_paragraph(("Hello", None))
        mgr = IdManager(start_after=50)
        del_id = handle_delete(p, id_manager=mgr, config=config)

        assert del_id == 51
        del_el = xpath(p, "w:del")[0]
        assert del_el.get(qn("w", "id")) == "51"
        assert del_el.get(qn("w", "author")) == "Lawyer X"

    def test_multiple_runs(self):
        p = _make_paragraph(
            ("The ", None),
            ("Company", {"bold": True}),
            (" shall", None),
        )
        mgr = IdManager()
        handle_delete(p, id_manager=mgr, config=_default_config())

        del_el = xpath(p, "w:del")[0]
        runs = xpath(del_el, "w:r")
        assert len(runs) == 3

    def test_returns_annotation_id(self):
        p = _make_paragraph(("Hello", None))
        mgr = IdManager(start_after=99)
        del_id = handle_delete(p, id_manager=mgr, config=_default_config())
        assert del_id == 100


# ===================================================================
# Append handler tests
# ===================================================================


class TestHandleAppendAfter:
    def test_inserts_paragraph_after_reference(self):
        body = _make_body_with_paragraphs(
            [("First paragraph.", None)],
            [("Third paragraph.", None)],
        )
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ins_id = handle_append_after(
            ref_p,
            "Second paragraph.",
            id_manager=mgr,
            config=_default_config(),
        )

        # The body should now have 3 paragraphs
        paragraphs = list(body)
        assert len(paragraphs) == 3
        # New paragraph should be at index 1
        assert paragraphs[1] is new_p

    def test_marks_paragraph_mark_as_inserted(self):
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ = handle_append_after(
            ref_p,
            "New text.",
            id_manager=mgr,
            config=_default_config(),
        )

        ppr_ins = xpath(new_p, "w:pPr/w:rPr/w:ins")
        assert len(ppr_ins) == 1

    def test_wraps_content_in_ins(self):
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ = handle_append_after(
            ref_p,
            "New content here.",
            id_manager=mgr,
            config=_default_config(),
        )

        ins_els = xpath(new_p, "w:ins")
        assert len(ins_els) >= 1

        # Content should be inside <w:ins>
        runs = xpath(ins_els[0], "w:r")
        assert len(runs) >= 1

    def test_plain_text_content(self):
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ = handle_append_after(
            ref_p,
            "Simple text.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Extract text from the new paragraph
        texts = []
        for t in new_p.iter(qn("w", "t")):
            if t.text:
                texts.append(t.text)
        assert "Simple text." in " ".join(texts)

    def test_formatted_text_content(self):
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ = handle_append_after(
            ref_p,
            "The **Company** shall.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should have multiple runs (plain + bold + plain)
        ins_el = xpath(new_p, "w:ins")[0]
        runs = xpath(ins_el, "w:r")
        assert len(runs) >= 1

    def test_returns_annotation_id(self):
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager(start_after=200)
        _, ins_id = handle_append_after(
            ref_p,
            "Text.",
            id_manager=mgr,
            config=_default_config(),
        )
        assert ins_id == 201

    def test_ins_attributes(self):
        config = RedlineConfig(author="AI Bot")
        body = _make_body_with_paragraphs([("Ref.", None)])
        ref_p = body[0]
        mgr = IdManager()
        new_p, _ = handle_append_after(
            ref_p,
            "Text.",
            id_manager=mgr,
            config=config,
        )

        ins_el = xpath(new_p, "w:ins")[0]
        assert ins_el.get(qn("w", "author")) == "AI Bot"


# ===================================================================
# Modify handler tests
# ===================================================================


class TestHandleModify:
    def test_no_change_returns_empty(self):
        p = _make_paragraph(("Hello world", None))
        mgr = IdManager()
        ids = handle_modify(p, "Hello world", id_manager=mgr, config=_default_config())
        assert ids == []

    def test_word_replacement_produces_del_and_ins(self):
        p = _make_paragraph(("The quick brown fox", None))
        mgr = IdManager()
        ids = handle_modify(
            p,
            "The slow brown fox",
            id_manager=mgr,
            config=_default_config(),
        )

        assert len(ids) >= 2  # At least one del + one ins

        # Should have <w:del> and <w:ins> elements
        del_els = xpath(p, "w:del")
        ins_els = xpath(p, "w:ins")
        assert len(del_els) >= 1
        assert len(ins_els) >= 1

    def test_del_contains_old_text(self):
        p = _make_paragraph(("The quick brown fox", None))
        mgr = IdManager()
        handle_modify(p, "The slow brown fox", id_manager=mgr, config=_default_config())

        del_els = xpath(p, "w:del")
        del_texts = []
        for d in del_els:
            for dt in d.iter(qn("w", "delText")):
                if dt.text:
                    del_texts.append(dt.text)
        combined = " ".join(del_texts)
        assert "quick" in combined

    def test_ins_contains_new_text(self):
        p = _make_paragraph(("The quick brown fox", None))
        mgr = IdManager()
        handle_modify(p, "The slow brown fox", id_manager=mgr, config=_default_config())

        ins_els = xpath(p, "w:ins")
        ins_texts = []
        for i in ins_els:
            for t in i.iter(qn("w", "t")):
                if t.text:
                    ins_texts.append(t.text)
        combined = " ".join(ins_texts)
        assert "slow" in combined

    def test_equal_text_preserved(self):
        p = _make_paragraph(("The quick brown fox", None))
        mgr = IdManager()
        handle_modify(p, "The slow brown fox", id_manager=mgr, config=_default_config())

        # Equal text should be in plain <w:r> elements (not inside del/ins)
        # Get all direct <w:r> children (not inside del/ins)
        direct_runs = xpath(p, "w:r")
        direct_texts = []
        for r in direct_runs:
            for t in r.iter(qn("w", "t")):
                if t.text:
                    direct_texts.append(t.text)
        combined = " ".join(direct_texts)
        assert "The" in combined
        assert "brown" in combined or "fox" in combined

    def test_preserves_ppr(self):
        """Paragraph properties should survive the rebuild."""
        p = _make_paragraph(("Hello world", None))
        # Add a <w:pPr> with a style
        ppr = etree.SubElement(p, qn("w", "pPr"))
        p.insert(0, ppr)
        pstyle = etree.SubElement(ppr, qn("w", "pStyle"))
        pstyle.set(qn("w", "val"), "Heading1")

        mgr = IdManager()
        handle_modify(p, "Goodbye world", id_manager=mgr, config=_default_config())

        # pPr should still exist with style
        found_ppr = xpath(p, "w:pPr")
        assert len(found_ppr) == 1
        found_style = xpath(found_ppr[0], "w:pStyle")
        assert len(found_style) == 1
        assert found_style[0].get(qn("w", "val")) == "Heading1"

    def test_word_insertion_only(self):
        p = _make_paragraph(("The Company shall provide notice", None))
        mgr = IdManager()
        handle_modify(
            p,
            "The Company shall provide written notice",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should have insertion but no deletion
        ins_els = xpath(p, "w:ins")
        assert len(ins_els) >= 1

        ins_texts = []
        for i in ins_els:
            for t in i.iter(qn("w", "t")):
                if t.text:
                    ins_texts.append(t.text)
        assert "written" in " ".join(ins_texts)

    def test_word_deletion_only(self):
        p = _make_paragraph(("The Company shall promptly provide notice", None))
        mgr = IdManager()
        handle_modify(
            p,
            "The Company shall provide notice",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should have deletion but no insertion
        del_els = xpath(p, "w:del")
        assert len(del_els) >= 1

        del_texts = []
        for d in del_els:
            for dt in d.iter(qn("w", "delText")):
                if dt.text:
                    del_texts.append(dt.text)
        assert "promptly" in " ".join(del_texts)

    def test_tracked_change_attributes(self):
        config = RedlineConfig(author="Review Bot")
        p = _make_paragraph(("old text", None))
        mgr = IdManager(start_after=0)
        handle_modify(p, "new text", id_manager=mgr, config=config)

        del_els = xpath(p, "w:del")
        if del_els:
            assert del_els[0].get(qn("w", "author")) == "Review Bot"

        ins_els = xpath(p, "w:ins")
        if ins_els:
            assert ins_els[0].get(qn("w", "author")) == "Review Bot"


# ===================================================================
# Integration: handlers with real fixture files
# ===================================================================


class TestHandlersWithFixtures:
    """Test handlers against real .docx fixtures."""

    def test_delete_on_real_doc(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[2]  # Second paragraph

        mgr = IdManager(start_after=doc.max_annotation_id())
        del_id = handle_delete(para, id_manager=mgr, config=_default_config())

        assert del_id > 0
        # The document should still save without error
        data = doc.to_bytes()
        assert len(data) > 0

    def test_append_on_real_doc(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[3]  # Third paragraph
        original_count = len(doc.paragraphs)

        mgr = IdManager(start_after=doc.max_annotation_id())
        _new_p, ins_id = handle_append_after(
            para,
            "This is a new paragraph.",
            id_manager=mgr,
            config=_default_config(),
        )

        assert ins_id > 0
        assert len(doc.paragraphs) == original_count + 1
        data = doc.to_bytes()
        assert len(data) > 0

    def test_modify_on_real_doc(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[1]  # First paragraph

        old_text = paragraph_to_pseudo_markdown(para)

        # Change something
        new_text = (
            old_text.replace("paragraph", "section")
            if "paragraph" in old_text.lower()
            else old_text + " (amended)"
        )

        mgr = IdManager(start_after=doc.max_annotation_id())
        ids = handle_modify(para, new_text, id_manager=mgr, config=_default_config())

        # Should have produced some tracked changes (unless text was identical)
        if old_text != new_text:
            assert len(ids) > 0

        data = doc.to_bytes()
        assert len(data) > 0


# ===================================================================
# Modify handler: escaped character regression tests (Bug 1 fix)
# ===================================================================


class TestModifyWithEscapedCharacters:
    """Regression tests for the garbled output bug.

    When a paragraph contains literal underscores (common in legal docs
    for fill-in-the-blank fields), the pseudo-Markdown escapes them as
    ``\\_``.  The old modify handler diffed pseudo-Markdown against raw
    run text, causing alignment failures that produced garbled single-
    character runs.

    After the fix, the handler diffs raw text vs raw text.  The new_text
    param may still be pseudo-Markdown (with ``\\_`` escapes), which
    ``pseudo_markdown_to_raw()`` converts back before diffing.
    """

    def test_modify_paragraph_with_underscores(self):
        """Modify a paragraph containing literal underscores without garbling."""
        p = _make_paragraph(("Name: ________", None))
        mgr = IdManager()
        ids = handle_modify(
            p,
            "Full Name: \\_\\_\\_\\_\\_\\_\\_\\_",
            id_manager=mgr,
            config=_default_config(),
        )

        assert len(ids) >= 1  # At least one tracked change

        # Verify the insertion contains "Full"
        ins_texts = []
        for i in xpath(p, "w:ins"):
            for t in i.iter(qn("w", "t")):
                if t.text:
                    ins_texts.append(t.text)
        assert "Full" in "".join(ins_texts)

        # Verify the underscores survive as EQUAL text (not garbled)
        equal_texts = []
        for r in xpath(p, "w:r"):
            for t in r.iter(qn("w", "t")):
                if t.text:
                    equal_texts.append(t.text)
        combined_equal = "".join(equal_texts)
        assert "____" in combined_equal

    def test_modify_preserves_long_underscore_fill(self):
        """A fill-in-the-blank like ____________ should not be garbled."""
        underscores = "_" * 20
        p = _make_paragraph((f"Party: {underscores} agrees", None))
        mgr = IdManager()

        # New text keeps the underscores but changes "Party" to "Seller"
        escaped_underscores = "\\_" * 20
        new_text = f"Seller: {escaped_underscores} agrees"

        ids = handle_modify(p, new_text, id_manager=mgr, config=_default_config())
        assert len(ids) >= 2

        # Collect all text content from the modified paragraph
        all_text = []
        for el in p.iter():
            tag = etree.QName(el.tag).localname if isinstance(el.tag, str) else ""
            if tag in ("t", "delText") and el.text:
                all_text.append(el.text)
        combined = "".join(all_text)

        # The underscores should appear as a contiguous block, not garbled
        assert underscores in combined

    def test_modify_no_change_with_underscores(self):
        """If new_text matches old text (after stripping), no changes should occur."""
        p = _make_paragraph(("Name: ________", None))
        mgr = IdManager()
        # Pass the exact escaped equivalent — should produce no diff
        ids = handle_modify(
            p,
            "Name: \\_\\_\\_\\_\\_\\_\\_\\_",
            id_manager=mgr,
            config=_default_config(),
        )
        assert ids == []

    def test_modify_with_asterisks_in_text(self):
        """Literal asterisks in original text should not cause garbling."""
        p = _make_paragraph(("Clause 5 * important *", None))
        mgr = IdManager()
        ids = handle_modify(
            p,
            "Clause 5 \\* critical \\*",
            id_manager=mgr,
            config=_default_config(),
        )

        assert len(ids) >= 2

        ins_texts = []
        for i in xpath(p, "w:ins"):
            for t in i.iter(qn("w", "t")):
                if t.text:
                    ins_texts.append(t.text)
        assert "critical" in "".join(ins_texts)

    def test_modify_round_trip_via_markup(self):
        """After modification, markup=True extraction should show tracked changes."""
        p = _make_paragraph(("The Company shall provide notice", None))
        mgr = IdManager()
        handle_modify(
            p,
            "The Company shall provide written notice",
            id_manager=mgr,
            config=_default_config(),
        )

        # Extract with markup=True
        md = paragraph_to_pseudo_markdown(p, markup=True)
        assert "++written++" in md or "++written " in md or "++ written++" in md
        # The unchanged parts should still be there
        assert "The" in md
        assert "Company" in md
        assert "notice" in md


# ===================================================================
# Modify handler: tab whitespace regression tests (Bug 4 fix)
# ===================================================================


class TestModifyWithTabs:
    """Regression tests for the tab-alignment garbling bug.

    Real legal documents use <w:tab/> elements for clause numbering
    (e.g. ``\\t\\t1.\\tThe term...``).  The tokenizer normalises these to
    spaces, but ``map_diff_to_runs()`` must align diff text (with spaces)
    against original run text (with tabs) without garbling.
    """

    @staticmethod
    def _make_tabbed_paragraph() -> etree._Element:
        """Build a paragraph with ``<w:tab/>`` elements like a real NDA clause.

        Structure: <w:tab/> | <w:tab/>1. | <w:tab/>The term means all info.
        Raw text: ``\\t\\t1.\\tThe term means all info.``
        """
        p = etree.Element(qn("w", "p"))

        r0 = etree.SubElement(p, qn("w", "r"))
        etree.SubElement(r0, qn("w", "tab"))

        r1 = etree.SubElement(p, qn("w", "r"))
        etree.SubElement(r1, qn("w", "tab"))
        t1 = etree.SubElement(r1, qn("w", "t"))
        t1.text = "1."

        r2 = etree.SubElement(p, qn("w", "r"))
        etree.SubElement(r2, qn("w", "tab"))
        t2 = etree.SubElement(r2, qn("w", "t"))
        t2.text = "The term means all info."

        return p

    def test_modify_tabbed_paragraph_no_garbling(self):
        """Modifying a tabbed paragraph should not produce garbled output."""
        p = self._make_tabbed_paragraph()
        mgr = IdManager()
        ids = handle_modify(
            p,
            "1. The term means all relevant info.",
            id_manager=mgr,
            config=_default_config(),
        )

        assert len(ids) >= 1

        # Collect all text from the paragraph
        all_text = []
        for el in p.iter():
            tag = etree.QName(el.tag).localname if isinstance(el.tag, str) else ""
            if tag in ("t", "delText") and el.text:
                all_text.append(el.text)
        combined = "".join(all_text)

        # Should contain the unchanged parts and the insertion
        assert "1." in combined
        assert "The term means all" in combined
        assert "relevant" in combined
        assert "info." in combined

        # Should NOT have garbled single-character words
        words = combined.split()
        single_chars = [w for w in words if len(w) == 1 and w not in ("1",)]
        assert len(single_chars) <= 2, f"Too many single-char words: {single_chars}"

    def test_modify_tabbed_paragraph_preserves_tabs(self):
        """Tabs in the original paragraph should be preserved after modification."""
        p = self._make_tabbed_paragraph()
        mgr = IdManager()
        handle_modify(
            p,
            "1. The term means all relevant info.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Extract with markup to see the result
        md = paragraph_to_pseudo_markdown(p, markup=True)
        assert "++relevant++" in md or "++relevant " in md or "++ relevant++" in md

    def test_no_change_tabbed_paragraph(self):
        """If new_text matches the tabbed paragraph content, no changes occur."""
        p = self._make_tabbed_paragraph()
        mgr = IdManager()
        # The raw text is "\t\t1.\tThe term means all info."
        # After pseudo_markdown_to_raw, "1. The term means all info." becomes
        # "1. The term means all info." which tokenizes the same way as the
        # original (tabs are stripped by tokenizer).
        ids = handle_modify(
            p,
            "1. The term means all info.",
            id_manager=mgr,
            config=_default_config(),
        )
        assert ids == []
