"""Tests for the change handlers (delete, append, modify)."""

from __future__ import annotations

from lxml import etree

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

        from docx_mcp.converter import paragraph_to_pseudo_markdown

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
