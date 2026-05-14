"""Tests for the docx-to-pseudo-Markdown converter."""

from pathlib import Path

from lxml import etree

from docx_mcp.converter import (
    _escape_markdown,
    _unescape_markdown,
    _wrap_formatting,
    document_to_fragments,
    fragments_to_tagged_text,
    paragraph_to_pseudo_markdown,
    pseudo_markdown_to_raw,
)
from docx_mcp.document import DocxDocument
from docx_mcp.namespaces import XML


class TestEscapeMarkdown:
    def test_escapes_asterisks(self):
        assert _escape_markdown("2 * 3 = 6") == "2 \\* 3 = 6"

    def test_escapes_underscores(self):
        assert _escape_markdown("var_name") == "var\\_name"

    def test_escapes_backslashes(self):
        assert _escape_markdown("path\\to") == "path\\\\to"

    def test_preserves_smart_quotes(self):
        text = "\u201cQuoted\u201d text with \u2018single\u2019"
        assert _escape_markdown(text) == text

    def test_preserves_em_dash(self):
        assert _escape_markdown("clause\u2014provided") == "clause\u2014provided"

    def test_preserves_section_symbol(self):
        assert _escape_markdown("\u00a7 5.1") == "\u00a7 5.1"

    def test_preserves_non_breaking_space(self):
        text = "Section\u00a05.1"
        assert _escape_markdown(text) == text

    def test_round_trip(self):
        original = "bold * _underline_ and \\backslash"
        assert _unescape_markdown(_escape_markdown(original)) == original


class TestWrapFormatting:
    def test_plain_text(self):
        assert _wrap_formatting("hello", bold=False, italic=False, underline=False) == "hello"

    def test_bold(self):
        assert _wrap_formatting("hello", bold=True, italic=False, underline=False) == "**hello**"

    def test_italic(self):
        assert _wrap_formatting("hello", bold=False, italic=True, underline=False) == "_hello_"

    def test_underline(self):
        assert _wrap_formatting("hello", bold=False, italic=False, underline=True) == "__hello__"

    def test_bold_italic(self):
        assert _wrap_formatting("hello", bold=True, italic=True, underline=False) == "**_hello_**"

    def test_bold_underline(self):
        assert _wrap_formatting("hello", bold=True, italic=False, underline=True) == "**__hello__**"

    def test_all_formatting(self):
        assert (
            _wrap_formatting("hello", bold=True, italic=True, underline=True) == "**___hello___**"
        )

    def test_empty_text(self):
        assert _wrap_formatting("", bold=True, italic=True, underline=True) == ""


class TestParagraphToPseudoMarkdown:
    """Test conversion of actual parsed XML paragraphs."""

    def test_plain_text_paragraph(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        # First paragraph should be plain text
        for para in doc.paragraphs:
            md = paragraph_to_pseudo_markdown(para)
            if md:  # skip empty paragraphs
                # Should not contain formatting markers (no **, _, __)
                # except escaped ones
                assert "**" not in md or "\\*" in md
                break

    def test_formatted_runs_bold(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        # First paragraph has "Seller" in bold
        md = paragraph_to_pseudo_markdown(doc.paragraphs[0])
        assert "**Seller**" in md

    def test_formatted_runs_italic(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        # First paragraph has "goods" in italic
        md = paragraph_to_pseudo_markdown(doc.paragraphs[0])
        assert "_goods_" in md

    def test_formatted_runs_underline(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        # First paragraph has "thirty days" underlined
        md = paragraph_to_pseudo_markdown(doc.paragraphs[0])
        assert "__thirty days__" in md

    def test_formatted_runs_bold_italic(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        # Second paragraph has "United States Dollars" in bold+italic
        md = paragraph_to_pseudo_markdown(doc.paragraphs[1])
        assert "**_United States Dollars_**" in md

    def test_plain_paragraph_no_markers(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        # Third paragraph is plain text
        md = paragraph_to_pseudo_markdown(doc.paragraphs[2])
        assert "**" not in md
        assert md.startswith("This paragraph has no special formatting")

    def test_special_chars_smart_quotes(self, special_chars_path: Path):
        doc = DocxDocument(path=special_chars_path)
        md = paragraph_to_pseudo_markdown(doc.paragraphs[0])
        # Smart quotes should be preserved
        assert "\u201c" in md  # left double quote
        assert "\u201d" in md  # right double quote
        assert "\u2019" in md  # right single quote (apostrophe)

    def test_special_chars_em_dash(self, special_chars_path: Path):
        doc = DocxDocument(path=special_chars_path)
        md = paragraph_to_pseudo_markdown(doc.paragraphs[1])
        assert "\u2014" in md  # em dash

    def test_special_chars_section_symbol(self, special_chars_path: Path):
        doc = DocxDocument(path=special_chars_path)
        md = paragraph_to_pseudo_markdown(doc.paragraphs[2])
        assert "\u00a7" in md  # section symbol

    def test_special_chars_non_breaking_space(self, special_chars_path: Path):
        doc = DocxDocument(path=special_chars_path)
        md = paragraph_to_pseudo_markdown(doc.paragraphs[2])
        assert "\u00a0" in md  # non-breaking space preserved

    def test_multiple_spaces_collapsed(self):
        """Verify multiple regular spaces are collapsed."""
        # Build a paragraph XML with multiple spaces in a run
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = etree.Element(f"{{{nsmap['w']}}}p", nsmap=nsmap)
        r = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        t = etree.SubElement(r, f"{{{nsmap['w']}}}t")
        t.text = "hello    world"
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        md = paragraph_to_pseudo_markdown(p)
        assert md == "hello world"

    def test_empty_paragraph(self):
        """Empty paragraph produces empty string."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = etree.Element(f"{{{nsmap['w']}}}p", nsmap=nsmap)
        md = paragraph_to_pseudo_markdown(p)
        assert md == ""

    def test_whitespace_only_run_no_artifact(self):
        """Bold/italic whitespace-only runs must not produce formatting artifacts."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = etree.Element(f"{{{nsmap['w']}}}p", nsmap=nsmap)

        # Run 1: bold space
        r1 = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        rpr1 = etree.SubElement(r1, f"{{{nsmap['w']}}}rPr")
        etree.SubElement(rpr1, f"{{{nsmap['w']}}}b")
        t1 = etree.SubElement(r1, f"{{{nsmap['w']}}}t")
        t1.text = " "
        t1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        # Run 2: normal text
        r2 = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        t2 = etree.SubElement(r2, f"{{{nsmap['w']}}}t")
        t2.text = "hello"

        md = paragraph_to_pseudo_markdown(p)
        assert "****" not in md
        assert "** **" not in md
        assert md == "hello"

    def test_tab_only_run_no_artifact(self):
        """Bold tab-only runs must not produce formatting artifacts."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = etree.Element(f"{{{nsmap['w']}}}p", nsmap=nsmap)

        # Run 1: bold tab
        r1 = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        rpr1 = etree.SubElement(r1, f"{{{nsmap['w']}}}rPr")
        etree.SubElement(rpr1, f"{{{nsmap['w']}}}b")
        etree.SubElement(r1, f"{{{nsmap['w']}}}tab")

        # Run 2: normal text
        r2 = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        t2 = etree.SubElement(r2, f"{{{nsmap['w']}}}t")
        t2.text = "world"

        md = paragraph_to_pseudo_markdown(p)
        assert "**\t**" not in md
        assert md == "world"

    def test_multiple_empty_runs_no_artifact(self):
        """Sequence of empty bold runs must not produce '****'."""
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        p = etree.Element(f"{{{nsmap['w']}}}p", nsmap=nsmap)

        for _ in range(3):
            r = etree.SubElement(p, f"{{{nsmap['w']}}}r")
            rpr = etree.SubElement(r, f"{{{nsmap['w']}}}rPr")
            etree.SubElement(rpr, f"{{{nsmap['w']}}}b")
            t = etree.SubElement(r, f"{{{nsmap['w']}}}t")
            t.text = ""

        # Run with actual text
        r = etree.SubElement(p, f"{{{nsmap['w']}}}r")
        t = etree.SubElement(r, f"{{{nsmap['w']}}}t")
        t.text = "actual"

        md = paragraph_to_pseudo_markdown(p)
        assert "****" not in md
        assert md == "actual"


class TestDocumentToFragments:
    def test_simple_5para_fragments(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        fragments = document_to_fragments(doc.paragraphs)
        assert len(fragments) == len(doc.paragraphs)
        # IDs should be 1-based sequential
        ids = [fid for fid, _ in fragments]
        assert ids == list(range(1, len(fragments) + 1))

    def test_fragments_contain_text(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        fragments = document_to_fragments(doc.paragraphs)
        # At least some fragments should have non-empty text
        non_empty = [(fid, text) for fid, text in fragments if text]
        assert len(non_empty) >= 5


class TestFragmentsToTaggedText:
    def test_basic_format(self):
        fragments = [(1, "First."), (2, "Second."), (3, "Third.")]
        result = fragments_to_tagged_text(fragments)
        assert result == "<f=1>First.</f=1>\n<f=2>Second.</f=2>\n<f=3>Third.</f=3>"

    def test_empty_fragment(self):
        fragments = [(1, "Text"), (2, ""), (3, "More")]
        result = fragments_to_tagged_text(fragments)
        assert "<f=2></f=2>" in result

    def test_formatting_markers_in_text(self):
        fragments = [(1, "This is **bold** text.")]
        result = fragments_to_tagged_text(fragments)
        assert result == "<f=1>This is **bold** text.</f=1>"


# ===================================================================
# pseudo_markdown_to_raw tests
# ===================================================================


class TestPseudoMarkdownToRaw:
    def test_plain_text_unchanged(self):
        assert pseudo_markdown_to_raw("Hello world") == "Hello world"

    def test_strips_bold(self):
        assert pseudo_markdown_to_raw("**bold**") == "bold"

    def test_strips_italic(self):
        assert pseudo_markdown_to_raw("_italic_") == "italic"

    def test_strips_underline(self):
        assert pseudo_markdown_to_raw("__underline__") == "underline"

    def test_strips_nested_bold_underline(self):
        assert pseudo_markdown_to_raw("**__bold underline__**") == "bold underline"

    def test_strips_nested_bold_italic(self):
        assert pseudo_markdown_to_raw("**_bold italic_**") == "bold italic"

    def test_strips_all_three_nested(self):
        assert pseudo_markdown_to_raw("**___bold italic underline___**") == "bold italic underline"

    def test_unescapes_underscores(self):
        assert pseudo_markdown_to_raw("\\_name\\_") == "_name_"

    def test_unescapes_asterisks(self):
        assert pseudo_markdown_to_raw("\\*star\\*") == "*star*"

    def test_unescapes_backslashes(self):
        assert pseudo_markdown_to_raw("path\\\\to") == "path\\to"

    def test_mixed_formatting_and_escapes(self):
        text = "The **Seller** shall deliver the \\_goods\\_ within __thirty days__."
        expected = "The Seller shall deliver the _goods_ within thirty days."
        assert pseudo_markdown_to_raw(text) == expected

    def test_multiple_underscores_escaped(self):
        """The NDA bug scenario: literal underscores escaped in pseudo-Markdown."""
        # Four literal underscores become four escaped underscores
        text = "\\_\\_\\_\\_"
        assert pseudo_markdown_to_raw(text) == "____"

    def test_long_underscore_fill_escaped(self):
        """Simulate a fill-in-the-blank like ____________ in a legal doc."""
        raw_underscores = "_" * 12
        escaped = _escape_markdown(raw_underscores)
        assert pseudo_markdown_to_raw(escaped) == raw_underscores

    def test_round_trip_plain(self):
        """pseudo_markdown_to_raw(escape(text)) should return original text."""
        original = "Section 5.1: obligations of the _underscore_ party"
        escaped = _escape_markdown(original)
        assert pseudo_markdown_to_raw(escaped) == original

    def test_empty_string(self):
        assert pseudo_markdown_to_raw("") == ""

    def test_preserves_unicode(self):
        text = "\u201cQuoted\u201d \u2014 \u00a7 5.1"
        assert pseudo_markdown_to_raw(text) == text


# ===================================================================
# paragraph_to_pseudo_markdown with markup=True
# ===================================================================


def _make_tracked_change_paragraph() -> etree._Element:
    """Build a paragraph with <w:ins> and <w:del> tracked changes.

    Structure: "The " + DEL("quick ") + INS("slow ") + "brown fox"
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    nsmap = {"w": W}

    p = etree.Element(f"{{{W}}}p", nsmap=nsmap)

    # Plain run: "The "
    r1 = etree.SubElement(p, f"{{{W}}}r")
    t1 = etree.SubElement(r1, f"{{{W}}}t")
    t1.text = "The "
    t1.set(f"{{{XML}}}space", "preserve")

    # Deleted region: "quick "
    del_el = etree.SubElement(p, f"{{{W}}}del")
    del_el.set(f"{{{W}}}id", "1")
    del_el.set(f"{{{W}}}author", "Test")
    del_el.set(f"{{{W}}}date", "2026-01-01T00:00:00Z")
    dr = etree.SubElement(del_el, f"{{{W}}}r")
    dt = etree.SubElement(dr, f"{{{W}}}delText")
    dt.text = "quick "
    dt.set(f"{{{XML}}}space", "preserve")

    # Inserted region: "slow "
    ins_el = etree.SubElement(p, f"{{{W}}}ins")
    ins_el.set(f"{{{W}}}id", "2")
    ins_el.set(f"{{{W}}}author", "Test")
    ins_el.set(f"{{{W}}}date", "2026-01-01T00:00:00Z")
    ir = etree.SubElement(ins_el, f"{{{W}}}r")
    it = etree.SubElement(ir, f"{{{W}}}t")
    it.text = "slow "
    it.set(f"{{{XML}}}space", "preserve")

    # Plain run: "brown fox"
    r2 = etree.SubElement(p, f"{{{W}}}r")
    t2 = etree.SubElement(r2, f"{{{W}}}t")
    t2.text = "brown fox"

    return p


class TestParagraphToPseudoMarkdownMarkup:
    """Tests for paragraph_to_pseudo_markdown with markup=True."""

    def test_markup_false_ignores_tracked_changes(self):
        p = _make_tracked_change_paragraph()
        md = paragraph_to_pseudo_markdown(p, markup=False)
        # Default mode should only see direct <w:r> children
        assert "quick" not in md
        assert "slow" not in md
        assert "The" in md
        assert "brown fox" in md

    def test_markup_true_shows_tracked_changes(self):
        p = _make_tracked_change_paragraph()
        md = paragraph_to_pseudo_markdown(p, markup=True)
        assert "~~quick~~" in md or "~~quick ~~" in md
        assert "++slow++" in md or "++slow ++" in md
        assert "The" in md
        assert "brown fox" in md

    def test_markup_true_insertion_markers(self):
        p = _make_tracked_change_paragraph()
        md = paragraph_to_pseudo_markdown(p, markup=True)
        # The inserted text should be wrapped in ++…++
        assert "++" in md
        # Extract the content between ++ markers
        import re

        ins_match = re.search(r"\+\+(.+?)\+\+", md)
        assert ins_match is not None
        assert "slow" in ins_match.group(1)

    def test_markup_true_deletion_markers(self):
        p = _make_tracked_change_paragraph()
        md = paragraph_to_pseudo_markdown(p, markup=True)
        # The deleted text should be wrapped in ~~…~~
        assert "~~" in md
        import re

        del_match = re.search(r"~~(.+?)~~", md)
        assert del_match is not None
        assert "quick" in del_match.group(1)

    def test_markup_true_with_bold_in_insertion(self):
        """Inserted text with formatting should have both markers."""
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": W}
        p = etree.Element(f"{{{W}}}p", nsmap=nsmap)

        ins_el = etree.SubElement(p, f"{{{W}}}ins")
        ins_el.set(f"{{{W}}}id", "1")
        ins_el.set(f"{{{W}}}author", "Test")
        ins_el.set(f"{{{W}}}date", "2026-01-01T00:00:00Z")
        r = etree.SubElement(ins_el, f"{{{W}}}r")
        rpr = etree.SubElement(r, f"{{{W}}}rPr")
        etree.SubElement(rpr, f"{{{W}}}b")
        t = etree.SubElement(r, f"{{{W}}}t")
        t.text = "Important"

        md = paragraph_to_pseudo_markdown(p, markup=True)
        assert "++**Important**++" in md


# ===================================================================
# document_to_fragments with markup
# ===================================================================


class TestDocumentToFragmentsMarkup:
    def test_markup_false_is_default(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        frags_default = document_to_fragments(doc.paragraphs)
        frags_explicit = document_to_fragments(doc.paragraphs, markup=False)
        assert frags_default == frags_explicit

    def test_markup_true_passes_through(self, simple_5para_path: Path):
        """On an unmodified doc, markup=True should give the same result."""
        doc = DocxDocument(path=simple_5para_path)
        frags_plain = document_to_fragments(doc.paragraphs, markup=False)
        frags_markup = document_to_fragments(doc.paragraphs, markup=True)
        # On a clean doc with no tracked changes, results should be identical
        assert frags_plain == frags_markup
