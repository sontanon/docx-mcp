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
)
from docx_mcp.document import DocxDocument


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
