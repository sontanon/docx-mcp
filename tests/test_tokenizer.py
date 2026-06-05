"""Tests for the tokenizer module."""

from docx_mcp.tokenizer import detokenize, tokenize


class TestTokenize:
    def test_simple_sentence(self):
        assert tokenize("The quick brown fox") == ["The", "quick", "brown", "fox"]

    def test_empty_string(self):
        assert tokenize("") == []

    def test_whitespace_only(self):
        assert tokenize("   \t  \n  ") == []

    def test_single_word(self):
        assert tokenize("hello") == ["hello"]

    def test_multiple_spaces_between_words(self):
        assert tokenize("hello   world") == ["hello", "world"]

    def test_leading_trailing_whitespace(self):
        assert tokenize("  hello world  ") == ["hello", "world"]

    def test_formatting_markers_attached(self):
        """Markdown markers stay attached to their word token."""
        assert tokenize("**bold** _italic_ __underline__") == [
            "**bold**",
            "_italic_",
            "__underline__",
        ]

    def test_mixed_formatting(self):
        assert tokenize("The **quick** _brown_ fox") == [
            "The",
            "**quick**",
            "_brown_",
            "fox",
        ]

    def test_smart_quotes(self):
        tokens = tokenize("\u201cHello,\u201d she said.")
        assert tokens == ["\u201cHello,\u201d", "she", "said."]

    def test_non_breaking_space_splits_words(self):
        """Non-breaking space (\\u00a0) is \\s, so \\S+ splits on it."""
        tokens = tokenize("Section\u00a05.1")
        assert tokens == ["Section", "5.1"]

    def test_em_dash(self):
        tokens = tokenize("party\u2014the Company")
        # Em-dash is non-whitespace, so it stays with its word
        assert tokens == ["party\u2014the", "Company"]

    def test_section_symbol(self):
        tokens = tokenize("\u00a7 5.1(a)")
        assert tokens == ["\u00a7", "5.1(a)"]

    def test_tabs_and_newlines(self):
        assert tokenize("hello\tworld\nfoo") == ["hello", "world", "foo"]


class TestDetokenize:
    def test_simple(self):
        assert detokenize(["hello", "world"]) == "hello world"

    def test_empty_list(self):
        assert detokenize([]) == ""

    def test_single_token(self):
        assert detokenize(["hello"]) == "hello"

    def test_round_trip_simple(self):
        text = "The quick brown fox"
        assert detokenize(tokenize(text)) == text

    def test_round_trip_collapses_multiple_spaces(self):
        """Tokenize+detokenize normalises multiple spaces to single."""
        text = "hello   world"
        assert detokenize(tokenize(text)) == "hello world"

    def test_round_trip_with_formatting(self):
        text = "The **quick** _brown_ fox"
        assert detokenize(tokenize(text)) == text
