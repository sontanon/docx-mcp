"""Tests for the differ module."""

from __future__ import annotations

from docx_mcp.differ import DmpWordDiffer
from docx_mcp.models import DiffChunk, DiffOp


class TestDmpWordDiffer:
    """Tests for the DMP-based word-level differ."""

    def setup_method(self):
        self.differ = DmpWordDiffer()

    # -- Identical / trivial cases --

    def test_identical_text(self):
        chunks = self.differ.diff("hello world", "hello world")
        assert chunks == [DiffChunk(op=DiffOp.EQUAL, text="hello world")]

    def test_both_empty(self):
        assert self.differ.diff("", "") == []

    def test_old_empty_insert_all(self):
        chunks = self.differ.diff("", "hello world")
        assert chunks == [DiffChunk(op=DiffOp.INSERT, text="hello world")]

    def test_new_empty_delete_all(self):
        chunks = self.differ.diff("hello world", "")
        assert chunks == [DiffChunk(op=DiffOp.DELETE, text="hello world")]

    # -- Word replacement --

    def test_single_word_replacement(self):
        chunks = self.differ.diff(
            "The quick brown fox",
            "The slow brown fox",
        )
        # Should have: equal "The", delete "quick", insert "slow", equal "brown fox"
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.EQUAL, "The") in ops
        assert (DiffOp.DELETE, "quick") in ops
        assert (DiffOp.INSERT, "slow") in ops
        assert (DiffOp.EQUAL, "brown fox") in ops

    def test_multi_word_replacement(self):
        chunks = self.differ.diff(
            "The Company shall provide notice",
            "The Company must give written notice",
        )
        # "The Company" should be equal
        assert any(c.op == DiffOp.EQUAL and "The Company" in c.text for c in chunks)
        # "shall provide" should be deleted
        assert any(c.op == DiffOp.DELETE for c in chunks)
        # "must give written" should be inserted
        assert any(c.op == DiffOp.INSERT for c in chunks)
        # "notice" should be equal
        assert any(c.op == DiffOp.EQUAL and "notice" in c.text for c in chunks)

    # -- Insertion only --

    def test_word_insertion(self):
        chunks = self.differ.diff(
            "The Company shall provide notice",
            "The Company shall provide written notice",
        )
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.INSERT, "written") in ops
        # No deletions
        assert all(c.op != DiffOp.DELETE for c in chunks)

    # -- Deletion only --

    def test_word_deletion(self):
        chunks = self.differ.diff(
            "The Company shall promptly provide notice",
            "The Company shall provide notice",
        )
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.DELETE, "promptly") in ops
        # No insertions
        assert all(c.op != DiffOp.INSERT for c in chunks)

    # -- Formatting markers as tokens --

    def test_formatting_markers_in_diff(self):
        """Bold markers are part of the token and diff correctly."""
        chunks = self.differ.diff(
            "The **Company** shall provide",
            "The **Company** must provide",
        )
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.DELETE, "shall") in ops
        assert (DiffOp.INSERT, "must") in ops
        # **Company** should be in an equal chunk
        assert any(c.op == DiffOp.EQUAL and "**Company**" in c.text for c in chunks)

    def test_formatting_added(self):
        """Adding bold markers to a word shows as delete old + insert new."""
        chunks = self.differ.diff("The Company shall", "The **Company** shall")
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.DELETE, "Company") in ops
        assert (DiffOp.INSERT, "**Company**") in ops

    # -- Smart quotes and unicode --

    def test_smart_quotes_preserved(self):
        old = "the \u201cCompany\u201d agrees"
        new = "the \u201cCompany\u201d hereby agrees"
        chunks = self.differ.diff(old, new)
        ops = [(c.op, c.text) for c in chunks]
        assert any(c.op == DiffOp.EQUAL and "\u201cCompany\u201d" in c.text for c in chunks)
        assert (DiffOp.INSERT, "hereby") in ops

    def test_em_dash_in_token(self):
        old = "party\u2014the Company"
        new = "party\u2014the Organization"
        chunks = self.differ.diff(old, new)
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.EQUAL, "party\u2014the") in ops
        assert (DiffOp.DELETE, "Company") in ops
        assert (DiffOp.INSERT, "Organization") in ops

    def test_section_symbol(self):
        old = "\u00a7 5.1(a) of the Agreement"
        new = "\u00a7 5.2(b) of the Agreement"
        chunks = self.differ.diff(old, new)
        ops = [(c.op, c.text) for c in chunks]
        assert (DiffOp.EQUAL, "\u00a7") in ops
        assert (DiffOp.DELETE, "5.1(a)") in ops
        assert (DiffOp.INSERT, "5.2(b)") in ops

    # -- Reconstruction --

    def test_reconstruct_new_from_equal_and_insert(self):
        """Concatenating EQUAL + INSERT text should reproduce the new text."""
        old = "The Company shall provide notice to the other party"
        new = "The Company must give written notice to each party"
        chunks = self.differ.diff(old, new)
        reconstructed = " ".join(c.text for c in chunks if c.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert reconstructed == new

    def test_reconstruct_old_from_equal_and_delete(self):
        """Concatenating EQUAL + DELETE text should reproduce the old text."""
        old = "The Company shall provide notice to the other party"
        new = "The Company must give written notice to each party"
        chunks = self.differ.diff(old, new)
        reconstructed = " ".join(c.text for c in chunks if c.op in (DiffOp.EQUAL, DiffOp.DELETE))
        assert reconstructed == old

    # -- Edge cases --

    def test_single_word_to_single_word(self):
        chunks = self.differ.diff("hello", "goodbye")
        assert chunks == [
            DiffChunk(op=DiffOp.DELETE, text="hello"),
            DiffChunk(op=DiffOp.INSERT, text="goodbye"),
        ]

    def test_completely_different_texts(self):
        chunks = self.differ.diff("aaa bbb ccc", "xxx yyy zzz")
        assert chunks == [
            DiffChunk(op=DiffOp.DELETE, text="aaa bbb ccc"),
            DiffChunk(op=DiffOp.INSERT, text="xxx yyy zzz"),
        ]

    def test_no_spurious_empty_chunks(self):
        chunks = self.differ.diff("The quick brown fox", "The slow brown fox")
        for chunk in chunks:
            assert chunk.text != ""

    def test_adjacent_chunks_have_different_ops(self):
        """After merging, no two adjacent chunks should have the same op."""
        chunks = self.differ.diff(
            "The Company shall provide timely notice to the other party",
            "The Company must give written notice to each party promptly",
        )
        for i in range(len(chunks) - 1):
            assert chunks[i].op != chunks[i + 1].op, (
                f"Adjacent chunks {i} and {i + 1} both have op={chunks[i].op}"
            )

    def test_legal_clause_modification(self):
        """Realistic legal text modification."""
        old = (
            "The Receiving Party shall hold and maintain the Confidential "
            "Information in strict confidence for the sole and exclusive benefit "
            "of the Disclosing Party."
        )
        new = (
            "The Receiving Party shall hold and maintain the Confidential "
            "Information in strict confidence for the benefit "
            "of the Disclosing Party."
        )
        chunks = self.differ.diff(old, new)
        # "sole and exclusive" should be deleted
        deleted = " ".join(c.text for c in chunks if c.op == DiffOp.DELETE)
        assert "sole and exclusive" in deleted
        # Most text should be equal
        equal_words = sum(len(c.text.split()) for c in chunks if c.op == DiffOp.EQUAL)
        total_old_words = len(old.split())
        assert equal_words > total_old_words * 0.7  # At least 70% equal
