"""Tests for the IdManager and run_ops modules."""

from __future__ import annotations

import copy

import pytest
from lxml import etree

from docx_mcp.id_manager import IdManager
from docx_mcp.models import DiffChunk, DiffOp
from docx_mcp.namespaces import XML, qn
from docx_mcp.run_ops import (
    TaggedSegment,
    _inject_inter_chunk_spaces,
    _merge_tagged_segments,
    _rpr_equal,
    build_run_element,
    build_tracked_change_element,
    clone_rpr,
    extract_runs,
    get_paragraph_text,
    map_diff_to_runs,
    split_run_at,
)

# ===================================================================
# IdManager tests
# ===================================================================


class TestIdManager:
    def test_starts_at_one_by_default(self):
        mgr = IdManager()
        assert mgr.next_id() == 1

    def test_starts_after_given_id(self):
        mgr = IdManager(start_after=102)
        assert mgr.next_id() == 103

    def test_monotonically_increasing(self):
        mgr = IdManager(start_after=0)
        ids = [mgr.next_id() for _ in range(5)]
        assert ids == [1, 2, 3, 4, 5]

    def test_peek_does_not_consume(self):
        mgr = IdManager(start_after=10)
        assert mgr.peek == 11
        assert mgr.peek == 11  # Still 11
        assert mgr.next_id() == 11
        assert mgr.peek == 12


# ===================================================================
# Helpers for building test XML
# ===================================================================


def _make_paragraph(*run_specs: tuple[str, dict[str, bool] | None]) -> etree._Element:
    """Build a ``<w:p>`` element with runs.

    Each *run_spec* is ``(text, formatting_dict)`` where formatting_dict
    may contain ``bold``, ``italic``, ``underline`` keys.
    """
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


# ===================================================================
# extract_runs tests
# ===================================================================


class TestExtractRuns:
    def test_simple_paragraph(self):
        p = _make_paragraph(("Hello world", None))
        runs = extract_runs(p)
        assert len(runs) == 1
        assert runs[0].text == "Hello world"
        assert runs[0].rpr is None

    def test_multiple_runs(self):
        p = _make_paragraph(
            ("The ", None),
            ("Company", {"bold": True}),
            (" shall", None),
        )
        runs = extract_runs(p)
        assert len(runs) == 3
        assert runs[0].text == "The "
        assert runs[1].text == "Company"
        assert runs[2].text == " shall"

    def test_bold_detection(self):
        p = _make_paragraph(("bold text", {"bold": True}))
        runs = extract_runs(p)
        assert runs[0].rpr is not None

    def test_empty_run(self):
        """Runs with no text content are included with empty text."""
        p = etree.Element(qn("w", "p"))
        r = etree.SubElement(p, qn("w", "r"))
        # No <w:t> child — just rPr
        etree.SubElement(r, qn("w", "rPr"))
        runs = extract_runs(p)
        assert len(runs) == 1
        assert runs[0].text == ""


class TestGetParagraphText:
    def test_concatenates_all_runs(self):
        p = _make_paragraph(
            ("The ", None),
            ("Company", {"bold": True}),
            (" shall provide notice.", None),
        )
        runs = extract_runs(p)
        assert get_paragraph_text(runs) == "The Company shall provide notice."


# ===================================================================
# clone_rpr tests
# ===================================================================


class TestCloneRpr:
    def test_none_returns_none(self):
        assert clone_rpr(None) is None

    def test_deep_copy(self):
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        cloned = clone_rpr(rpr)
        assert cloned is not None
        assert cloned is not rpr
        assert etree.tostring(cloned) == etree.tostring(rpr)

    def test_modification_independent(self):
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        cloned = clone_rpr(rpr)
        assert cloned is not None
        etree.SubElement(cloned, qn("w", "i"))
        # Original should not have italic
        assert len(rpr) == 1


# ===================================================================
# split_run_at tests
# ===================================================================


class TestSplitRunAt:
    def test_split_in_middle(self):
        p = _make_paragraph(("Hello world", {"bold": True}))
        run = extract_runs(p)[0]
        left, right = split_run_at(run, 5)
        assert left.text == "Hello"
        assert right.text == " world"

    def test_split_preserves_rpr(self):
        p = _make_paragraph(("Hello world", {"bold": True}))
        run = extract_runs(p)[0]
        left, right = split_run_at(run, 5)
        # Both halves should have independent rPr copies
        assert left.rpr is not None
        assert right.rpr is not None
        assert left.rpr is not right.rpr
        assert left.rpr is not run.rpr

    def test_split_at_offset_zero_raises(self):
        p = _make_paragraph(("Hello", None))
        run = extract_runs(p)[0]
        with pytest.raises(ValueError, match="out of range"):
            split_run_at(run, 0)

    def test_split_at_end_raises(self):
        p = _make_paragraph(("Hello", None))
        run = extract_runs(p)[0]
        with pytest.raises(ValueError, match="out of range"):
            split_run_at(run, 5)

    def test_split_single_char_boundary(self):
        p = _make_paragraph(("AB", None))
        run = extract_runs(p)[0]
        left, right = split_run_at(run, 1)
        assert left.text == "A"
        assert right.text == "B"


# ===================================================================
# map_diff_to_runs tests
# ===================================================================


class TestMapDiffToRuns:
    def test_all_equal(self):
        """When text is unchanged, all segments should be EQUAL."""
        p = _make_paragraph(("Hello world", None))
        runs = extract_runs(p)
        chunks = [DiffChunk(op=DiffOp.EQUAL, text="Hello world")]
        segments = map_diff_to_runs(chunks, runs)
        assert all(s.op == DiffOp.EQUAL for s in segments)
        combined = "".join(s.text for s in segments)
        assert combined == "Hello world"

    def test_simple_word_replacement(self):
        """Replace one word — should produce EQUAL + DELETE + INSERT + EQUAL."""
        p = _make_paragraph(("The quick brown fox", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The"),
            DiffChunk(op=DiffOp.DELETE, text="quick"),
            DiffChunk(op=DiffOp.INSERT, text="slow"),
            DiffChunk(op=DiffOp.EQUAL, text="brown fox"),
        ]
        segments = map_diff_to_runs(chunks, runs)

        # Check we have the right ops
        ops = [s.op for s in segments]
        assert DiffOp.EQUAL in ops
        assert DiffOp.DELETE in ops
        assert DiffOp.INSERT in ops

        # Reconstruct old text (EQUAL + DELETE)
        old_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.DELETE))
        # Should match original (modulo possible whitespace differences)
        assert "The" in old_text
        assert "quick" in old_text
        assert "brown fox" in old_text

    def test_insert_inherits_formatting(self):
        """Inserted text should inherit formatting from adjacent run."""
        p = _make_paragraph(("Hello world", {"bold": True}))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="Hello"),
            DiffChunk(op=DiffOp.INSERT, text="beautiful"),
            DiffChunk(op=DiffOp.EQUAL, text="world"),
        ]
        segments = map_diff_to_runs(chunks, runs)
        insert_segs = [s for s in segments if s.op == DiffOp.INSERT]
        assert len(insert_segs) == 1
        # Should have inherited bold formatting
        assert insert_segs[0].rpr is not None

    def test_delete_preserves_original_text(self):
        """Deleted segments should contain text from the original runs."""
        p = _make_paragraph(("The Company shall provide notice", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The Company"),
            DiffChunk(op=DiffOp.DELETE, text="shall provide"),
            DiffChunk(op=DiffOp.EQUAL, text="notice"),
        ]
        segments = map_diff_to_runs(chunks, runs)
        delete_segs = [s for s in segments if s.op == DiffOp.DELETE]
        delete_text = "".join(s.text for s in delete_segs)
        assert "shall" in delete_text
        assert "provide" in delete_text

    def test_multi_run_paragraph(self):
        """Diff across run boundaries should still produce correct segments."""
        p = _make_paragraph(
            ("The ", None),
            ("Company", {"bold": True}),
            (" shall provide notice.", None),
        )
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The Company shall"),
            DiffChunk(op=DiffOp.DELETE, text="provide"),
            DiffChunk(op=DiffOp.INSERT, text="give"),
            DiffChunk(op=DiffOp.EQUAL, text="notice."),
        ]
        segments = map_diff_to_runs(chunks, runs)

        # Verify all text is accounted for
        all_ops = {s.op for s in segments}
        assert DiffOp.EQUAL in all_ops
        assert DiffOp.DELETE in all_ops
        assert DiffOp.INSERT in all_ops

    def test_empty_diff_empty_runs(self):
        """Empty inputs should produce empty output."""
        segments = map_diff_to_runs([], [])
        assert segments == []

    def test_tabs_in_runs_aligned_with_space_in_diff(self):
        """Tabs from <w:tab/> should align with spaces in tokenised diff text.

        Regression test for the NDA garbling bug: the tokenizer normalises
        whitespace (tabs → spaces), but map_diff_to_runs must still align
        the diff text against original run text that contains \\t characters.
        """
        # Build paragraph with <w:tab/> elements like a real indented clause:
        # Run 0: <w:tab/>  → "\t"
        # Run 1: <w:tab/> + "1."  → "\t1."
        # Run 2: <w:tab/> + "The term means all info."  → "\tThe term means all info."
        p = etree.Element(qn("w", "p"))

        r0 = etree.SubElement(p, qn("w", "r"))
        rpr0 = etree.SubElement(r0, qn("w", "rPr"))
        etree.SubElement(rpr0, qn("w", "b"))
        etree.SubElement(r0, qn("w", "tab"))

        r1 = etree.SubElement(p, qn("w", "r"))
        rpr1 = etree.SubElement(r1, qn("w", "rPr"))
        etree.SubElement(rpr1, qn("w", "b"))
        etree.SubElement(r1, qn("w", "tab"))
        t1 = etree.SubElement(r1, qn("w", "t"))
        t1.text = "1."

        r2 = etree.SubElement(p, qn("w", "r"))
        etree.SubElement(r2, qn("w", "tab"))
        t2 = etree.SubElement(r2, qn("w", "t"))
        t2.text = "The term means all info."

        runs = extract_runs(p)
        raw = get_paragraph_text(runs)
        assert raw == "\t\t1.\tThe term means all info."

        # Diff: replace "all info" with "all relevant info"
        # After tokenize/detokenize, the text is "1. The term means all info."
        # (tabs become word separators, detokenize joins with single spaces)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="1. The term means all"),
            DiffChunk(op=DiffOp.INSERT, text="relevant"),
            DiffChunk(op=DiffOp.EQUAL, text="info."),
        ]
        segments = map_diff_to_runs(chunks, runs)

        # Verify no garbling: equal text should include original tabs
        equal_text = "".join(s.text for s in segments if s.op == DiffOp.EQUAL)
        assert "\t" in equal_text, "Tabs from original runs must be preserved"
        assert "1." in equal_text
        assert "The term means all" in equal_text
        assert "info." in equal_text

        # Verify insertion is present
        insert_segs = [s for s in segments if s.op == DiffOp.INSERT]
        assert len(insert_segs) == 1
        assert "relevant" in insert_segs[0].text

    def test_multiple_consecutive_tabs(self):
        """Multiple tabs at the start should all be consumed as whitespace."""
        p = etree.Element(qn("w", "p"))
        r = etree.SubElement(p, qn("w", "r"))
        etree.SubElement(r, qn("w", "tab"))
        etree.SubElement(r, qn("w", "tab"))
        etree.SubElement(r, qn("w", "tab"))
        t = etree.SubElement(r, qn("w", "t"))
        t.text = "Hello world"

        runs = extract_runs(p)
        raw = get_paragraph_text(runs)
        assert raw == "\t\t\tHello world"

        # Diff just sees "Hello world" (tabs stripped by tokenizer)
        chunks = [DiffChunk(op=DiffOp.EQUAL, text="Hello world")]
        segments = map_diff_to_runs(chunks, runs)

        combined = "".join(s.text for s in segments)
        assert combined == "\t\t\tHello world"
        assert all(s.op == DiffOp.EQUAL for s in segments)

    def test_insert_in_middle_has_spaces(self):
        """Insert between two words must have spaces on both sides.

        "The quick fox" → "The quick brown fox"
        Regression test for Bug 5 (inter-chunk word boundary spacing).
        """
        p = _make_paragraph(("The quick fox", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The quick"),
            DiffChunk(op=DiffOp.INSERT, text="brown"),
            DiffChunk(op=DiffOp.EQUAL, text="fox"),
        ]
        segments = map_diff_to_runs(chunks, runs)

        # Reconstruct new text (EQUAL + INSERT)
        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "The quick brown fox"

        # Reconstruct old text (EQUAL + DELETE)
        old_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.DELETE))
        assert old_text == "The quick fox"

    def test_insert_at_start_has_trailing_space(self):
        """Insert at the start of a paragraph needs a space after it.

        "world." → "Hello world."
        Regression test for Bug 5 (inter-chunk word boundary spacing).
        """
        p = _make_paragraph(("world.", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.INSERT, text="Hello"),
            DiffChunk(op=DiffOp.EQUAL, text="world."),
        ]
        segments = map_diff_to_runs(chunks, runs)

        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "Hello world."

    def test_insert_at_end_has_leading_space(self):
        """Insert at the end needs a space before it.

        "Hello world." → "Hello world. Goodbye."
        Regression test for Bug 5 (inter-chunk word boundary spacing).
        """
        p = _make_paragraph(("Hello world.", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="Hello world."),
            DiffChunk(op=DiffOp.INSERT, text="Goodbye."),
        ]
        segments = map_diff_to_runs(chunks, runs)

        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "Hello world. Goodbye."

    def test_replace_word_has_spaces(self):
        """Replacing a word must preserve spaces around the replacement.

        "The quick fox" → "The slow fox"
        Regression test for Bug 5 (inter-chunk word boundary spacing).
        """
        p = _make_paragraph(("The quick fox", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The"),
            DiffChunk(op=DiffOp.DELETE, text="quick"),
            DiffChunk(op=DiffOp.INSERT, text="slow"),
            DiffChunk(op=DiffOp.EQUAL, text="fox"),
        ]
        segments = map_diff_to_runs(chunks, runs)

        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "The slow fox"

        old_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.DELETE))
        assert old_text == "The quick fox"

    def test_insert_multiple_words_in_middle(self):
        """Insert multiple words in the middle of a sentence.

        "The fox jumps." → "The quick brown fox jumps."
        """
        p = _make_paragraph(("The fox jumps.", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The"),
            DiffChunk(op=DiffOp.INSERT, text="quick brown"),
            DiffChunk(op=DiffOp.EQUAL, text="fox jumps."),
        ]
        segments = map_diff_to_runs(chunks, runs)

        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "The quick brown fox jumps."

    def test_consecutive_insert_delete_insert(self):
        """Multiple adjacent inserts and deletes preserve spacing.

        "A B C" → "X Y Z"  (full replacement)
        """
        p = _make_paragraph(("A B C", None))
        runs = extract_runs(p)
        chunks = [
            DiffChunk(op=DiffOp.DELETE, text="A B C"),
            DiffChunk(op=DiffOp.INSERT, text="X Y Z"),
        ]
        segments = map_diff_to_runs(chunks, runs)

        new_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.INSERT))
        assert new_text == "X Y Z"

        old_text = "".join(s.text for s in segments if s.op in (DiffOp.EQUAL, DiffOp.DELETE))
        assert old_text == "A B C"


# ===================================================================
# _inject_inter_chunk_spaces tests
# ===================================================================


class TestInjectInterChunkSpaces:
    def test_empty_list(self):
        assert _inject_inter_chunk_spaces([]) == []

    def test_single_chunk(self):
        chunks = [DiffChunk(op=DiffOp.EQUAL, text="hello")]
        result = _inject_inter_chunk_spaces(chunks)
        assert len(result) == 1
        assert result[0].text == "hello"

    def test_adds_space_between_abutting_chunks(self):
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The quick"),
            DiffChunk(op=DiffOp.INSERT, text="brown"),
            DiffChunk(op=DiffOp.EQUAL, text="fox"),
        ]
        result = _inject_inter_chunk_spaces(chunks)
        assert result[0].text == "The quick"  # EQUAL unchanged
        assert result[1].text == " brown "  # INSERT gets both boundary spaces
        assert result[2].text == "fox"  # EQUAL unchanged

    def test_no_space_when_already_present(self):
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="hello "),
            DiffChunk(op=DiffOp.INSERT, text="world"),
        ]
        result = _inject_inter_chunk_spaces(chunks)
        assert result[1].text == "world"  # No extra space

    def test_no_space_between_delete_and_insert_replacement(self):
        """DELETE→INSERT: left boundary skips DELETE, finds EQUAL("The")."""
        chunks = [
            DiffChunk(op=DiffOp.EQUAL, text="The"),
            DiffChunk(op=DiffOp.DELETE, text="quick"),
            DiffChunk(op=DiffOp.INSERT, text="slow"),
            DiffChunk(op=DiffOp.EQUAL, text="fox"),
        ]
        result = _inject_inter_chunk_spaces(chunks)
        # EQUAL and DELETE chunks are never modified
        assert result[0].text == "The"
        assert result[1].text == "quick"
        # INSERT left boundary: skip DELETE, find EQUAL("The") ends 'e' → needs space
        # INSERT right boundary: skip nothing, EQUAL("fox") starts 'f' → needs space
        assert result[2].text == " slow "
        assert result[3].text == "fox"

    def test_preserves_ops(self):
        chunks = [
            DiffChunk(op=DiffOp.INSERT, text="Hello"),
            DiffChunk(op=DiffOp.EQUAL, text="world"),
        ]
        result = _inject_inter_chunk_spaces(chunks)
        assert result[0].op == DiffOp.INSERT
        assert result[1].op == DiffOp.EQUAL
        # Only the INSERT gets the trailing space; EQUAL is untouched
        assert result[0].text == "Hello "
        assert result[1].text == "world"


# ===================================================================
# build_run_element tests
# ===================================================================


class TestBuildRunElement:
    def test_plain_text_run(self):
        r = build_run_element("Hello")
        assert r.tag == qn("w", "r")
        t = r.find(qn("w", "t"))
        assert t is not None
        assert t.text == "Hello"

    def test_delete_text_run(self):
        r = build_run_element("deleted", is_delete=True)
        dt = r.find(qn("w", "delText"))
        assert dt is not None
        assert dt.text == "deleted"
        # Should NOT have <w:t>
        assert r.find(qn("w", "t")) is None

    def test_preserves_whitespace(self):
        r = build_run_element(" hello ")
        t = r.find(qn("w", "t"))
        assert t is not None
        assert t.get(f"{{{XML}}}space") == "preserve"

    def test_no_space_preserve_for_normal_text(self):
        r = build_run_element("hello")
        t = r.find(qn("w", "t"))
        assert t is not None
        assert t.get(f"{{{XML}}}space") is None

    def test_includes_rpr(self):
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        r = build_run_element("bold", rpr=rpr)
        found_rpr = r.find(qn("w", "rPr"))
        assert found_rpr is not None
        assert found_rpr.find(qn("w", "b")) is not None
        # Should be a copy, not the original
        assert found_rpr is not rpr


# ===================================================================
# build_tracked_change_element tests
# ===================================================================


class TestBuildTrackedChangeElement:
    def test_ins_element(self):
        el = build_tracked_change_element(
            "ins",
            change_id=42,
            author="AI Review",
            date_iso="2026-01-15T10:30:00Z",
        )
        assert el.tag == qn("w", "ins")
        assert el.get(qn("w", "id")) == "42"
        assert el.get(qn("w", "author")) == "AI Review"
        assert el.get(qn("w", "date")) == "2026-01-15T10:30:00Z"

    def test_del_element(self):
        el = build_tracked_change_element(
            "del",
            change_id=99,
            author="Reviewer",
            date_iso="2026-02-01T00:00:00Z",
        )
        assert el.tag == qn("w", "del")
        assert el.get(qn("w", "id")) == "99"


# ===================================================================
# _rpr_equal tests
# ===================================================================


class TestRprEqual:
    def test_both_none(self):
        assert _rpr_equal(None, None) is True

    def test_first_none(self):
        rpr = etree.Element(qn("w", "rPr"))
        assert _rpr_equal(None, rpr) is False

    def test_second_none(self):
        rpr = etree.Element(qn("w", "rPr"))
        assert _rpr_equal(rpr, None) is False

    def test_same_object(self):
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        assert _rpr_equal(rpr, rpr) is True

    def test_deep_copy_equal(self):
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        cloned = copy.deepcopy(rpr)
        assert cloned is not rpr
        assert _rpr_equal(rpr, cloned) is True

    def test_different_content(self):
        rpr_bold = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr_bold, qn("w", "b"))

        rpr_italic = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr_italic, qn("w", "i"))

        assert _rpr_equal(rpr_bold, rpr_italic) is False

    def test_empty_rprs_equal(self):
        rpr_a = etree.Element(qn("w", "rPr"))
        rpr_b = etree.Element(qn("w", "rPr"))
        assert _rpr_equal(rpr_a, rpr_b) is True

    def test_complex_rpr_equal(self):
        """rPr with bold + underline should equal a deepcopy."""
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        u = etree.SubElement(rpr, qn("w", "u"))
        u.set(qn("w", "val"), "single")
        cloned = copy.deepcopy(rpr)
        assert _rpr_equal(rpr, cloned) is True


# ===================================================================
# _merge_tagged_segments tests
# ===================================================================


class TestMergeTaggedSegments:
    def test_empty_list(self):
        assert _merge_tagged_segments([]) == []

    def test_single_segment(self):
        segs = [TaggedSegment(text="hello", op=DiffOp.EQUAL)]
        result = _merge_tagged_segments(segs)
        assert len(result) == 1
        assert result[0].text == "hello"

    def test_merges_same_op_none_rpr(self):
        segs = [
            TaggedSegment(text="he", op=DiffOp.EQUAL),
            TaggedSegment(text="llo", op=DiffOp.EQUAL),
        ]
        result = _merge_tagged_segments(segs)
        assert len(result) == 1
        assert result[0].text == "hello"

    def test_does_not_merge_different_ops(self):
        segs = [
            TaggedSegment(text="old", op=DiffOp.DELETE),
            TaggedSegment(text="new", op=DiffOp.INSERT),
        ]
        result = _merge_tagged_segments(segs)
        assert len(result) == 2

    def test_merges_cloned_rpr(self):
        """Segments with deepcopy'd rPr of same content should merge.

        This is the Bug 2 regression test: clone_rpr() creates different
        objects but _merge_tagged_segments should still merge them because
        _rpr_equal compares by content, not identity.
        """
        rpr = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr, qn("w", "b"))
        segs = [
            TaggedSegment(text="he", op=DiffOp.EQUAL, rpr=clone_rpr(rpr)),
            TaggedSegment(text="llo", op=DiffOp.EQUAL, rpr=clone_rpr(rpr)),
        ]
        # The two rpr objects are different instances
        assert segs[0].rpr is not segs[1].rpr
        result = _merge_tagged_segments(segs)
        assert len(result) == 1
        assert result[0].text == "hello"

    def test_does_not_merge_different_rpr(self):
        rpr_bold = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr_bold, qn("w", "b"))
        rpr_italic = etree.Element(qn("w", "rPr"))
        etree.SubElement(rpr_italic, qn("w", "i"))

        segs = [
            TaggedSegment(text="he", op=DiffOp.EQUAL, rpr=rpr_bold),
            TaggedSegment(text="llo", op=DiffOp.EQUAL, rpr=rpr_italic),
        ]
        result = _merge_tagged_segments(segs)
        assert len(result) == 2
