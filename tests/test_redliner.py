"""End-to-end tests for the redliner orchestrator."""

import io
import zipfile

import pytest
from pydantic import ValidationError

from docx_mcp.converter import paragraph_to_pseudo_markdown
from docx_mcp.document import DocxDocument
from docx_mcp.models import ParagraphChange, ParagraphChangeType, RedlineConfig
from docx_mcp.namespaces import xpath
from docx_mcp.redliner import apply_redlines

# ===================================================================
# Helpers
# ===================================================================


def _default_config() -> RedlineConfig:
    return RedlineConfig(author="AI Review")


def _is_valid_docx(data: bytes) -> bool:
    """Check that data is a valid ZIP with document.xml."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            return "word/document.xml" in zf.namelist()
    except zipfile.BadZipFile:
        return False


# ===================================================================
# Validation tests
# ===================================================================


class TestValidation:
    def test_invalid_fragment_id_raises(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=99,
                change_type=ParagraphChangeType.DELETE,
                justification="Bad ID.",
            ),
        ]
        with pytest.raises(ValueError, match="fragment_id=99"):
            apply_redlines(simple_5para_path, changes)

    def test_zero_fragment_id_raises(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=0,
                change_type=ParagraphChangeType.DELETE,
                justification="Zero ID.",
            ),
        ]
        with pytest.raises(ValueError, match="fragment_id=0"):
            apply_redlines(simple_5para_path, changes)

    def test_modify_without_new_text_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match="new_text is required"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text=None,
                justification="Missing text.",
            )

    def test_append_without_new_text_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match="new_text is required"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text=None,
                justification="Missing text.",
            )

    def test_empty_changes_returns_unchanged_doc(self, simple_5para_path):
        doc = apply_redlines(simple_5para_path, [])
        data = doc.to_bytes()
        assert _is_valid_docx(data)

    def test_collapse_empty_changes_id_space(self, simple_5para_path):
        """With collapse_empty=True, empty paragraphs are skipped in ID space."""
        doc = DocxDocument(path=simple_5para_path)
        # Count how many paragraphs are empty
        empty_count = sum(
            1 for p in doc.paragraphs
            if not paragraph_to_pseudo_markdown(p).strip()
        )
        if empty_count == 0:
            pytest.skip("Fixture has no empty paragraphs to test collapse")

        full_map = doc.full_element_map(collapse_empty=False)
        collapsed_map = doc.full_element_map(collapse_empty=True)
        assert len(collapsed_map) == len(full_map) - empty_count

    def test_collapse_empty_wrong_id_raises(self, simple_5para_path):
        """Applying changes with wrong collapse_empty setting raises clear error."""
        doc = DocxDocument(path=simple_5para_path)
        empty_count = sum(
            1 for p in doc.paragraphs
            if not paragraph_to_pseudo_markdown(p).strip()
        )
        if empty_count == 0:
            pytest.skip("Fixture has no empty paragraphs to test mismatch")

        # If we extract with collapse_empty=True, the last body ID is smaller
        collapsed_map = doc.full_element_map(collapse_empty=True)
        max_collapsed = max(
            int(k) for k in collapsed_map if not k.startswith(("header_", "footer_"))
        )

        # Try to apply a change with the original (non-collapsed) max ID
        # but collapse_empty=True — should fail
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=str(max_collapsed + empty_count),
                change_type=ParagraphChangeType.MODIFY,
                new_text="Updated text.",
                justification="Should fail with wrong collapse setting.",
            ),
        ]
        with pytest.raises(ValueError, match="fragment_id"):
            apply_redlines(simple_5para_path, changes, collapse_empty=True)


# ===================================================================
# Single change type tests
# ===================================================================


class TestSingleDelete:
    def test_delete_produces_tracked_deletion(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=2,
                change_type=ParagraphChangeType.DELETE,
                justification="Removed redundant paragraph.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Still has 5 paragraphs (deleted paragraphs remain in XML)
        assert len(doc.paragraphs) == 5

        # Second paragraph should have <w:del>
        para2 = doc.fragment_map()[2]
        del_els = xpath(para2, "w:del")
        assert len(del_els) >= 1

        # Should have a comment
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) == 1

        # Valid docx
        assert _is_valid_docx(doc.to_bytes())


class TestSingleAppend:
    def test_append_inserts_new_paragraph(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="This is a new paragraph.",
                justification="Added clarification.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Should have 6 paragraphs now
        assert len(doc.paragraphs) == 6

        # The new paragraph (index 4, 0-based) should have <w:ins>
        para4 = doc.paragraphs[3]  # After fragment 3
        ins_els = xpath(para4, "w:ins")
        assert len(ins_els) >= 1

        # Should have a comment
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) == 1

        assert _is_valid_docx(doc.to_bytes())


class TestSingleModify:
    def test_modify_produces_tracked_changes(self, simple_5para_path):
        doc_orig = DocxDocument(path=simple_5para_path)
        para1 = doc_orig.fragment_map()[1]
        old_text = paragraph_to_pseudo_markdown(para1)

        # Make a word-level change
        if " " in old_text:
            words = old_text.split()
            words[0] = "Modified"
            new_text = " ".join(words)
        else:
            new_text = "Completely different text."

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text=new_text,
                justification="Updated wording.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Paragraph should have tracked changes
        para = doc.fragment_map()[1]
        del_els = xpath(para, "w:del")
        ins_els = xpath(para, "w:ins")
        assert len(del_els) >= 1 or len(ins_els) >= 1

        # Should have a comment
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) >= 1

        assert _is_valid_docx(doc.to_bytes())


# ===================================================================
# Multiple changes tests
# ===================================================================


class TestMultipleChanges:
    def test_all_three_change_types(self, simple_5para_path):
        """Apply modify, delete, and append in one call."""
        doc_orig = DocxDocument(path=simple_5para_path)
        para1 = doc_orig.fragment_map()[1]
        old_text = paragraph_to_pseudo_markdown(para1)

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text=old_text + " (amended)",
                justification="Added amendment note.",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.DELETE,
                justification="Removed paragraph 3.",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=5,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="New final paragraph.",
                justification="Added conclusion.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Should have 6 paragraphs (5 original + 1 appended)
        assert len(doc.paragraphs) == 6

        # Should have 3 comments
        assert doc.comments_tree is not None
        comments = xpath(doc.comments_tree, "w:comment")
        assert len(comments) == 3

        assert _is_valid_docx(doc.to_bytes())

    def test_multiple_modifies(self, simple_5para_path):
        doc_orig = DocxDocument(path=simple_5para_path)

        changes = []
        for fid in [1, 2, 4]:
            para = doc_orig.fragment_map()[fid]
            paragraph_to_pseudo_markdown(para)  # verify conversion works
            changes.append(
                ParagraphChange(
                    kind="paragraph",
                    fragment_id=fid,
                    change_type=ParagraphChangeType.MODIFY,
                    new_text=f"Replaced text for paragraph {fid}.",
                    justification=f"Updated paragraph {fid}.",
                ),
            )

        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Each modified paragraph should have tracked changes
        for fid in [1, 2, 4]:
            para = doc.fragment_map()[fid]
            has_del = len(xpath(para, "w:del")) > 0
            has_ins = len(xpath(para, "w:ins")) > 0
            assert has_del or has_ins, f"Fragment {fid} should have tracked changes"

        assert _is_valid_docx(doc.to_bytes())

    def test_multiple_appends_same_fragment(self, simple_5para_path):
        """Two appends after the same paragraph (reverse order processing)."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=2,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="First new paragraph.",
                justification="Added first.",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=2,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="Second new paragraph.",
                justification="Added second.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # Should have 7 paragraphs (5 + 2)
        assert len(doc.paragraphs) == 7
        assert _is_valid_docx(doc.to_bytes())


# ===================================================================
# Fixture-specific integration tests
# ===================================================================


class TestWithFormattedDoc:
    def test_modify_formatted_paragraph(self, formatted_runs_path):
        doc_orig = DocxDocument(path=formatted_runs_path)
        fmap = doc_orig.fragment_map()
        para = fmap[1]
        old_text = paragraph_to_pseudo_markdown(para)

        # Just append text
        new_text = old_text + " (reviewed)"

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text=new_text,
                justification="Added review note.",
            ),
        ]
        doc = apply_redlines(formatted_runs_path, changes, config=_default_config())
        assert _is_valid_docx(doc.to_bytes())


class TestWithExistingComments:
    def test_preserves_existing_comments(self, existing_comments_path):
        doc_orig = DocxDocument(path=existing_comments_path)
        original_comment_count = len(xpath(doc_orig.comments_tree, "w:comment"))

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text="Modified first paragraph.",
                justification="New comment.",
            ),
        ]
        doc = apply_redlines(existing_comments_path, changes, config=_default_config())

        new_comment_count = len(xpath(doc.comments_tree, "w:comment"))
        assert new_comment_count == original_comment_count + 1
        assert _is_valid_docx(doc.to_bytes())


class TestWithNdaSkeleton:
    def test_nda_modify_and_delete(self, nda_skeleton_path):
        """Realistic NDA changes."""
        doc_orig = DocxDocument(path=nda_skeleton_path)
        fmap = doc_orig.fragment_map()

        # Find a paragraph with substantial text
        target_fid = None
        for fid, para in fmap.items():
            text = paragraph_to_pseudo_markdown(para)
            if len(text) > 50:
                target_fid = fid
                break

        if target_fid is None:
            pytest.skip("No suitable paragraph found in NDA skeleton")

        assert target_fid is not None  # narrowing for type checkers
        old_text = paragraph_to_pseudo_markdown(fmap[target_fid])
        new_text = (
            old_text.replace("shall", "must") if "shall" in old_text else old_text + " (amended)"
        )

        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=target_fid,
                change_type=ParagraphChangeType.MODIFY,
                new_text=new_text,
                justification="Modernized obligation language.",
            ),
        ]
        doc = apply_redlines(nda_skeleton_path, changes, config=_default_config())

        data = doc.to_bytes()
        assert _is_valid_docx(data)
        assert len(data) > 0


# ===================================================================
# Round-trip test
# ===================================================================


class TestRoundTrip:
    def test_save_and_reload(self, simple_5para_path, tmp_path):
        """Apply changes, save to file, reload, and verify."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=2,
                change_type=ParagraphChangeType.DELETE,
                justification="Removed.",
            ),
            ParagraphChange(
                kind="paragraph",
                fragment_id=4,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="New paragraph.",
                justification="Added.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())
        output_path = tmp_path / "redlined.docx"
        doc.save(output_path)

        # Reload
        doc2 = DocxDocument(path=output_path)
        assert len(doc2.paragraphs) == 6  # 5 + 1 appended
        assert doc2.comments_tree is not None
        comments = xpath(doc2.comments_tree, "w:comment")
        assert len(comments) == 2

    def test_from_bytes(self, simple_5para_path):
        """Load from bytes, apply changes, get bytes back."""
        raw = simple_5para_path.read_bytes()
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text="Completely new first paragraph.",
                justification="Rewrote.",
            ),
        ]
        doc = apply_redlines(raw, changes, config=_default_config())
        data = doc.to_bytes()
        assert _is_valid_docx(data)


# ===================================================================
# Validation: spacing field constraints
# ===================================================================


class TestSpacingValidation:
    """Validate that spacing fields are rejected on wrong change types."""

    def test_blank_lines_before_on_delete_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match=r"blank_lines_before.*only allowed"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove.",
                blank_lines_before=1,
            )

    def test_blank_lines_after_on_modify_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match=r"blank_lines.*only allowed"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text="Changed.",
                justification="Edit.",
                blank_lines_after=1,
            )

    def test_delete_next_blanks_on_append_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match="delete_next_blanks only allowed"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="New.",
                justification="Add.",
                delete_next_blanks=1,
            )

    def test_delete_next_blanks_on_modify_raises(self, simple_5para_path):
        # Pydantic validator catches this at model creation time
        with pytest.raises(ValidationError, match="delete_next_blanks only allowed"):
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.MODIFY,
                new_text="Changed.",
                justification="Edit.",
                delete_next_blanks=1,
            )

    def test_zero_values_accepted_on_any_type(self, simple_5para_path):
        """Default zero values should pass validation on any change type."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove.",
                blank_lines_before=0,
                blank_lines_after=0,
                delete_next_blanks=0,
            ),
        ]
        # Should not raise
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())
        assert _is_valid_docx(doc.to_bytes())


# ===================================================================
# delete_next_blanks tests
# ===================================================================


class TestDeleteNextBlanks:
    """Tests for deleting trailing blank paragraphs alongside a clause."""

    def test_delete_with_one_trailing_blank(self, blank_separated_path):
        """Deleting fragment 1 with delete_next_blanks=1 removes the blank at fragment 2."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.DELETE,
                justification="Removed clause A.",
                delete_next_blanks=1,
            ),
        ]
        doc = apply_redlines(blank_separated_path, changes, config=_default_config())

        # Both fragment 1 (clause) and fragment 2 (blank) should be marked deleted
        fmap = doc.fragment_map()
        del_els_1 = xpath(fmap[1], "w:del")
        assert len(del_els_1) >= 1

        # Fragment 2 is the blank — it should also be deleted (pPr mark)
        ppr_del_2 = xpath(fmap[2], "w:pPr/w:rPr/w:del")
        assert len(ppr_del_2) >= 1

        # Total paragraph count unchanged (deletions are tracked, not removed)
        assert len(doc.paragraphs) == 7

        assert _is_valid_docx(doc.to_bytes())

    def test_delete_next_blanks_non_blank_raises(self, blank_separated_path):
        """If the next paragraph is not blank, ValueError is raised."""
        # Fragment 3 is "The Buyer shall pay...", fragment 4 is blank.
        # Deleting fragment 4 (blank) with delete_next_blanks=1 should fail
        # because fragment 5 ("This Agreement...") is not blank.
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=4,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove blank.",
                delete_next_blanks=1,
            ),
        ]
        with pytest.raises(ValueError, match="not blank"):
            apply_redlines(blank_separated_path, changes, config=_default_config())

    def test_delete_next_blanks_not_enough_siblings(self, blank_separated_path):
        """If there aren't enough paragraphs after the target, ValueError is raised."""
        # Fragment 7 is the last paragraph — only <w:sectPr> follows it
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=7,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove last.",
                delete_next_blanks=1,
            ),
        ]
        with pytest.raises(ValueError, match="delete_next_blanks=1 on fragment 7"):
            apply_redlines(blank_separated_path, changes, config=_default_config())

    def test_delete_next_blanks_two_blanks_but_only_one_exists(self, blank_separated_path):
        """Requesting 2 trailing blanks but only 1 exists raises ValueError."""
        # Fragment 1 is a clause, fragment 2 is blank, fragment 3 is a clause
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove.",
                delete_next_blanks=2,
            ),
        ]
        with pytest.raises(ValueError, match="not blank"):
            apply_redlines(blank_separated_path, changes, config=_default_config())

    def test_delete_next_blanks_zero_unchanged_behavior(self, blank_separated_path):
        """delete_next_blanks=0 should not touch trailing blanks."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=1,
                change_type=ParagraphChangeType.DELETE,
                justification="Remove clause only.",
                delete_next_blanks=0,
            ),
        ]
        doc = apply_redlines(blank_separated_path, changes, config=_default_config())

        fmap = doc.fragment_map()
        # Fragment 1 should be deleted
        del_els_1 = xpath(fmap[1], "w:del")
        assert len(del_els_1) >= 1

        # Fragment 2 (blank) should NOT be deleted
        ppr_del_2 = xpath(fmap[2], "w:pPr/w:rPr/w:del")
        assert len(ppr_del_2) == 0

        assert _is_valid_docx(doc.to_bytes())

    def test_delete_next_blanks_produces_valid_docx(self, blank_separated_path):
        """Round-trip: delete with trailing blank, save, reload."""
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.DELETE,
                justification="Removed clause B.",
                delete_next_blanks=1,
            ),
        ]
        doc = apply_redlines(blank_separated_path, changes, config=_default_config())
        data = doc.to_bytes()
        assert _is_valid_docx(data)

        # Reload and verify structure
        from docx_mcp.document import DocxDocument

        doc2 = DocxDocument(data=data)
        assert len(doc2.paragraphs) == 7


# ===================================================================
# Append with blank lines (end-to-end through redliner)
# ===================================================================


class TestAppendWithBlankLines:
    """End-to-end tests for blank_lines_before/after through apply_redlines."""

    def test_append_with_blank_line_before(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="New clause inserted.",
                justification="Added clause.",
                blank_lines_before=1,
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # 5 original + 1 blank + 1 content = 7
        assert len(doc.paragraphs) == 7
        assert _is_valid_docx(doc.to_bytes())

    def test_append_with_blank_lines_before_and_after(self, simple_5para_path):
        changes = [
            ParagraphChange(
                kind="paragraph",
                fragment_id=3,
                change_type=ParagraphChangeType.APPEND_AFTER,
                new_text="New clause inserted.",
                justification="Added clause.",
                blank_lines_before=1,
                blank_lines_after=1,
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes, config=_default_config())

        # 5 original + 1 blank before + 1 content + 1 blank after = 8
        assert len(doc.paragraphs) == 8
        assert _is_valid_docx(doc.to_bytes())
