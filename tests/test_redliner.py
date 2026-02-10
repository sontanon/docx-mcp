"""End-to-end tests for the redliner orchestrator."""

from __future__ import annotations

import io
import zipfile

import pytest

from docx_mcp.converter import paragraph_to_pseudo_markdown
from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType, RedlineConfig
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
            Change(
                fragment_id=99,
                change_type=ChangeType.DELETE,
                justification="Bad ID.",
            ),
        ]
        with pytest.raises(ValueError, match="fragment_id=99"):
            apply_redlines(simple_5para_path, changes)

    def test_zero_fragment_id_raises(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=0,
                change_type=ChangeType.DELETE,
                justification="Zero ID.",
            ),
        ]
        with pytest.raises(ValueError, match="fragment_id=0"):
            apply_redlines(simple_5para_path, changes)

    def test_modify_without_new_text_raises(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text=None,
                justification="Missing text.",
            ),
        ]
        with pytest.raises(ValueError, match="requires new_text"):
            apply_redlines(simple_5para_path, changes)

    def test_append_without_new_text_raises(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.APPEND_AFTER,
                new_text=None,
                justification="Missing text.",
            ),
        ]
        with pytest.raises(ValueError, match="requires new_text"):
            apply_redlines(simple_5para_path, changes)

    def test_empty_changes_returns_unchanged_doc(self, simple_5para_path):
        doc = apply_redlines(simple_5para_path, [])
        data = doc.to_bytes()
        assert _is_valid_docx(data)


# ===================================================================
# Single change type tests
# ===================================================================


class TestSingleDelete:
    def test_delete_produces_tracked_deletion(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=2,
                change_type=ChangeType.DELETE,
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
            Change(
                fragment_id=3,
                change_type=ChangeType.APPEND_AFTER,
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
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
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
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text=old_text + " (amended)",
                justification="Added amendment note.",
            ),
            Change(
                fragment_id=3,
                change_type=ChangeType.DELETE,
                justification="Removed paragraph 3.",
            ),
            Change(
                fragment_id=5,
                change_type=ChangeType.APPEND_AFTER,
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
                Change(
                    fragment_id=fid,
                    change_type=ChangeType.MODIFY,
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
            Change(
                fragment_id=2,
                change_type=ChangeType.APPEND_AFTER,
                new_text="First new paragraph.",
                justification="Added first.",
            ),
            Change(
                fragment_id=2,
                change_type=ChangeType.APPEND_AFTER,
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
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
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
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
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

        old_text = paragraph_to_pseudo_markdown(fmap[target_fid])
        new_text = (
            old_text.replace("shall", "must") if "shall" in old_text else old_text + " (amended)"
        )

        changes = [
            Change(
                fragment_id=target_fid,
                change_type=ChangeType.MODIFY,
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
            Change(
                fragment_id=2,
                change_type=ChangeType.DELETE,
                justification="Removed.",
            ),
            Change(
                fragment_id=4,
                change_type=ChangeType.APPEND_AFTER,
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
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Completely new first paragraph.",
                justification="Rewrote.",
            ),
        ]
        doc = apply_redlines(raw, changes, config=_default_config())
        data = doc.to_bytes()
        assert _is_valid_docx(data)
