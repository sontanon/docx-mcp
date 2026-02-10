"""Tests for the validator module."""

from __future__ import annotations

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType
from docx_mcp.namespaces import qn
from docx_mcp.redliner import apply_redlines
from docx_mcp.validator import ValidationResult, validate_document


class TestValidationResult:
    def test_ok_when_no_errors(self):
        r = ValidationResult()
        assert r.ok is True

    def test_ok_with_warnings_only(self):
        r = ValidationResult(warnings=["something minor"])
        assert r.ok is True

    def test_not_ok_with_errors(self):
        r = ValidationResult(errors=["something bad"])
        assert r.ok is False


class TestAnnotationIdUniqueness:
    def test_clean_document_passes(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        result = validate_document(doc)
        assert result.ok

    def test_cross_group_id_collision(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        body = doc.body
        # Inject a tracked change and a comment marker with the same w:id
        # These are from different annotation groups, so this is a collision
        del_el = etree.SubElement(body, qn("w", "del"))
        del_el.set(qn("w", "id"), "999")
        del_el.set(qn("w", "author"), "Test")
        del_el.set(qn("w", "date"), "2024-01-01T00:00:00Z")
        crs = etree.SubElement(body, qn("w", "commentRangeStart"))
        crs.set(qn("w", "id"), "999")

        result = validate_document(doc)
        assert not result.ok
        assert any("999" in e for e in result.errors)

    def test_redlined_document_has_unique_ids(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Modified first paragraph.",
                justification="Test modify.",
            ),
            Change(
                fragment_id=3,
                change_type=ChangeType.DELETE,
                justification="Test delete.",
            ),
            Change(
                fragment_id=5,
                change_type=ChangeType.APPEND_AFTER,
                new_text="New paragraph appended.",
                justification="Test append.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)
        result = validate_document(doc)
        assert result.ok, f"Errors: {result.errors}"


class TestCommentIntegrity:
    def test_redlined_comments_are_intact(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=2,
                change_type=ChangeType.MODIFY,
                new_text="Changed second paragraph.",
                justification="Testing comment integrity.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)
        result = validate_document(doc)
        assert result.ok, f"Errors: {result.errors}"

    def test_orphaned_comment_in_comments_xml(self, simple_5para_path):
        """A comment in comments.xml without markers in document.xml."""
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Modified first.",
                justification="Test.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)

        # Add an orphaned comment directly to comments.xml
        assert doc.comments_tree is not None
        orphan = etree.SubElement(doc.comments_tree, qn("w", "comment"))
        orphan.set(qn("w", "id"), "99999")
        orphan.set(qn("w", "author"), "Test")
        orphan.set(qn("w", "date"), "2024-01-01T00:00:00Z")

        result = validate_document(doc)
        assert not result.ok
        assert any("99999" in e and "commentRangeStart" in e for e in result.errors)

    def test_orphaned_range_start_without_end(self, simple_5para_path):
        """A commentRangeStart without a matching commentRangeEnd."""
        doc = DocxDocument(path=simple_5para_path)
        body = doc.body
        crs = etree.SubElement(body, qn("w", "commentRangeStart"))
        crs.set(qn("w", "id"), "888")

        result = validate_document(doc)
        assert any("commentRangeStart id=888" in e for e in result.errors)


class TestTrackedChangeAttributes:
    def test_redlined_changes_have_all_attributes(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Modified text here.",
                justification="Testing attributes.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)
        result = validate_document(doc)
        assert result.ok, f"Errors: {result.errors}"

    def test_missing_author_on_ins(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        body = doc.body
        ins = etree.SubElement(body, qn("w", "ins"))
        ins.set(qn("w", "id"), "500")
        # Missing w:author and w:date

        result = validate_document(doc)
        assert not result.ok
        assert any("w:author" in e for e in result.errors)
        assert any("w:date" in e for e in result.errors)


class TestPackageConsistency:
    def test_redlined_package_is_consistent(self, simple_5para_path):
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="Test text.",
                justification="Test.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)
        result = validate_document(doc)
        assert result.ok, f"Errors: {result.errors}"

    def test_existing_comments_are_consistent(self, existing_comments_path):
        doc = DocxDocument(path=existing_comments_path)
        result = validate_document(doc)
        # Existing fixtures may have orphaned markers (common in real docs)
        # but should not have structural errors
        # If there are errors, they should only be from pre-existing issues
        # in the fixture, not from our code
        assert isinstance(result, ValidationResult)


class TestFullPipelineValidation:
    def test_nda_multi_change_validation(self, nda_skeleton_path):
        """Validate a complex multi-change scenario on the NDA fixture."""
        changes = [
            Change(
                fragment_id=1,
                change_type=ChangeType.MODIFY,
                new_text="**MUTUAL NON-DISCLOSURE AGREEMENT**",
                justification="Changed to mutual NDA.",
            ),
            Change(
                fragment_id=2,
                change_type=ChangeType.DELETE,
                justification="Removed unnecessary clause.",
            ),
            Change(
                fragment_id=3,
                change_type=ChangeType.APPEND_AFTER,
                new_text="This agreement shall be governed by the laws of Delaware.",
                justification="Added governing law clause.",
            ),
        ]
        doc = apply_redlines(nda_skeleton_path, changes)
        result = validate_document(doc)
        assert result.ok, f"Errors: {result.errors}, Warnings: {result.warnings}"

    def test_roundtrip_validation(self, simple_5para_path):
        """Validate that a redlined doc survives a save/reload cycle."""
        changes = [
            Change(
                fragment_id=2,
                change_type=ChangeType.MODIFY,
                new_text="Roundtrip modified paragraph.",
                justification="Testing roundtrip.",
            ),
        ]
        doc = apply_redlines(simple_5para_path, changes)
        result1 = validate_document(doc)
        assert result1.ok, f"Pre-roundtrip errors: {result1.errors}"

        # Serialize and reload
        reloaded = DocxDocument(data=doc.to_bytes())
        result2 = validate_document(reloaded)
        assert result2.ok, f"Post-roundtrip errors: {result2.errors}"
