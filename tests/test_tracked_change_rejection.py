"""Tests for T1.1: hard-reject pre-existing tracked changes."""

from __future__ import annotations

import pytest
from fastmcp import Client
from fastmcp.exceptions import ToolError

from docx_mcp.document import DocxDocument
from docx_mcp.redliner import apply_redlines
from docx_mcp.server import mcp

pytestmark = pytest.mark.anyio


class TestTrackedChangeRejection:
    """Documents with pre-existing tracked changes must be rejected."""

    def test_pristine_document_accepted(self, simple_5para_path):
        """Clean documents have no tracked changes."""
        doc = DocxDocument(path=simple_5para_path)
        assert doc.has_tracked_changes() == []

    def test_body_tracked_changes_rejected(self, body_tracked_changes_path):
        """Document with <w:ins> in body is flagged."""
        doc = DocxDocument(path=body_tracked_changes_path)
        dirty = doc.has_tracked_changes()
        assert "word/document.xml" in dirty

    def test_header_tracked_changes_rejected(self, header_tracked_changes_path):
        """Document with <w:del> in header is flagged."""
        doc = DocxDocument(path=header_tracked_changes_path)
        dirty = doc.has_tracked_changes()
        assert any("header" in part for part in dirty)

    def test_footer_tracked_changes_rejected(self, footer_tracked_changes_path):
        """Document with <w:moveFrom> in footer is flagged."""
        doc = DocxDocument(path=footer_tracked_changes_path)
        dirty = doc.has_tracked_changes()
        assert any("footer" in part for part in dirty)

    def test_comments_tracked_changes_rejected(self, comments_tracked_changes_path):
        """Document with <w:moveTo> in comments is flagged."""
        doc = DocxDocument(path=comments_tracked_changes_path)
        dirty = doc.has_tracked_changes()
        assert "word/comments.xml" in dirty

    def test_apply_redlines_rejects_body_tracked_changes(
        self,
        body_tracked_changes_path,
    ):
        """apply_redlines raises ValueError for body tracked changes."""
        from docx_mcp.models import ParagraphChange, ParagraphChangeType

        with pytest.raises(ValueError, match="pre-existing tracked changes"):
            apply_redlines(
                body_tracked_changes_path,
                [
                    ParagraphChange(
                        kind="paragraph",
                        fragment_id=1,
                        change_type=ParagraphChangeType.MODIFY,
                        new_text="Modified text.",
                        justification="Test.",
                    ),
                ],
            )

    async def test_extract_fragments_rejects_body_tracked_changes(
        self,
        body_tracked_changes_path,
    ):
        """extract_fragments raises ToolError for body tracked changes."""
        async with Client(mcp) as client:
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(body_tracked_changes_path)},
                )

    async def test_extract_fragments_rejects_header_tracked_changes(
        self,
        header_tracked_changes_path,
    ):
        """extract_fragments raises ToolError for header tracked changes."""
        async with Client(mcp) as client:
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(header_tracked_changes_path)},
                )

    async def test_extract_fragments_rejects_footer_tracked_changes(
        self,
        footer_tracked_changes_path,
    ):
        """extract_fragments raises ToolError for footer tracked changes."""
        async with Client(mcp) as client:
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(footer_tracked_changes_path)},
                )

    async def test_extract_fragments_rejects_comments_tracked_changes(
        self,
        comments_tracked_changes_path,
    ):
        """extract_fragments raises ToolError for comments tracked changes."""
        async with Client(mcp) as client:
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(comments_tracked_changes_path)},
                )
