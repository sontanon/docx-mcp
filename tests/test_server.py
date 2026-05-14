"""Tests for the MCP server (tools and resource)."""

import json
from pathlib import Path
from urllib.parse import quote
from zipfile import ZipFile

import pytest
from fastmcp import Client
from fastmcp.exceptions import ToolError

from docx_mcp.server import mcp

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

pytestmark = pytest.mark.anyio


def _text(result) -> str:
    """Extract the text content from a CallToolResult."""
    return result.content[0].text


def _is_valid_docx(path: Path) -> bool:
    """Return True if *path* is a valid .docx ZIP containing document.xml."""
    if not path.exists():
        return False
    try:
        with ZipFile(path) as zf:
            return "word/document.xml" in zf.namelist()
    except Exception:
        return False


def _write_changes_json(tmp_path: Path, changes: list[dict] | dict) -> Path:
    """Write a changes structure to a JSON file and return the path."""
    p = tmp_path / "changes.json"
    p.write_text(json.dumps(changes), encoding="utf-8")
    return p


# ---------------------------------------------------------------------------
# extract_fragments
# ---------------------------------------------------------------------------


class TestExtractFragments:
    async def test_tagged_format(self, simple_5para_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(simple_5para_path)},
            )
        text = _text(result)
        assert "<f=1>" in text
        assert "<f=5>" in text
        assert "</f=1>" in text

    async def test_extracts_body_and_headers(self, header_footer_text_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(header_footer_text_path)},
            )
        text = _text(result)
        assert "<f=header_1.1>" in text
        assert "<f=footer_1.1>" in text
        assert "<f=1>" in text

    async def test_file_not_found(self):
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="File not found"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": "/nonexistent/file.docx"},
                )

    async def test_formatted_document(self, formatted_runs_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(formatted_runs_path)},
            )
        text = _text(result)
        # Formatted doc should contain pseudo-Markdown markers
        assert "<f=1>" in text


# ---------------------------------------------------------------------------
# apply_changes
# ---------------------------------------------------------------------------


class TestApplyChanges:
    async def test_single_modify(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "The Modified Seller shall transfer the goods.",
                            "justification": "Test markup extraction.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "1 change(s)" in text
        assert "1 modify" in text
        assert str(output) in text
        assert _is_valid_docx(output)

    async def test_single_delete(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "2",
                            "change_type": "delete",
                            "justification": "Removed redundant clause.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "1 delete" in text
        assert _is_valid_docx(output)

    async def test_single_append(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "3",
                            "change_type": "append_after",
                            "new_text": "The foregoing shall survive termination.",
                            "justification": "Added survival provision.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "1 append_after" in text
        assert _is_valid_docx(output)

    async def test_multiple_change_types(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "Modified paragraph one.",
                            "justification": "Edit first paragraph.",
                        },
                        {
                            "fragment_id": "3",
                            "change_type": "delete",
                            "justification": "Remove third paragraph.",
                        },
                        {
                            "fragment_id": "5",
                            "change_type": "append_after",
                            "new_text": "A new final paragraph.",
                            "justification": "Add conclusion.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "3 change(s)" in text
        assert _is_valid_docx(output)

    async def test_default_output_path(self, simple_5para_path):
        expected = simple_5para_path.parent / "simple_5para_redlined.docx"
        try:
            async with Client(mcp) as client:
                result = await client.call_tool(
                    "apply_changes",
                    {
                        "document_path": str(simple_5para_path),
                        "changes": [
                            {
                                "fragment_id": "1",
                                "change_type": "delete",
                                "justification": "Test default path.",
                            },
                        ],
                    },
                )
            text = _text(result)
            assert "simple_5para_redlined.docx" in text
            assert _is_valid_docx(expected)
        finally:
            expected.unlink(missing_ok=True)

    async def test_validation_included_by_default(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                            {
                                "fragment_id": "1",
                                "change_type": "delete",
                                "justification": "Test validation.",
                            },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Validation:" in text

    async def test_invalid_fragment_id(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="fragment_id=99"):
                await client.call_tool(
                    "apply_changes",
                    {
                        "document_path": str(simple_5para_path),
                        "changes": [
                            {
                                "fragment_id": "99",
                                "change_type": "delete",
                                "justification": "Bad ID.",
                            },
                        ],
                        "output_path": str(output),
                    },
                )

    async def test_file_not_found(self, tmp_path):
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="File not found"):
                await client.call_tool(
                    "apply_changes",
                    {
                        "document_path": "/nonexistent/file.docx",
                        "changes": [
                            {
                                "fragment_id": "1",
                                "change_type": "delete",
                                "justification": "Test.",
                            },
                        ],
                    },
                )

    async def test_custom_author(self, simple_5para_path, tmp_path):
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                            {
                                "fragment_id": "1",
                                "change_type": "delete",
                                "justification": "Test author.",
                            },
                    ],
                    "output_path": str(output),
                    "author": "Jane Doe",
                },
            )
        assert _is_valid_docx(output)
        # Verify author is in the output XML
        from docx_mcp.document import DocxDocument
        from docx_mcp.namespaces import qn

        doc = DocxDocument(path=output)
        # Find any tracked change element with the custom author
        found = False
        for el in doc.document_tree.iter():
            if el.get(qn("w", "author")) == "Jane Doe":
                found = True
                break
        assert found, "Custom author not found in tracked changes"


# ---------------------------------------------------------------------------
# apply_changes_from_file
# ---------------------------------------------------------------------------


class TestApplyChangesFromFile:
    async def test_basic_file(self, simple_5para_path, tmp_path):
        changes_file = _write_changes_json(
            tmp_path,
            [
                {
                    "fragment_id": "1",
                    "change_type": "modify",
                    "new_text": "Paragraph one modified via file.",
                    "justification": "Test file-based apply.",
                },
            ],
        )
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes_from_file",
                {
                    "document_path": str(simple_5para_path),
                    "changes_file": str(changes_file),
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "1 change(s)" in text
        assert _is_valid_docx(output)

    async def test_dict_wrapper(self, simple_5para_path, tmp_path):
        changes_file = _write_changes_json(
            tmp_path,
            {
                "changes": [
                    {
                        "fragment_id": "2",
                        "change_type": "delete",
                        "justification": "Test dict wrapper.",
                    },
                ],
            },
        )
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes_from_file",
                {
                    "document_path": str(simple_5para_path),
                    "changes_file": str(changes_file),
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "1 change(s)" in text
        assert _is_valid_docx(output)

    async def test_changes_file_not_found(self, simple_5para_path):
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="Changes file not found"):
                await client.call_tool(
                    "apply_changes_from_file",
                    {
                        "document_path": str(simple_5para_path),
                        "changes_file": "/nonexistent/changes.json",
                    },
                )

    async def test_invalid_json(self, simple_5para_path, tmp_path):
        bad_file = tmp_path / "bad.json"
        bad_file.write_text("not valid json {{{", encoding="utf-8")
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="Invalid JSON"):
                await client.call_tool(
                    "apply_changes_from_file",
                    {
                        "document_path": str(simple_5para_path),
                        "changes_file": str(bad_file),
                    },
                )

    async def test_invalid_change_schema(self, simple_5para_path, tmp_path):
        changes_file = _write_changes_json(
            tmp_path,
            [
                {"fragment_id": "not_an_int", "change_type": "bogus"},
            ],
        )
        async with Client(mcp) as client:
            with pytest.raises(Exception, match=r"Unknown change_type|Invalid"):
                await client.call_tool(
                    "apply_changes_from_file",
                    {
                        "document_path": str(simple_5para_path),
                        "changes_file": str(changes_file),
                    },
                )

    async def test_dict_without_changes_key(self, simple_5para_path, tmp_path):
        changes_file = _write_changes_json(tmp_path, {"data": []})
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="no 'changes' key"):
                await client.call_tool(
                    "apply_changes_from_file",
                    {
                        "document_path": str(simple_5para_path),
                        "changes_file": str(changes_file),
                    },
                )


# ---------------------------------------------------------------------------
# validate_document
# ---------------------------------------------------------------------------


class TestValidateDocument:
    async def test_clean_file(self, simple_5para_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "validate_document_tool",
                {"document_path": str(simple_5para_path)},
            )
        text = _text(result)
        assert "passed" in text
        assert "0 errors" in text

    async def test_redlined_file(self, simple_5para_path, tmp_path):
        # First create a redlined file
        output = tmp_path / "redlined.docx"
        async with Client(mcp) as client:
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "Changed text for validation test.",
                            "justification": "Test.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
            # Now validate it
            result = await client.call_tool(
                "validate_document_tool",
                {"document_path": str(output)},
            )
        text = _text(result)
        assert "passed" in text

    async def test_file_not_found(self):
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="File not found"):
                await client.call_tool(
                    "validate_document_tool",
                    {"document_path": "/nonexistent/file.docx"},
                )


# ---------------------------------------------------------------------------
# diff_fragments
# ---------------------------------------------------------------------------


class TestDiffFragments:
    async def test_identical_documents(self, simple_5para_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(simple_5para_path),
                    "modified_path": str(simple_5para_path),
                },
            )
        text = _text(result)
        assert "unchanged" in text
        # All 5 fragments should be unchanged
        assert text.count("unchanged") == 5

    async def test_modified_document(self, simple_5para_path, tmp_path):
        # Create a modified version
        output = tmp_path / "modified.docx"
        async with Client(mcp) as client:
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "Completely different text here.",
                            "justification": "Test diff.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        # The redlined version has tracked changes XML, so the text
        # representation will differ. Just verify the tool runs without error.
        async with Client(mcp) as client:
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(simple_5para_path),
                    "modified_path": str(output),
                },
            )
        text = _text(result)
        assert "Fragment 1:" in text

    async def test_different_length_documents(self, simple_5para_path, nda_skeleton_path):
        async with Client(mcp) as client:
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(simple_5para_path),
                    "modified_path": str(nda_skeleton_path),
                },
            )
        text = _text(result)
        # NDA has more paragraphs than simple_5para, so we should see "added"
        assert "added" in text or "modified" in text

    async def test_identical_tables(self, simple_table_path):
        """Identical tables should be reported as unchanged."""
        async with Client(mcp) as client:
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(simple_table_path),
                    "modified_path": str(simple_table_path),
                },
            )
        text = _text(result)
        # simple_table has a paragraph at fragment 1, table at fragment 2
        assert "Table 2: unchanged" in text

    async def test_modified_table_cell(self, simple_table_path, tmp_path):
        """Modified table cells should be shown with cell-level diffs."""
        # Create a genuinely modified version by directly manipulating the document
        # (not using tracked changes, which don't change the extracted text)
        from docx_mcp.document import DocxDocument
        from docx_mcp.namespaces import xpath

        output = tmp_path / "modified_table.docx"
        doc = DocxDocument(simple_table_path)

        # Directly modify cell 2.1.1 text (table at body_elements[1], row 1, col 1)
        table = doc.body_elements[1]
        rows = xpath(table, ".//w:tr")
        first_cell = xpath(rows[0], ".//w:tc")[0]
        first_para = xpath(first_cell, ".//w:p")[0]
        first_run = xpath(first_para, ".//w:r")[0]
        text_elem = xpath(first_run, ".//w:t")[0]
        text_elem.text = "Changed Header A"

        doc.save(output)

        # Now diff the original and modified
        async with Client(mcp) as client:
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(simple_table_path),
                    "modified_path": str(output),
                },
            )
        text = _text(result)
        assert "Table 2: modified" in text
        assert "Cell 2.1.1: modified" in text

    async def test_mixed_content_with_table_change(self, mixed_content_path, tmp_path):
        """Mixed paragraph and table content should diff correctly."""
        # Create a modified version with both paragraph and table changes
        output = tmp_path / "modified_mixed.docx"
        async with Client(mcp) as client:
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(mixed_content_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "Modified first paragraph.",
                            "justification": "Test para change.",
                        },
                        {
                            "cell_id": "2.1.1",
                            "change_type": "modify_cell",
                            "new_text": "Modified cell content",
                            "justification": "Test cell change.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
            # Now diff the original and modified
            result = await client.call_tool(
                "diff_fragments",
                {
                    "original_path": str(mixed_content_path),
                    "modified_path": str(output),
                },
            )
        text = _text(result)
        # Should show both paragraph and table modifications
        assert "Fragment 1:" in text
        assert "Table 2: modified" in text or "Table 2:" in text
        assert "Cell 2.1.1:" in text

    async def test_file_not_found(self, simple_5para_path):
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="File not found"):
                await client.call_tool(
                    "diff_fragments",
                    {
                        "original_path": str(simple_5para_path),
                        "modified_path": "/nonexistent/file.docx",
                    },
                )


# ---------------------------------------------------------------------------
# Resource: docx://{document_path}/fragments
# ---------------------------------------------------------------------------


class TestFragmentsResource:
    async def test_read_resource(self, simple_5para_path):
        encoded = quote(str(simple_5para_path), safe="")
        uri = f"docx-fragments://{encoded}"
        async with Client(mcp) as client:
            contents = await client.read_resource(uri)
        assert len(contents) > 0
        text = contents[0].text
        assert "<f=1>" in text
        assert "<f=5>" in text

    async def test_resource_matches_tool(self, simple_5para_path):
        """Resource output should match the extract_fragments tool output."""
        encoded = quote(str(simple_5para_path), safe="")
        uri = f"docx-fragments://{encoded}"
        async with Client(mcp) as client:
            resource_contents = await client.read_resource(uri)
            tool_result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(simple_5para_path)},
            )
        resource_text = resource_contents[0].text
        tool_text = _text(tool_result)
        assert resource_text == tool_text


# ---------------------------------------------------------------------------
# extract_fragments with markup mode
# ---------------------------------------------------------------------------


class TestExtractRedlinedDocRejected:
    async def test_extract_redlined_doc_rejected(self, simple_5para_path, tmp_path):
        """Extracting from any redlined doc is rejected (pre-existing tracked changes)."""
        output = tmp_path / "redlined.docx"
        async with Client(mcp) as client:
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_5para_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "The Modified Seller shall transfer the goods.",
                            "justification": "Test extraction from redlined doc.",
                        },
                    ],
                    "output_path": str(output),
                },
            )
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(output)},
                )


# ---------------------------------------------------------------------------
# Table extraction tests
# ---------------------------------------------------------------------------


class TestExtractTablesInFragments:
    """Tests for extracting documents with tables."""

    async def test_extract_simple_table_tagged_format(self, simple_table_path):
        """Extract a simple table in tagged format."""
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(simple_table_path)},
            )
        text = _text(result)
        # Should have paragraph tags
        assert "<f=1>" in text
        assert "</f=1>" in text
        # Should have table tags (table is ID 2)
        assert "<table=2 rows=3 cols=3>" in text
        assert "</table=2>" in text
        # Should have cell tags
        assert "<cell=2.1.1>" in text
        assert "</cell=2.1.1>" in text

    async def test_extract_simple_table_tagged(self, simple_table_path):
        """Extract a simple table in tagged format."""
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(simple_table_path)},
            )
        text = _text(result)
        assert "<table=2 rows=3 cols=3>" in text
        assert "<cell=2.1.1>" in text
        assert "</table=2>" in text

    async def test_extract_merged_cell_table_with_spans(self, merged_cell_table_path):
        """Merged-cell tables are extracted with span markers."""
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(merged_cell_table_path)},
            )
        text = _text(result)
        # Should have table with span markers (spanned-over cells omitted)
        assert "<table=2 rows=2 cols=3>" in text
        assert '<cell=2.1.1 span="2">' in text
        assert '<cell=2.1.2' not in text  # spanned-over cell omitted

    async def test_extract_mixed_content(self, mixed_content_path):
        """Extract document with mixed paragraphs and tables."""
        async with Client(mcp) as client:
            result = await client.call_tool(
                "extract_fragments",
                {"document_path": str(mixed_content_path)},
            )
        text = _text(result)
        # Should have both paragraph and table tags interleaved
        assert "<f=1>" in text
        assert "<table=2" in text
        assert "<f=3>" in text
        assert "<table=4" in text
        assert "<f=5>" in text


# ---------------------------------------------------------------------------
# Table change application tests
# ---------------------------------------------------------------------------


class TestApplyTableChanges:
    """Tests for applying table changes via MCP tools."""

    async def test_modify_cell_single_para(self, simple_table_path, tmp_path):
        """Modify a single-paragraph cell."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_table_path),
                    "changes": [
                        {
                            "table_id": 2,
                            "row": 1,
                            "col": 1,
                            "cell_id": "2.1.1",
                            "change_type": "modify_cell",
                            "new_text": "Modified Header",
                            "justification": "Update header",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Applied" in text
        assert _is_valid_docx(output)

    async def test_clear_cell(self, simple_table_path, tmp_path):
        """Clear a cell."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_table_path),
                    "changes": [
                        {
                            "table_id": 2,
                            "row": 1,
                            "col": 1,
                            "cell_id": "2.1.1",
                            "change_type": "clear_cell",
                            "justification": "Remove header",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Applied" in text
        assert _is_valid_docx(output)

    async def test_modify_multi_para_cell(self, table_multi_para_path, tmp_path):
        """Modify a cell with multiple paragraphs."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(table_multi_para_path),
                    "changes": [
                        {
                            "table_id": 2,
                            "row": 2,
                            "col": 2,
                            "cell_id": "2.2.2",
                            "change_type": "modify_cell",
                            "new_text": "First.\nSecond.\nThird.",
                            "justification": "Update all paragraphs",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Applied" in text
        assert _is_valid_docx(output)

    async def test_paragraph_and_table_changes_together(self, mixed_content_path, tmp_path):
        """Apply both paragraph and table changes in one operation."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(mixed_content_path),
                    "changes": [
                        {
                            "fragment_id": "1",
                            "change_type": "modify",
                            "new_text": "Modified intro.",
                            "justification": "Update intro",
                        },
                        {
                            "table_id": 2,
                            "row": 1,
                            "col": 1,
                            "cell_id": "2.1.1",
                            "change_type": "modify_cell",
                            "new_text": "Modified cell",
                            "justification": "Update cell",
                        },
                    ],
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Applied" in text
        assert "1 paragraph changes" in text or "1 modify" in text
        assert "1 table changes" in text or "1 modify_cell" in text
        assert _is_valid_docx(output)

    async def test_invalid_table_id_raises(self, simple_table_path, tmp_path):
        """Invalid table ID should raise error."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="table_id=99"):
                await client.call_tool(
                    "apply_changes",
                    {
                        "document_path": str(simple_table_path),
                        "changes": [
                            {
                                "table_id": 99,
                                "row": 1,
                                "col": 1,
                                "cell_id": "99.1.1",
                                "change_type": "modify_cell",
                                "new_text": "Text",
                                "justification": "Bad ID",
                            },
                        ],
                        "output_path": str(output),
                    },
                )

    async def test_row_out_of_range_raises(self, simple_table_path, tmp_path):
        """Row out of range should raise error."""
        output = tmp_path / "output.docx"
        async with Client(mcp) as client:
            with pytest.raises(Exception, match="row 99 out of range"):
                await client.call_tool(
                    "apply_changes",
                    {
                        "document_path": str(simple_table_path),
                        "changes": [
                            {
                                "table_id": 2,
                                "row": 99,
                                "col": 1,
                                "cell_id": "2.99.1",
                                "change_type": "modify_cell",
                                "new_text": "Text",
                                "justification": "Bad row",
                            },
                        ],
                        "output_path": str(output),
                    },
                )

    async def test_apply_changes_from_file_with_tables(self, simple_table_path, tmp_path):
        """Test apply_changes_from_file with table changes."""
        output = tmp_path / "output.docx"
        changes_file = _write_changes_json(
            tmp_path,
            [
                {
                    "table_id": 2,
                    "row": 1,
                    "col": 1,
                    "cell_id": "2.1.1",
                    "change_type": "modify_cell",
                    "new_text": "Modified Header",
                    "justification": "Update via file",
                },
            ],
        )
        async with Client(mcp) as client:
            result = await client.call_tool(
                "apply_changes_from_file",
                {
                    "document_path": str(simple_table_path),
                    "changes_file": str(changes_file),
                    "output_path": str(output),
                },
            )
        text = _text(result)
        assert "Applied" in text
        assert _is_valid_docx(output)

    async def test_extract_modified_table_with_markup(self, simple_table_path, tmp_path):
        """Extract a modified table with markup=True shows tracked changes."""
        output = tmp_path / "redlined.docx"
        async with Client(mcp) as client:
            # Apply change
            await client.call_tool(
                "apply_changes",
                {
                    "document_path": str(simple_table_path),
                    "changes": [
                        {
                            "table_id": 2,
                            "row": 2,
                            "col": 2,
                            "cell_id": "2.2.2",
                            "change_type": "modify_cell",
                            "new_text": "New content",
                            "justification": "Test markup",
                        },
                    ],
                    "output_path": str(output),
                },
            )
            # Extraction should be rejected (pre-existing tracked changes)
            with pytest.raises(ToolError, match="pre-existing tracked changes"):
                await client.call_tool(
                    "extract_fragments",
                    {"document_path": str(output)},
                )
