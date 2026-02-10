"""Tests for DocxDocument parsing and manipulation."""

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument


class TestDocxDocumentInit:
    """Tests for document loading."""

    def test_load_from_path(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        assert doc.document_tree is not None
        assert doc.body is not None

    def test_load_from_bytes(self, simple_5para_path: Path):
        data = simple_5para_path.read_bytes()
        doc = DocxDocument(data=data)
        assert doc.document_tree is not None

    def test_load_requires_path_or_data(self):
        with pytest.raises(ValueError, match="Provide either path or data"):
            DocxDocument()

    def test_load_rejects_both_path_and_data(self, simple_5para_path: Path):
        data = simple_5para_path.read_bytes()
        with pytest.raises(ValueError, match="Provide either path or data, not both"):
            DocxDocument(path=simple_5para_path, data=data)

    def test_load_invalid_zip_raises(self):
        with pytest.raises(Exception):  # noqa: B017
            DocxDocument(data=b"not a zip file")


class TestParagraphIndexing:
    """Tests for paragraph discovery and fragment mapping."""

    def test_simple_5para_count(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        # python-docx always adds an initial empty paragraph in some templates,
        # so let's just verify we have at least 5 paragraphs
        assert len(doc.paragraphs) >= 5

    def test_fragment_map_keys_are_1_based(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        assert min(fmap.keys()) == 1
        assert max(fmap.keys()) == len(doc.paragraphs)

    def test_fragment_map_values_are_elements(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        for el in fmap.values():
            assert el.tag.endswith("}p")  # qualified name ends with }p

    def test_formatted_runs_has_paragraphs(self, formatted_runs_path: Path):
        doc = DocxDocument(path=formatted_runs_path)
        assert len(doc.paragraphs) >= 5

    def test_nda_skeleton_has_many_paragraphs(self, nda_skeleton_path: Path):
        doc = DocxDocument(path=nda_skeleton_path)
        # NDA has title, parties, recitals, sections — should be many paragraphs
        assert len(doc.paragraphs) >= 10

    def test_styled_headings_has_paragraphs(self, styled_headings_path: Path):
        doc = DocxDocument(path=styled_headings_path)
        assert len(doc.paragraphs) >= 6  # 3 headings + 3 body paragraphs


class TestAnnotationIdScanning:
    """Tests for max_annotation_id discovery."""

    def test_no_annotations_returns_zero(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        assert doc.max_annotation_id() == 0

    def test_existing_comments_finds_max_id(self, existing_comments_path: Path):
        doc = DocxDocument(path=existing_comments_path)
        max_id = doc.max_annotation_id()
        # We created comments with IDs 100, 101, 102
        assert max_id >= 102

    def test_existing_comments_has_comments_tree(self, existing_comments_path: Path):
        doc = DocxDocument(path=existing_comments_path)
        assert doc.comments_tree is not None


class TestDocumentSerialization:
    """Tests for saving and round-tripping documents."""

    def test_to_bytes_produces_valid_zip(self, simple_5para_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        data = doc.to_bytes()
        # Should be a valid ZIP
        import io
        import zipfile

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            assert "word/document.xml" in zf.namelist()

    def test_round_trip_preserves_paragraphs(self, simple_5para_path: Path):
        doc1 = DocxDocument(path=simple_5para_path)
        count1 = len(doc1.paragraphs)

        data = doc1.to_bytes()
        doc2 = DocxDocument(data=data)
        count2 = len(doc2.paragraphs)

        assert count1 == count2

    def test_save_to_file(self, simple_5para_path: Path, tmp_path: Path):
        doc = DocxDocument(path=simple_5para_path)
        out_path = tmp_path / "output.docx"
        doc.save(out_path)

        assert out_path.exists()
        assert out_path.stat().st_size > 0

        # Should be loadable
        doc2 = DocxDocument(path=out_path)
        assert len(doc2.paragraphs) == len(doc.paragraphs)

    def test_round_trip_preserves_comments(self, existing_comments_path: Path):
        doc1 = DocxDocument(path=existing_comments_path)
        max_id1 = doc1.max_annotation_id()

        data = doc1.to_bytes()
        doc2 = DocxDocument(data=data)
        max_id2 = doc2.max_annotation_id()

        assert max_id1 == max_id2

    def test_deepcopy_is_independent(self, simple_5para_path: Path):
        doc1 = DocxDocument(path=simple_5para_path)
        doc2 = doc1.deepcopy()

        # Modifying doc2 should not affect doc1
        para = doc2.paragraphs[0]
        doc2.body.remove(para)

        assert len(doc1.paragraphs) != len(doc2.paragraphs)
