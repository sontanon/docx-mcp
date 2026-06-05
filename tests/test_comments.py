"""Tests for the comments module."""

from docx_mcp.comments import _initials, add_comment, ensure_comments_part
from docx_mcp.document import DocxDocument
from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig
from docx_mcp.namespaces import CT_COMMENTS, REL_COMMENTS, qn, xpath

# ===================================================================
# Helpers
# ===================================================================


def _default_config() -> RedlineConfig:
    return RedlineConfig(author="AI Review")


# ===================================================================
# ensure_comments_part tests
# ===================================================================


class TestEnsureCommentsPart:
    def test_creates_comments_tree_if_missing(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        assert doc.comments_tree is None

        root = ensure_comments_part(doc)
        assert root is not None
        assert doc.comments_tree is root
        assert root.tag == qn("w", "comments")

    def test_returns_existing_tree(self, existing_comments_path):
        doc = DocxDocument(path=existing_comments_path)
        original = doc.comments_tree
        assert original is not None

        result = ensure_comments_part(doc)
        assert result is original  # Same object

    def test_adds_relationship(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        ensure_comments_part(doc)

        # Check rels tree has a comments relationship
        rels = doc.rels_tree
        assert rels is not None
        found = False
        for child in rels:
            if isinstance(child.tag, str) and child.get("Type") == REL_COMMENTS:
                found = True
                assert child.get("Target") == "comments.xml"
                break
        assert found, "Comments relationship not found in rels"

    def test_adds_content_type(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        ensure_comments_part(doc)

        ct = doc.content_types_tree
        assert ct is not None
        found = False
        for child in ct:
            if isinstance(child.tag, str) and child.get("PartName") == "/word/comments.xml":
                assert child.get("ContentType") == CT_COMMENTS
                found = True
                break
        assert found, "Comments content type not found"

    def test_idempotent(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        ensure_comments_part(doc)
        ensure_comments_part(doc)  # Call again

        # Should only have one relationship
        rels = doc.rels_tree
        assert rels is not None
        count = sum(
            1 for child in rels if isinstance(child.tag, str) and child.get("Type") == REL_COMMENTS
        )
        assert count == 1


# ===================================================================
# add_comment tests
# ===================================================================


class TestAddComment:
    def test_creates_comment_element(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[1]
        mgr = IdManager(start_after=doc.max_annotation_id())

        cid = add_comment(
            doc,
            para,
            "Changed for clarity.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Comment should exist in comments.xml
        comments_root = doc.comments_tree
        assert comments_root is not None
        comments = xpath(comments_root, "w:comment")
        assert len(comments) >= 1

        # Find our comment
        found = None
        for c in comments:
            if c.get(qn("w", "id")) == str(cid):
                found = c
                break
        assert found is not None
        assert found.get(qn("w", "author")) == "AI Review"

        # Check justification text
        t_elements = found.iter(qn("w", "t"))
        texts = [t.text for t in t_elements if t.text]
        assert "Changed for clarity." in texts

    def test_inserts_range_markers(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[1]
        mgr = IdManager(start_after=doc.max_annotation_id())

        cid = add_comment(
            doc,
            para,
            "Test comment.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should have commentRangeStart and commentRangeEnd
        crs = xpath(para, "w:commentRangeStart")
        cre = xpath(para, "w:commentRangeEnd")
        assert len(crs) >= 1
        assert len(cre) >= 1

        # Check IDs match
        assert crs[0].get(qn("w", "id")) == str(cid)
        assert cre[0].get(qn("w", "id")) == str(cid)

    def test_inserts_comment_reference(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[1]
        mgr = IdManager(start_after=doc.max_annotation_id())

        cid = add_comment(
            doc,
            para,
            "Ref test.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should have a commentReference run
        refs = list(para.iter(qn("w", "commentReference")))
        assert len(refs) >= 1
        assert refs[0].get(qn("w", "id")) == str(cid)

    def test_comment_with_range_elements(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        para = fmap[1]
        mgr = IdManager(start_after=doc.max_annotation_id())

        # Get runs to use as range
        runs = xpath(para, "w:r")
        if runs:
            cid = add_comment(
                doc,
                para,
                "Range comment.",
                id_manager=mgr,
                config=_default_config(),
                range_elements=[runs[0]],
            )

            # commentRangeStart should be before the first run
            crs = xpath(para, "w:commentRangeStart")
            assert len(crs) >= 1
            assert crs[0].get(qn("w", "id")) == str(cid)

    def test_preserves_existing_comments(self, existing_comments_path):
        doc = DocxDocument(path=existing_comments_path)
        fmap = doc.fragment_map()
        para = fmap[1]

        original_comments = len(xpath(doc.comments_tree, "w:comment"))
        mgr = IdManager(start_after=doc.max_annotation_id())

        add_comment(
            doc,
            para,
            "New comment.",
            id_manager=mgr,
            config=_default_config(),
        )

        new_comments = len(xpath(doc.comments_tree, "w:comment"))
        assert new_comments == original_comments + 1

    def test_document_saves_with_comments(self, simple_5para_path):
        doc = DocxDocument(path=simple_5para_path)
        fmap = doc.fragment_map()
        mgr = IdManager(start_after=doc.max_annotation_id())

        add_comment(
            doc,
            fmap[1],
            "Comment 1.",
            id_manager=mgr,
            config=_default_config(),
        )
        add_comment(
            doc,
            fmap[2],
            "Comment 2.",
            id_manager=mgr,
            config=_default_config(),
        )

        # Should save without error
        data = doc.to_bytes()
        assert len(data) > 0

        # Reload and verify
        doc2 = DocxDocument(data=data)
        assert doc2.comments_tree is not None
        comments = xpath(doc2.comments_tree, "w:comment")
        assert len(comments) == 2


# ===================================================================
# _initials tests
# ===================================================================


class TestInitials:
    def test_two_word_name(self):
        assert _initials("AI Review") == "AR"

    def test_single_word(self):
        assert _initials("Bot") == "B"

    def test_three_words(self):
        assert _initials("John Q Smith") == "JQS"

    def test_empty_string(self):
        assert _initials("") == ""
