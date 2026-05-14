"""Shared pytest fixtures and configuration."""

from pathlib import Path

import pytest

FIXTURES_DIR = Path(__file__) / "fixtures" / "generated"


@pytest.fixture(scope="session", autouse=True)
def _generate_fixtures():
    """Ensure test fixture .docx files exist before any tests run."""
    from tests.generators.generate_fixtures import FIXTURES_DIR, generate_all

    # Only regenerate if the directory is empty or missing
    if not FIXTURES_DIR.exists() or not any(FIXTURES_DIR.glob("*.docx")):
        generate_all()


@pytest.fixture
def fixtures_dir() -> Path:
    from tests.generators.generate_fixtures import FIXTURES_DIR

    return FIXTURES_DIR


@pytest.fixture
def simple_5para_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "simple_5para.docx"


@pytest.fixture
def formatted_runs_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "formatted_runs.docx"


@pytest.fixture
def nda_skeleton_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "nda_skeleton.docx"


@pytest.fixture
def styled_headings_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "styled_headings.docx"


@pytest.fixture
def numbered_list_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "numbered_list.docx"


@pytest.fixture
def existing_comments_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "existing_comments.docx"


@pytest.fixture
def special_chars_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "special_chars.docx"


@pytest.fixture
def long_paragraph_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "long_paragraph.docx"


@pytest.fixture
def blank_separated_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "blank_separated.docx"


@pytest.fixture
def simple_table_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "simple_table.docx"


@pytest.fixture
def formatted_table_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "formatted_table.docx"


@pytest.fixture
def table_multi_para_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "table_multi_para.docx"


@pytest.fixture
def merged_cell_table_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "merged_cell_table.docx"


@pytest.fixture
def mixed_content_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "mixed_content.docx"


@pytest.fixture
def body_tracked_changes_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "body_tracked_changes.docx"


@pytest.fixture
def header_tracked_changes_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "header_tracked_changes.docx"


@pytest.fixture
def footer_tracked_changes_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "footer_tracked_changes.docx"


@pytest.fixture
def comments_tracked_changes_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "comments_tracked_changes.docx"


@pytest.fixture
def hyperlink_paragraph_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "hyperlink_paragraph.docx"


@pytest.fixture
def hyperlink_formatted_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "hyperlink_formatted.docx"


@pytest.fixture
def multiple_hyperlinks_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "multiple_hyperlinks.docx"


@pytest.fixture
def header_footer_text_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "header_footer_text.docx"


@pytest.fixture
def two_section_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "two_section.docx"


@pytest.fixture
def table_empty_cell_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "table_empty_cell.docx"


@pytest.fixture
def wide_table_path(fixtures_dir: Path) -> Path:
    return fixtures_dir / "wide_table.docx"
