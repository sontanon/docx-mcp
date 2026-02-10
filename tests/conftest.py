"""Shared pytest fixtures and configuration."""

from __future__ import annotations

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
