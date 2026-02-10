# AGENTS.md

Instructions for AI coding agents operating in this repository.

## Project

`docx-mcp` is a Python library that applies AI-generated legal document changes
as professional tracked changes (w:ins/w:del) with comments inside `.docx` files.
It manipulates OOXML directly via `lxml` -- `python-docx` is only used in test
fixture generation.

## Build / Lint / Test

```bash
# Install dependencies
uv sync

# Run all tests
uv run pytest tests/ -v

# Run a single test file
uv run pytest tests/test_redliner.py -v

# Run a single test by name
uv run pytest tests/test_redliner.py::TestSingleModify::test_modify_produces_tracked_changes -v

# Lint
uv run ruff check src/ tests/

# Auto-fix lint
uv run ruff check src/ tests/ --fix
```

Test fixtures (`.docx` files) are auto-generated on first run by a session-scoped
pytest fixture. They live in `tests/fixtures/generated/` (gitignored).

LSP errors about unresolved imports (`lxml`, `pydantic`, `docx`, `pytest`,
`diff_match_patch`) are expected -- these packages lack type stubs but work at
runtime. Do not try to fix these.

## Code Style

### Imports

Every file starts with `from __future__ import annotations`. Imports are ordered:
future, stdlib, third-party, local -- each group separated by a blank line,
alphabetical within each group. Enforced by ruff rule `I` (isort).

```python
from __future__ import annotations

import contextlib
from pathlib import Path

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.models import Change, ChangeType, RedlineConfig
```

### Formatting

- Line length: **100** characters.
- Target: **Python 3.13+** (`py313`).
- Ruff rules enforced: `F`, `E`, `W`, `I`, `UP`, `B`, `SIM`, `ANN`, `RUF`.
- `ANN` (type annotations) is disabled for test files.

### Type Annotations

- Use modern union syntax: `str | None`, `Path | str | bytes`. Never `Optional` or `Union`.
- lxml elements are typed as `etree._Element` (the private type is the convention here).
- All public functions have full parameter and return type annotations.
- Keyword-only arguments after `*` for `id_manager` and `config` parameters.

### Naming

- `snake_case` for functions and variables.
- `PascalCase` for classes.
- `UPPER_SNAKE` for module-level constants (`W`, `NSMAP`, `_PUA_START`).
- `_private_prefix` for internal/private helpers (e.g., `_merge_adjacent`, `_convert_t_to_del_text`).

### Docstrings

Google-style with `Args:`, `Returns:`, `Raises:` sections. Use double backticks
for inline OOXML references (e.g., `` ``<w:del>`` ``). Module docstrings describe
purpose and how the module fits into the pipeline.

```python
def handle_delete(
    paragraph: etree._Element,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
) -> int:
    """Mark *paragraph* as a tracked deletion in-place.

    Args:
        paragraph: The ``<w:p>`` element to mark as deleted.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.

    Returns:
        The annotation ID used for the ``<w:del>`` wrapper.
    """
```

### Error Handling

Always assign the message to a local `msg` variable, then raise. Never inline
the string in the raise statement.

```python
msg = f"Change references fragment_id={change.fragment_id}, but document has fragments 1..{max_fid}"
raise ValueError(msg)
```

Use `contextlib.suppress(ValueError)` instead of try/except for optional int parsing.

### Data Models

- **Pydantic `BaseModel`** for validated/serializable models: `Change`, `DiffChunk`, `RedlineConfig`.
- **`@dataclass`** for plain internal data: `RunInfo`, `TaggedSegment`, `ValidationResult`.
- **`StrEnum`** for enumerations: `ChangeType`, `DiffOp`.

## Architecture

### Source Layout

```
src/docx_mcp/
  namespaces.py     OOXML namespace constants, qn(), xpath(), make_element()
  models.py         Pydantic/enum data models
  document.py       DocxDocument: ZIP ↔ XML tree manipulation
  converter.py      Paragraph XML → pseudo-Markdown
  tokenizer.py      Word-level tokenization (\S+ regex)
  differ.py         Word-level diff (diff-match-patch wrapper)
  run_ops.py        Diff ↔ XML run mapping and element building
  id_manager.py     Monotonic annotation ID allocator
  comments.py       Comment creation and range markers
  redliner.py       Main orchestrator: apply_redlines()
  validator.py      Structural validation checks
  cli.py            CLI: apply, convert, validate subcommands
  handlers/
    modify.py       Word-level tracked changes on existing paragraphs
    delete.py       Full paragraph deletion markup
    append.py       New paragraph insertion after reference
```

### Key Patterns

- All XML namespace handling goes through `qn("w", "p")` → `{namespace}p`.
- Fragment IDs are 1-based paragraph indices in `<w:body>`.
- Handler functions take positional params then keyword-only `id_manager` and `config`.
- Pseudo-Markdown format: `**bold**`, `_italic_`, `__underline__`.

## Test Conventions

- One `test_*.py` per source module. Tests grouped in classes (`class TestXxx:`).
- Shared fixtures in `conftest.py` (e.g., `simple_5para_path`, `nda_skeleton_path`).
- Per-file helper functions (`_make_paragraph()`, `_default_config()`) -- not shared.
- Use `pytest.raises(ValueError, match=...)` for error assertions.
- Use `tmp_path` fixture for file I/O tests.
- Prefix unused unpacked variables with `_` (e.g., `_new_p, ins_id = ...`).
