"""Word-level tokenization for diffing pseudo-Markdown text.

Splits text into word tokens using whitespace boundaries. Each token is a
contiguous run of non-whitespace characters (``\\S+``), preserving formatting
markers (``**``, ``_``, ``__``) as part of their adjacent word tokens.

Non-breaking spaces (``\\u00a0``) are treated as whitespace by ``\\S+`` and
thus act as word boundaries, which is the correct behaviour for diffing.
"""

from __future__ import annotations

import re

_WORD_RE = re.compile(r"\S+")


def tokenize(text: str) -> list[str]:
    """Split *text* into word-level tokens.

    Returns a list of non-whitespace tokens.  An empty or whitespace-only
    string returns an empty list.
    """
    return _WORD_RE.findall(text)


def detokenize(tokens: list[str]) -> str:
    """Join *tokens* back into a single string separated by spaces."""
    return " ".join(tokens)
