"""Word-level diff engine using diff-match-patch.

Implements the "word → unique-char" mapping technique:

1. Tokenize both texts into word lists.
2. Build a bijection from unique word strings to unique Unicode code points.
3. Convert each word list into a *char string* (one char per word).
4. Run ``diff_match_patch.diff_main`` on the two char strings.
5. Apply ``diff_cleanupSemantic`` for human-readable boundaries.
6. Map the char-level diff back to word-level ``DiffChunk`` objects.

This gives us word-granularity diffs while leveraging DMP's highly optimised
Myers algorithm and semantic cleanup.
"""

from __future__ import annotations

from typing import Protocol

import diff_match_patch as dmp_module

from docx_mcp.models import DiffChunk, DiffOp
from docx_mcp.tokenizer import detokenize, tokenize

# ---------------------------------------------------------------------------
# Public protocol (allows swapping implementation later)
# ---------------------------------------------------------------------------


class Differ(Protocol):
    """Interface for a word-level diff engine."""

    def diff(self, old: str, new: str) -> list[DiffChunk]:
        """Return a list of diff chunks describing changes from *old* to *new*."""
        ...


# ---------------------------------------------------------------------------
# Helpers: word ↔ char mapping
# ---------------------------------------------------------------------------

# We start mapping at a Private Use Area code point to avoid collisions
# with any real text characters.  PUA-A (U+F0000-U+FFFFD) gives us 65534
# slots — more than enough for any legal document.
_PUA_START = 0xF0000


def _words_to_chars(
    words_a: list[str],
    words_b: list[str],
) -> tuple[str, str, list[str]]:
    """Map word tokens to unique single characters.

    Returns ``(chars_a, chars_b, word_list)`` where ``word_list[i]``
    is the word that character ``chr(_PUA_START + i)`` represents.
    """
    word_to_char: dict[str, str] = {}
    word_list: list[str] = []

    def _encode(words: list[str]) -> str:
        chars: list[str] = []
        for w in words:
            if w not in word_to_char:
                idx = len(word_list)
                c = chr(_PUA_START + idx)
                word_to_char[w] = c
                word_list.append(w)
            chars.append(word_to_char[w])
        return "".join(chars)

    return _encode(words_a), _encode(words_b), word_list


_DMP_OP_MAP: dict[int, DiffOp] = {
    -1: DiffOp.DELETE,
    0: DiffOp.EQUAL,
    1: DiffOp.INSERT,
}


# ---------------------------------------------------------------------------
# Main implementation
# ---------------------------------------------------------------------------


class DmpWordDiffer:
    """Word-level differ powered by ``diff-match-patch``."""

    def diff(self, old: str, new: str) -> list[DiffChunk]:
        """Compute a word-level diff between *old* and *new*.

        Both inputs are plain or pseudo-Markdown strings.  Returns a list
        of :class:`DiffChunk` objects whose ``text`` fields, when
        concatenated (equal + insert), reproduce *new*, and (equal + delete)
        reproduce *old* (modulo whitespace normalisation between words).
        """
        old_words = tokenize(old)
        new_words = tokenize(new)

        # Fast path: identical texts
        if old_words == new_words:
            if old_words:
                return [DiffChunk(op=DiffOp.EQUAL, text=detokenize(old_words))]
            return []

        # Fast path: one side empty
        if not old_words:
            return [DiffChunk(op=DiffOp.INSERT, text=detokenize(new_words))]
        if not new_words:
            return [DiffChunk(op=DiffOp.DELETE, text=detokenize(old_words))]

        # Map words → chars, diff chars, map back
        chars_a, chars_b, word_list = _words_to_chars(old_words, new_words)

        dmp = dmp_module.diff_match_patch()
        char_diffs: list[tuple[int, str]] = dmp.diff_main(chars_a, chars_b, checklines=False)
        dmp.diff_cleanupSemantic(char_diffs)

        chunks: list[DiffChunk] = []
        for op_int, char_text in char_diffs:
            op = _DMP_OP_MAP[op_int]
            words = [word_list[ord(c) - _PUA_START] for c in char_text]
            text = detokenize(words)
            if text:  # skip empty chunks (shouldn't happen, but be safe)
                chunks.append(DiffChunk(op=op, text=text))

        return _merge_adjacent(chunks)


def _merge_adjacent(chunks: list[DiffChunk]) -> list[DiffChunk]:
    """Merge consecutive chunks that share the same operation."""
    if not chunks:
        return chunks
    merged: list[DiffChunk] = [chunks[0]]
    for chunk in chunks[1:]:
        if chunk.op == merged[-1].op:
            merged[-1] = DiffChunk(op=chunk.op, text=f"{merged[-1].text} {chunk.text}")
        else:
            merged.append(chunk)
    return merged
