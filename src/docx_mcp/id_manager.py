"""Globally unique annotation-ID allocator.

Every tracked change (``w:ins``, ``w:del``), comment, and bookmark in OOXML
must carry a unique ``w:id`` integer attribute.  :class:`IdManager` hands out
monotonically increasing IDs starting from one past the document's current
maximum, ensuring no collisions.
"""


class IdManager:
    """Allocates globally unique annotation IDs.

    Usage::

        mgr = IdManager(start_after=doc.max_annotation_id())
        ins_id = mgr.next_id()   # e.g. 103
        del_id = mgr.next_id()   # 104
        comment_id = mgr.next_id()  # 105
    """

    def __init__(self, start_after: int = 0) -> None:
        """Create an ID manager.

        Args:
            start_after: The highest existing annotation ID in the document.
                New IDs will begin at ``start_after + 1``.
        """
        self._next = start_after + 1

    def next_id(self) -> int:
        """Return the next available ID and advance the counter."""
        current = self._next
        self._next += 1
        return current

    @property
    def peek(self) -> int:
        """Return the next ID that will be allocated (without consuming it)."""
        return self._next
