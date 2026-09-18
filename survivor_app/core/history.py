from __future__ import annotations

from dataclasses import dataclass

from .counts import decrement, increment


@dataclass(frozen=True)
class HistoryEntry:
    kruh_id: int


class CountHistory:
    """Tracks a session's sequence of Kruh increments, supporting undo/redo.

    Only increments are tracked (per the Counter tab's manual +/- and keyboard-entry
    actions) -- manual decrements are corrections and stay outside the undo stack.
    """

    def __init__(self) -> None:
        self._done: list[HistoryEntry] = []
        self._undone: list[HistoryEntry] = []

    @property
    def entries(self) -> list[HistoryEntry]:
        """Increments applied so far, oldest first."""
        return list(self._done)

    def can_undo(self) -> bool:
        return bool(self._done)

    def can_redo(self) -> bool:
        return bool(self._undone)

    def record(self, kruh_id: int) -> None:
        """Record an increment that has already been applied to `counts`."""
        self._done.append(HistoryEntry(kruh_id))
        self._undone.clear()

    def undo(self, counts: dict[int, int]) -> HistoryEntry | None:
        """Revert the most recent recorded increment. Returns the reverted entry, if any."""
        if not self._done:
            return None
        entry = self._done.pop()
        decrement(entry.kruh_id, counts)
        self._undone.append(entry)
        return entry

    def redo(self, counts: dict[int, int]) -> HistoryEntry | None:
        """Reapply the most recently undone increment. Returns the reapplied entry, if any."""
        if not self._undone:
            return None
        entry = self._undone.pop()
        increment(entry.kruh_id, counts)
        self._done.append(entry)
        return entry
