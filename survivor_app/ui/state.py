from __future__ import annotations

import copy
import os
import shutil

from PySide6.QtCore import QObject, Signal

from ..core.config import Config, load_config, save_config, validate
from ..core.counts import decrement, increment, load_counts, save_counts
from ..core.history import CountHistory, HistoryEntry
from ..core.models import Solution, T_Distribution
from ..core.solutions import load_solutions, save_solutions
from . import paths


class AppState(QObject):
    configChanged = Signal()
    countsChanged = Signal()
    historyChanged = Signal()
    solutionsReady = Signal(list)  # list[Solution]
    distributionEdited = Signal()

    def __init__(self):
        super().__init__()
        self.config_path = paths.config_path()
        self.counts_path = paths.counts_path()
        self.distributions_path = paths.distributions_path()

        self.config: Config = Config.empty()
        self.counts: dict[int, int] = {}
        self.history: CountHistory = CountHistory()
        self.solutions: list[Solution] = []
        self.selected_solution: Solution | None = None
        self.edited_distribution: T_Distribution | None = None

        self._load_initial()

    def _load_initial(self) -> None:
        if not os.path.exists(self.config_path):
            try:
                shutil.copyfile(paths.bundled_example_config_path(), self.config_path)
            except OSError:
                pass

        try:
            self.config = load_config(self.config_path)
        except (OSError, KeyError, ValueError):
            self.config = Config.empty()

        self.counts = load_counts(self.counts_path)
        # Solutions are deliberately not auto-loaded from distributions_path on
        # startup, even though they're auto-saved there after every solve -- they
        # can go stale against a config/counts that changed since, so loading them
        # back is an explicit action (see load_solutions_from / the Distribution
        # tab's "Load saved distributions..." button), not something that just
        # happens silently at launch.

    def config_errors(self) -> list[str]:
        return validate(self.config)

    def is_config_valid(self) -> bool:
        return not self.config_errors()

    def set_config(self, config: Config) -> None:
        self.config = config
        save_config(self.config, self.config_path)
        # A changed config may no longer match a previously-computed distribution.
        self.solutions = []
        self.selected_solution = None
        self.edited_distribution = None
        self.configChanged.emit()

    def increment_kruh(self, kruh_id: int) -> None:
        increment(kruh_id, self.counts)
        self.history.record(kruh_id)
        self._save_counts()
        self.historyChanged.emit()

    def decrement_kruh(self, kruh_id: int) -> None:
        decrement(kruh_id, self.counts)
        self._save_counts()

    def undo(self) -> HistoryEntry | None:
        entry = self.history.undo(self.counts)
        if entry is not None:
            self._save_counts()
            self.historyChanged.emit()
        return entry

    def redo(self) -> HistoryEntry | None:
        entry = self.history.redo(self.counts)
        if entry is not None:
            self._save_counts()
            self.historyChanged.emit()
        return entry

    def reset_counts(self) -> None:
        """Clear all counts back to zero and autosave, same as any other edit."""
        self.counts = {}
        self.history = CountHistory()
        self._save_counts()
        self.historyChanged.emit()

    def load_counts_from(self, filename: str) -> None:
        """Replace the current counts with those loaded from an arbitrary file, then
        autosave to the default counts file -- loading never repoints future autosaves."""
        self.counts = load_counts(filename)
        self.history = CountHistory()
        self._save_counts()
        self.historyChanged.emit()

    def save_counts_as(self, filename: str) -> None:
        """Write a snapshot of the current counts to an arbitrary file, without touching
        the default counts file that autosave keeps writing to."""
        save_counts(self.counts, filename)

    def _save_counts(self) -> None:
        save_counts(self.counts, self.counts_path)
        self.countsChanged.emit()

    def set_solutions(self, solutions: list[Solution]) -> None:
        """Set the freshly computed Solutions and persist all of them to disk --
        the solver run is expensive, so this is the only copy of that work until
        the next solve overwrites it."""
        self.solutions = solutions
        self.selected_solution = None
        self.edited_distribution = None
        save_solutions(solutions, self.distributions_path)
        self.solutionsReady.emit(solutions)

    def load_solutions_from(self, filename: str) -> None:
        """Replace the current Solutions with those loaded from an arbitrary file
        (typically the default distributions.json auto-saved after a solve), then
        re-save to that default file -- loading never repoints where a future solve
        autosaves, mirroring load_counts_from."""
        solutions = load_solutions(filename)
        self.set_solutions(solutions)

    def select_solution(self, solution: Solution) -> None:
        self.selected_solution = solution
        self.edited_distribution = copy.deepcopy(solution.distribution)
        self.distributionEdited.emit()

    def mark_distribution_edited(self) -> None:
        self.distributionEdited.emit()
