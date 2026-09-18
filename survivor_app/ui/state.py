from __future__ import annotations

import copy
import os
import shutil

from PySide6.QtCore import QObject, Signal

from ..core.config import Config, load_config, save_config, validate
from ..core.counts import decrement, increment, load_counts, save_counts
from ..core.models import Solution, T_Distribution
from . import paths


class AppState(QObject):
    configChanged = Signal()
    countsChanged = Signal()
    solutionsReady = Signal(list)  # list[Solution]
    distributionEdited = Signal()

    def __init__(self):
        super().__init__()
        self.config_path = paths.config_path()
        self.counts_path = paths.counts_path()

        self.config: Config = Config.empty()
        self.counts: dict[int, int] = {}
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
        self._save_counts()

    def decrement_kruh(self, kruh_id: int) -> None:
        decrement(kruh_id, self.counts)
        self._save_counts()

    def _save_counts(self) -> None:
        save_counts(self.counts, self.counts_path)
        self.countsChanged.emit()

    def set_solutions(self, solutions: list[Solution]) -> None:
        self.solutions = solutions
        self.selected_solution = None
        self.edited_distribution = None
        self.solutionsReady.emit(solutions)

    def select_solution(self, solution: Solution) -> None:
        self.selected_solution = solution
        self.edited_distribution = copy.deepcopy(solution.distribution)
        self.distributionEdited.emit()

    def mark_distribution_edited(self) -> None:
        self.distributionEdited.emit()
