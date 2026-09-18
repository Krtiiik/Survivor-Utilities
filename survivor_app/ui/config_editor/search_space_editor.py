from __future__ import annotations

from PySide6.QtWidgets import QFormLayout, QLabel, QLineEdit, QWidget

from ...core.distribute import SOLVER_TIME_LIMIT


class SearchSpaceEditor(QWidget):
    """Edits the solver's search space. Every (Possible Teams count) x (Possible
    Teams size) combination gets tried, each taking up to SOLVER_TIME_LIMIT
    seconds, so this directly controls how long a solver run can take."""

    def __init__(self):
        super().__init__()
        layout = QFormLayout(self)

        self._counts_edit = QLineEdit()
        layout.addRow("Possible Teams counts (comma-separated):", self._counts_edit)

        self._sizes_edit = QLineEdit()
        layout.addRow("Possible Teams sizes (comma-separated):", self._sizes_edit)

        self._estimate_label = QLabel("")
        layout.addRow("", self._estimate_label)

        self._counts_edit.textChanged.connect(self._update_estimate)
        self._sizes_edit.textChanged.connect(self._update_estimate)

    def set_search_space(self, counts: list[int], sizes: list[int]) -> None:
        self._counts_edit.setText(", ".join(str(c) for c in counts))
        self._sizes_edit.setText(", ".join(str(s) for s in sizes))
        self._update_estimate()

    def get_possible_teams_counts(self) -> list[int]:
        return _parse_ints(self._counts_edit.text())

    def get_possible_teams_sizes(self) -> list[int]:
        return _parse_ints(self._sizes_edit.text())

    def _update_estimate(self) -> None:
        combos = len(self.get_possible_teams_counts()) * len(self.get_possible_teams_sizes())
        worst_case_minutes = combos * SOLVER_TIME_LIMIT / 60
        self._estimate_label.setText(
            f"This runs {combos} combination(s), up to ~{worst_case_minutes:.1f} min worst-case."
        )


def _parse_ints(text: str) -> list[int]:
    result = []
    for part in text.split(","):
        part = part.strip()
        if part:
            try:
                result.append(int(part))
            except ValueError:
                continue
    return result
