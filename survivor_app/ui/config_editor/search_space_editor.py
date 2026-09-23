from __future__ import annotations

from PySide6.QtWidgets import QFormLayout, QLabel, QLineEdit, QSpinBox, QWidget


class SearchSpaceEditor(QWidget):
    """Edits the solver's search space. Every Possible Teams size gets tried, each
    taking up to the solver time limit, so together they control how long a
    solver run can take."""

    def __init__(self):
        super().__init__()
        layout = QFormLayout(self)

        self._sizes_edit = QLineEdit()
        layout.addRow("Possible Teams sizes (comma-separated):", self._sizes_edit)

        self._time_limit_spin = QSpinBox()
        self._time_limit_spin.setRange(1, 3600)
        self._time_limit_spin.setSuffix(" s")
        self._time_limit_spin.setToolTip(
            "How long the solver may search for each Possible Teams size. Longer "
            "usually gives more even Teams and Subteams; when time runs out, the "
            "best distribution found so far is used."
        )
        layout.addRow("Solver time limit per size:", self._time_limit_spin)

        self._estimate_label = QLabel("")
        layout.addRow("", self._estimate_label)

        self._min_split_part_spin = QSpinBox()
        self._min_split_part_spin.setRange(1, 99)
        self._min_split_part_spin.setToolTip(
            "A Kruh too large for one Subteam is spread across several Subteams of the "
            "same Team. This is the fewest people each part may have (lowered "
            "automatically if a Kruh can't be split that evenly)."
        )
        layout.addRow("Min people per split Kruh part:", self._min_split_part_spin)

        self._sizes_edit.textChanged.connect(self._update_estimate)
        self._time_limit_spin.valueChanged.connect(self._update_estimate)

    def set_search_space(self, sizes: list[int]) -> None:
        self._sizes_edit.setText(", ".join(str(s) for s in sizes))
        self._update_estimate()

    def set_solver_time_limit(self, seconds: int) -> None:
        self._time_limit_spin.setValue(seconds)

    def get_solver_time_limit(self) -> int:
        return self._time_limit_spin.value()

    def set_min_split_part_size(self, size: int) -> None:
        self._min_split_part_spin.setValue(size)

    def get_min_split_part_size(self) -> int:
        return self._min_split_part_spin.value()

    def get_possible_teams_sizes(self) -> list[int]:
        return _parse_ints(self._sizes_edit.text())

    def _update_estimate(self) -> None:
        combos = len(self.get_possible_teams_sizes())
        worst_case_minutes = combos * self.get_solver_time_limit() / 60
        self._estimate_label.setText(
            f"This runs {combos} solve(s), up to ~{worst_case_minutes:.1f} min worst-case."
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
