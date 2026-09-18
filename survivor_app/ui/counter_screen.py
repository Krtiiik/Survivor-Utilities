from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGridLayout,
    QGroupBox,
    QLabel,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..core.counts import summarize
from .state import AppState


class CounterScreen(QWidget):
    """Replaces the old terminal REPL: a clickable grid of Kruh +/- buttons grouped
    by Obor, a live totals table, and autosave-on-every-click (same cadence as the
    original counter.py)."""

    def __init__(self, state: AppState):
        super().__init__()
        self._state = state
        self._count_labels: dict[int, QLabel] = {}
        self._minus_buttons: dict[int, QPushButton] = {}

        layout = QVBoxLayout(self)

        self._status_label = QLabel("")
        layout.addWidget(self._status_label)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        layout.addWidget(self._scroll, 3)

        totals_label = QLabel("Totals")
        totals_label.setFont(_bold(totals_label.font()))
        layout.addWidget(totals_label)

        self._totals_table = QTableWidget(0, 2)
        self._totals_table.setHorizontalHeaderLabels(["Kruh", "Count"])
        self._totals_table.horizontalHeader().setStretchLastSection(True)
        self._totals_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self._totals_table, 2)

        self._state.configChanged.connect(self.rebuild)
        self._state.countsChanged.connect(self._refresh_counts)

        self.rebuild()

    def rebuild(self) -> None:
        self._count_labels.clear()
        self._minus_buttons.clear()

        container = QWidget()
        container_layout = QVBoxLayout(container)

        for obor in self._state.config.obory:
            box = QGroupBox(obor.name)
            grid = QGridLayout(box)
            grid.addWidget(_bold_label("Kruh"), 0, 0)
            grid.addWidget(_bold_label("Count"), 0, 1)

            for row, kruh_id in enumerate(sorted(obor.kruhy), start=1):
                grid.addWidget(QLabel(str(kruh_id)), row, 0)

                count_label = QLabel(str(self._state.counts.get(kruh_id, 0)))
                count_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                self._count_labels[kruh_id] = count_label
                grid.addWidget(count_label, row, 1)

                minus_button = QPushButton("-")
                minus_button.setEnabled(self._state.counts.get(kruh_id, 0) > 0)
                minus_button.clicked.connect(lambda _checked=False, k=kruh_id: self._decrement(k))
                self._minus_buttons[kruh_id] = minus_button
                grid.addWidget(minus_button, row, 2)

                plus_button = QPushButton("+")
                plus_button.clicked.connect(lambda _checked=False, k=kruh_id: self._increment(k))
                grid.addWidget(plus_button, row, 3)

            container_layout.addWidget(box)

        container_layout.addStretch(1)
        self._scroll.setWidget(container)

        self._refresh_counts()

    def _increment(self, kruh_id: int) -> None:
        self._state.increment_kruh(kruh_id)
        self._flash_saved()

    def _decrement(self, kruh_id: int) -> None:
        self._state.decrement_kruh(kruh_id)
        self._flash_saved()

    def _flash_saved(self) -> None:
        self._status_label.setText("Saved")
        QTimer.singleShot(1500, lambda: self._status_label.setText(""))

    def _refresh_counts(self) -> None:
        for kruh_id, label in self._count_labels.items():
            count = self._state.counts.get(kruh_id, 0)
            label.setText(str(count))
            self._minus_buttons[kruh_id].setEnabled(count > 0)

        rows, total = summarize(self._state.counts)
        self._totals_table.setRowCount(len(rows) + 1)
        for i, (kruh_id, count) in enumerate(rows):
            self._totals_table.setItem(i, 0, QTableWidgetItem(str(kruh_id)))
            self._totals_table.setItem(i, 1, QTableWidgetItem(str(count)))

        total_label_item = QTableWidgetItem("Total")
        total_label_item.setFont(_bold(total_label_item.font()))
        total_value_item = QTableWidgetItem(str(total))
        total_value_item.setFont(_bold(total_value_item.font()))
        self._totals_table.setItem(len(rows), 0, total_label_item)
        self._totals_table.setItem(len(rows), 1, total_value_item)


def _bold(font: QFont) -> QFont:
    font.setBold(True)
    return font


def _bold_label(text: str) -> QLabel:
    label = QLabel(text)
    label.setFont(_bold(label.font()))
    return label
