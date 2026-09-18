from __future__ import annotations

from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ...core.config import ACTIVITY_TYPES, ActivityConfig


class ActivitiesEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["Activity name", "Type"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self._table)

        buttons = QHBoxLayout()
        add_button = QPushButton("Add Activity")
        add_button.clicked.connect(self._add_row)
        buttons.addWidget(add_button)
        move_up_button = QPushButton("Move Up")
        move_up_button.clicked.connect(lambda: self._move(-1))
        buttons.addWidget(move_up_button)
        move_down_button = QPushButton("Move Down")
        move_down_button.clicked.connect(lambda: self._move(1))
        buttons.addWidget(move_down_button)
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self._remove_selected)
        buttons.addWidget(remove_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(""))
        combo = QComboBox()
        combo.addItems(ACTIVITY_TYPES)
        self._table.setCellWidget(row, 1, combo)

    def _remove_selected(self) -> None:
        for index in sorted({item.row() for item in self._table.selectedItems()}, reverse=True):
            self._table.removeRow(index)

    def _move(self, direction: int) -> None:
        row = self._table.currentRow()
        target = row + direction
        if row < 0 or not (0 <= target < self._table.rowCount()):
            return

        name = self._table.item(row, 0).text()
        combo: QComboBox = self._table.cellWidget(row, 1)
        activity_type = combo.currentText()

        target_name = self._table.item(target, 0).text()
        target_combo: QComboBox = self._table.cellWidget(target, 1)
        target_type = target_combo.currentText()

        self._table.item(row, 0).setText(target_name)
        combo.setCurrentText(target_type)
        self._table.item(target, 0).setText(name)
        target_combo.setCurrentText(activity_type)
        self._table.setCurrentCell(target, 0)

    def set_activities(self, activities: list[ActivityConfig]) -> None:
        self._table.setRowCount(0)
        for activity in activities:
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QTableWidgetItem(activity.name))
            combo = QComboBox()
            combo.addItems(ACTIVITY_TYPES)
            combo.setCurrentText(activity.type)
            self._table.setCellWidget(row, 1, combo)

    def get_activities(self) -> list[ActivityConfig]:
        activities = []
        for row in range(self._table.rowCount()):
            name_item = self._table.item(row, 0)
            combo = self._table.cellWidget(row, 1)
            name = name_item.text().strip() if name_item else ""
            activity_type = combo.currentText() if isinstance(combo, QComboBox) else "rest"
            if name:
                activities.append(ActivityConfig(name=name, type=activity_type))
        return activities
