from __future__ import annotations

from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ...core.config import SubteamConfig
from ..widgets.color_picker import ColorPickerButton


class SubteamsEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["Subteam name", "Color"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self._table)

        buttons = QHBoxLayout()
        add_button = QPushButton("Add Subteam")
        add_button.clicked.connect(self._add_row)
        buttons.addWidget(add_button)
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self._remove_selected)
        buttons.addWidget(remove_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

    def _add_row(self) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(str(row + 1)))
        self._table.setCellWidget(row, 1, ColorPickerButton("#ffffff"))

    def _remove_selected(self) -> None:
        for index in sorted({item.row() for item in self._table.selectedItems()}, reverse=True):
            self._table.removeRow(index)

    def set_subteams(self, subteams: list[SubteamConfig]) -> None:
        self._table.setRowCount(0)
        for subteam in subteams:
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QTableWidgetItem(subteam.name))
            self._table.setCellWidget(row, 1, ColorPickerButton(subteam.color))

    def get_subteams(self) -> list[SubteamConfig]:
        subteams = []
        for row in range(self._table.rowCount()):
            name_item = self._table.item(row, 0)
            color_widget = self._table.cellWidget(row, 1)
            name = name_item.text().strip() if name_item else ""
            color = color_widget.color() if isinstance(color_widget, ColorPickerButton) else "#ffffff"
            if name:
                subteams.append(SubteamConfig(name=name, color=color))
        return subteams
