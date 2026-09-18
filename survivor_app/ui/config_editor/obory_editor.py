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

from ...core.config import OborConfig
from ..widgets.kruh_id_list_editor import KruhIdListEditor


class OboryEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["Obor name", "Kruh ids (comma-separated)"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self._table)

        buttons = QHBoxLayout()
        add_button = QPushButton("Add Obor")
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
        self._table.setItem(row, 0, QTableWidgetItem(""))
        self._table.setCellWidget(row, 1, KruhIdListEditor([]))

    def _remove_selected(self) -> None:
        for index in sorted({item.row() for item in self._table.selectedItems()}, reverse=True):
            self._table.removeRow(index)

    def set_obory(self, obory: list[OborConfig]) -> None:
        self._table.setRowCount(0)
        for obor in obory:
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QTableWidgetItem(obor.name))
            self._table.setCellWidget(row, 1, KruhIdListEditor(obor.kruhy))

    def get_obory(self) -> list[OborConfig]:
        obory = []
        for row in range(self._table.rowCount()):
            name_item = self._table.item(row, 0)
            editor = self._table.cellWidget(row, 1)
            name = name_item.text().strip() if name_item else ""
            kruhy = editor.kruh_ids() if isinstance(editor, KruhIdListEditor) else []
            if name or kruhy:
                obory.append(OborConfig(name=name, kruhy=kruhy))
        return obory
