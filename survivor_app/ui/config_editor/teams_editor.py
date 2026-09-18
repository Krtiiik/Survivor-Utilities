from __future__ import annotations

from PySide6.QtWidgets import (
    QAbstractItemView,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)


class TeamsEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        form = QFormLayout()
        self._teams_count_spin = QSpinBox()
        self._teams_count_spin.setRange(0, 999)
        form.addRow("Teams count (used by the Timesheet):", self._teams_count_spin)
        layout.addLayout(form)

        layout.addWidget(
            QLabel("Teams names (shared by the Timesheet and the Distribution solver; drag to reorder):")
        )
        self._list = QListWidget()
        self._list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        layout.addWidget(self._list)

        buttons = QHBoxLayout()
        add_button = QPushButton("Add")
        add_button.clicked.connect(self._add)
        buttons.addWidget(add_button)
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self._remove_selected)
        buttons.addWidget(remove_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

    def _add(self) -> None:
        name, ok = QInputDialog.getText(self, "Add team", "Team name:")
        if ok and name.strip():
            self._list.addItem(QListWidgetItem(name.strip()))

    def _remove_selected(self) -> None:
        for item in self._list.selectedItems():
            self._list.takeItem(self._list.row(item))

    def set_teams(self, teams_count: int, teams_names: list[str]) -> None:
        self._teams_count_spin.setValue(teams_count)
        self._list.clear()
        for name in teams_names:
            self._list.addItem(QListWidgetItem(name))

    def get_teams_count(self) -> int:
        return self._teams_count_spin.value()

    def get_teams_names(self) -> list[str]:
        return [self._list.item(i).text() for i in range(self._list.count())]
