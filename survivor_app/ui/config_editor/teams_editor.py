from __future__ import annotations

from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


class TeamsEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        layout.addWidget(
            QLabel(
                "Teams names, one per Team (drag to reorder). The Distribution solver uses at "
                "most this many, fewer if that works better. The Timesheet shows the first "
                "few, as set by Teams count on the Timesheet tab."
            )
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

    def set_teams(self, teams_names: list[str]) -> None:
        self._list.clear()
        for name in teams_names:
            self._list.addItem(QListWidgetItem(name))

    def get_teams_names(self) -> list[str]:
        return [self._list.item(i).text() for i in range(self._list.count())]
