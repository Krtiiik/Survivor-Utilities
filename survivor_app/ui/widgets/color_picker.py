from __future__ import annotations

from PySide6.QtGui import QColor
from PySide6.QtWidgets import QColorDialog, QPushButton


class ColorPickerButton(QPushButton):
    """A button showing its current color as a swatch; click opens QColorDialog."""

    def __init__(self, initial_color: str = "#ffffff"):
        super().__init__()
        self._color = initial_color
        self.setFixedWidth(80)
        self.clicked.connect(self._pick)
        self._refresh()

    def color(self) -> str:
        return self._color

    def set_color(self, color: str) -> None:
        self._color = color
        self._refresh()

    def _pick(self) -> None:
        chosen = QColorDialog.getColor(QColor(self._color), self, "Choose color")
        if chosen.isValid():
            self.set_color(chosen.name())

    def _refresh(self) -> None:
        self.setText(self._color)
        self.setStyleSheet(f"background-color: {self._color};")
