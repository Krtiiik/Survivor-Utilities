from __future__ import annotations

from PySide6.QtWidgets import QLineEdit


class KruhIdListEditor(QLineEdit):
    """Comma-separated Kruh ids, e.g. '11, 12, 13'. Highlights red while unparsable."""

    def __init__(self, kruh_ids: list[int] | None = None):
        super().__init__()
        self.setText(", ".join(str(k) for k in (kruh_ids or [])))
        self.textChanged.connect(self._validate)
        self._validate()

    def kruh_ids(self) -> list[int]:
        try:
            return self._parse()
        except ValueError:
            return []

    def is_valid(self) -> bool:
        try:
            self._parse()
            return True
        except ValueError:
            return False

    def _parse(self) -> list[int]:
        text = self.text().strip()
        if not text:
            return []
        return [int(part.strip()) for part in text.split(",") if part.strip()]

    def _validate(self) -> None:
        self.setStyleSheet("" if self.is_valid() else "background-color: #ffcdd2;")
