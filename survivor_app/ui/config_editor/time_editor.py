from __future__ import annotations

from PySide6.QtCore import QTime
from PySide6.QtWidgets import QFormLayout, QTimeEdit, QWidget

from ...core.config import TimeConfig


class TimeEditor(QWidget):
    def __init__(self):
        super().__init__()
        layout = QFormLayout(self)

        self._start_edit = QTimeEdit()
        self._start_edit.setDisplayFormat("HH:mm")
        layout.addRow("Event start:", self._start_edit)

        self._duration_edit = QTimeEdit()
        self._duration_edit.setDisplayFormat("HH:mm")
        layout.addRow("Activity duration (HH:MM):", self._duration_edit)

    def set_time(self, time_config: TimeConfig) -> None:
        self._start_edit.setTime(_parse(time_config.start))
        self._duration_edit.setTime(_parse(time_config.activity_duration))

    def get_time(self) -> TimeConfig:
        return TimeConfig(
            start=self._start_edit.time().toString("HH:mm"),
            activity_duration=self._duration_edit.time().toString("HH:mm"),
        )


def _parse(value: str) -> QTime:
    time = QTime.fromString(value, "HH:mm")
    return time if time.isValid() else QTime(0, 0)
