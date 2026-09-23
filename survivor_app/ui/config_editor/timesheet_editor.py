from __future__ import annotations

from PySide6.QtCore import QTime
from PySide6.QtWidgets import QFormLayout, QLabel, QSpinBox, QTimeEdit, QVBoxLayout, QWidget

from ...core.config import ActivityConfig, TimeConfig
from .activities_editor import ActivitiesEditor


class TimesheetEditor(QWidget):
    """Everything that only affects the Timesheet: how many Teams it renders,
    event timing, and the Activities in schedule order."""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        form = QFormLayout()
        self._teams_count_spin = QSpinBox()
        self._teams_count_spin.setRange(1, 999)
        self._teams_count_spin.setToolTip(
            "How many Teams the Timesheet shows (the first names from the Teams list). "
            "Must not exceed the number of Teams names or Activities. With fewer Teams "
            "than Activities, some Activities have no Team in some time slots (shown as ∅)."
        )
        form.addRow("Teams count:", self._teams_count_spin)

        self._start_edit = QTimeEdit()
        self._start_edit.setDisplayFormat("HH:mm")
        form.addRow("Event start:", self._start_edit)

        self._duration_edit = QTimeEdit()
        self._duration_edit.setDisplayFormat("HH:mm")
        form.addRow("Activity duration (HH:MM):", self._duration_edit)
        layout.addLayout(form)

        layout.addWidget(QLabel("Activities, in schedule order:"))
        self._activities_editor = ActivitiesEditor()
        layout.addWidget(self._activities_editor, 1)

    def set_timesheet(
        self, teams_count: int, time_config: TimeConfig, activities: list[ActivityConfig]
    ) -> None:
        self._teams_count_spin.setValue(teams_count)
        self._start_edit.setTime(_parse(time_config.start))
        self._duration_edit.setTime(_parse(time_config.activity_duration))
        self._activities_editor.set_activities(activities)

    def get_teams_count(self) -> int:
        return self._teams_count_spin.value()

    def get_time(self) -> TimeConfig:
        return TimeConfig(
            start=self._start_edit.time().toString("HH:mm"),
            activity_duration=self._duration_edit.time().toString("HH:mm"),
        )

    def get_activities(self) -> list[ActivityConfig]:
        return self._activities_editor.get_activities()


def _parse(value: str) -> QTime:
    time = QTime.fromString(value, "HH:mm")
    return time if time.isValid() else QTime(0, 0)
