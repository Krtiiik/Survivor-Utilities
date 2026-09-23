from __future__ import annotations

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ...core.config import Config, validate
from ..state import AppState
from .activities_editor import ActivitiesEditor
from .obory_editor import OboryEditor
from .search_space_editor import SearchSpaceEditor
from .subteams_editor import SubteamsEditor
from .teams_editor import TeamsEditor
from .time_editor import TimeEditor


class ConfigEditorScreen(QWidget):
    """Form-based editor for config.json, so the successor never has to hand-edit
    JSON. Edits are held in a working copy and only validated + written to disk
    when "Save Config" is pressed -- never persisting a config that fails
    validate()."""

    def __init__(self, state: AppState):
        super().__init__()
        self._state = state

        layout = QVBoxLayout(self)

        tabs = QTabWidget()
        self._obory_editor = OboryEditor()
        tabs.addTab(self._obory_editor, "Obory")
        self._teams_editor = TeamsEditor()
        tabs.addTab(self._teams_editor, "Teams")
        self._subteams_editor = SubteamsEditor()
        tabs.addTab(self._subteams_editor, "Subteams")
        self._activities_editor = ActivitiesEditor()
        tabs.addTab(self._activities_editor, "Activities")
        self._time_editor = TimeEditor()
        tabs.addTab(self._time_editor, "Time")
        self._search_space_editor = SearchSpaceEditor()
        tabs.addTab(self._search_space_editor, "Solver search space")
        layout.addWidget(tabs, 1)

        self._errors_label = QLabel("")
        self._errors_label.setStyleSheet("color: #b91c1c;")
        self._errors_label.setWordWrap(True)
        layout.addWidget(self._errors_label)

        buttons = QHBoxLayout()
        save_button = QPushButton("Save Config")
        save_button.clicked.connect(self._save)
        buttons.addWidget(save_button)
        reload_button = QPushButton("Reload")
        reload_button.setToolTip("Discard unsaved edits and reload the last saved config.")
        reload_button.clicked.connect(self._reload)
        buttons.addWidget(reload_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        self._reload()

    def _reload(self) -> None:
        config = self._state.config
        self._obory_editor.set_obory(config.obory)
        self._teams_editor.set_teams(config.teams_names)
        self._subteams_editor.set_subteams(config.subteams)
        self._activities_editor.set_activities(config.activities)
        self._time_editor.set_time(config.time)
        self._search_space_editor.set_search_space(config.possible_teams_sizes)
        self._search_space_editor.set_solver_time_limit(config.solver_time_limit)
        self._search_space_editor.set_min_split_part_size(config.min_split_part_size)
        self._errors_label.setText("")

    def _build_config(self) -> Config:
        return Config(
            possible_teams_sizes=self._search_space_editor.get_possible_teams_sizes(),
            teams_names=self._teams_editor.get_teams_names(),
            subteams=self._subteams_editor.get_subteams(),
            activities=self._activities_editor.get_activities(),
            time=self._time_editor.get_time(),
            obory=self._obory_editor.get_obory(),
            min_split_part_size=self._search_space_editor.get_min_split_part_size(),
            solver_time_limit=self._search_space_editor.get_solver_time_limit(),
        )

    def _save(self) -> None:
        config = self._build_config()
        errors = validate(config)
        if errors:
            self._errors_label.setText("Cannot save -- fix the following first:\n- " + "\n- ".join(errors))
            return

        self._errors_label.setText("")
        self._state.set_config(config)
        QMessageBox.information(self, "Config saved", "Configuration saved successfully.")
