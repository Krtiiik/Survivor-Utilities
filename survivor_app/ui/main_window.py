from __future__ import annotations

from PySide6.QtWidgets import QMainWindow, QTabWidget

from .config_editor.config_editor_screen import ConfigEditorScreen
from .counter_screen import CounterScreen
from .distribution_screen import DistributionScreen
from .state import AppState
from .timesheet_screen import TimesheetScreen


class MainWindow(QMainWindow):
    """Tabbed navigation, not a wizard: the real workflow isn't strictly linear
    (recounting mid-event, regenerating the timesheet independently of the
    distribution, etc.), so all four steps stay reachable at once. Distribution
    and Timesheet are disabled until the Config passes validation."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Survivor")
        self.resize(1100, 750)

        self.state = AppState()

        self._tabs = QTabWidget()
        self.setCentralWidget(self._tabs)

        self._tabs.addTab(CounterScreen(self.state), "1. Counter")
        self._tabs.addTab(DistributionScreen(self.state), "2. Distribution")
        self._tabs.addTab(TimesheetScreen(self.state), "3. Timesheet")
        self._tabs.addTab(ConfigEditorScreen(self.state), "4. Config")

        self.state.configChanged.connect(self._update_tab_availability)
        self._update_tab_availability()

    def _update_tab_availability(self) -> None:
        valid = self.state.is_config_valid()
        self._tabs.setTabEnabled(1, valid)
        self._tabs.setTabEnabled(2, valid)
        if not valid:
            self.statusBar().showMessage(
                "Config is invalid -- fix it on the Config tab before using Distribution/Timesheet.", 5000
            )
