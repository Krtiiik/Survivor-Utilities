from __future__ import annotations

from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ..core.errors import FileWriteError
from ..core.excel_export import render_timetable_xlsx
from ..core.timesheet import compute_timetable_layout
from .state import AppState
from .timesheet_preview import TimesheetPreview


class TimesheetScreen(QWidget):
    def __init__(self, state: AppState):
        super().__init__()
        self._state = state

        layout = QVBoxLayout(self)

        controls = QHBoxLayout()
        self._export_button = QPushButton("Export to .xlsx...")
        self._export_button.clicked.connect(self._export)
        controls.addWidget(self._export_button)
        controls.addStretch(1)
        layout.addLayout(controls)

        self._error_label = QLabel("")
        self._error_label.setStyleSheet("color: #b91c1c;")
        self._error_label.setWordWrap(True)
        layout.addWidget(self._error_label)

        self._preview = TimesheetPreview()
        layout.addWidget(self._preview, 1)

        self._state.configChanged.connect(self.refresh)
        self.refresh()

    def refresh(self) -> None:
        if not self._state.is_config_valid():
            self._error_label.setText("Config is invalid -- fix it on the Config tab before previewing.")
            self._export_button.setEnabled(False)
            return

        try:
            layout = compute_timetable_layout(self._state.config)
        except ValueError as error:
            self._error_label.setText(str(error))
            self._export_button.setEnabled(False)
            return

        self._error_label.setText("")
        self._export_button.setEnabled(True)
        self._preview.load_layout(layout)

    def _export(self) -> None:
        filename, _ = QFileDialog.getSaveFileName(
            self, "Export timesheet", "timesheet.xlsx", "Excel Workbook (*.xlsx)"
        )
        if not filename:
            return

        layout = compute_timetable_layout(self._state.config)
        try:
            render_timetable_xlsx(layout, self._state.config, filename)
        except FileWriteError as error:
            QMessageBox.critical(self, "Export failed", str(error))
        else:
            QMessageBox.information(self, "Export complete", f"Timesheet exported to {filename}")
