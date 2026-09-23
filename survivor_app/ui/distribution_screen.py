from __future__ import annotations

from PySide6.QtCore import QThread, Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from ..core.errors import FileWriteError
from ..core.excel_export import export_distribution
from ..core.models import ProgressEvent, Solution, SolutionStatus
from .distribution_grid import DistributionGrid
from .state import AppState
from .workers import DistributionWorker

SOLUTION_ROLE = Qt.ItemDataRole.UserRole


class DistributionScreen(QWidget):
    def __init__(self, state: AppState):
        super().__init__()
        self._state = state
        self._thread: QThread | None = None
        self._worker: DistributionWorker | None = None

        layout = QVBoxLayout(self)

        self._search_space_label = QLabel("")
        layout.addWidget(self._search_space_label)

        controls = QHBoxLayout()
        self._run_button = QPushButton("Run Solver")
        self._run_button.clicked.connect(self._start_run)
        controls.addWidget(self._run_button)

        self._cancel_button = QPushButton("Cancel")
        self._cancel_button.setEnabled(False)
        self._cancel_button.clicked.connect(self._cancel_run)
        controls.addWidget(self._cancel_button)

        self._load_button = QPushButton("Load saved distributions...")
        self._load_button.clicked.connect(self._load_saved)
        controls.addWidget(self._load_button)

        self._export_button = QPushButton("Export to .xlsx...")
        self._export_button.setEnabled(False)
        self._export_button.clicked.connect(self._export)
        controls.addWidget(self._export_button)
        controls.addStretch(1)
        layout.addLayout(controls)

        self._progress_bar = QProgressBar()
        layout.addWidget(self._progress_bar)

        self._log = QPlainTextEdit()
        self._log.setReadOnly(True)
        self._log.setMaximumHeight(120)
        layout.addWidget(self._log)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self._results_list = QListWidget()
        self._results_list.itemSelectionChanged.connect(self._on_result_selected)
        splitter.addWidget(self._results_list)

        self._grid = DistributionGrid()
        self._grid.edited.connect(self._on_grid_edited)
        splitter.addWidget(self._grid)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        layout.addWidget(splitter, 1)

        self._banner = QLabel("")
        self._banner.setStyleSheet("color: #b45309;")
        self._banner.setWordWrap(True)
        layout.addWidget(self._banner)

        self._state.configChanged.connect(self._on_config_changed)
        self._on_config_changed()

    def _on_config_changed(self) -> None:
        had_results = self._results_list.count() > 0

        config = self._state.config
        combos = len(config.possible_teams_sizes)
        worst_case_minutes = combos * 30 / 60
        self._search_space_label.setText(
            f"Search space: up to {config.teams_count} Teams, {combos} team size(s), "
            f"up to ~{worst_case_minutes:.1f} min worst-case."
        )
        self._run_button.setEnabled(self._state.is_config_valid())
        self._results_list.clear()
        self._grid.clear()
        self._export_button.setEnabled(False)
        self._banner.setText("Config changed -- please re-run the solver." if had_results else "")

    def _start_run(self) -> None:
        self._log.clear()
        self._results_list.clear()
        self._grid.clear()
        self._export_button.setEnabled(False)
        self._banner.setText("")
        self._progress_bar.setValue(0)

        combos = len(self._state.config.possible_teams_sizes)
        self._progress_bar.setMaximum(max(combos, 1))

        self._thread = QThread()
        self._worker = DistributionWorker(dict(self._state.counts), self._state.config)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)

        self._run_button.setEnabled(False)
        self._load_button.setEnabled(False)
        self._cancel_button.setEnabled(True)
        self._thread.start()

    def _cancel_run(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
        self._cancel_button.setEnabled(False)

    def _on_progress(self, event: ProgressEvent) -> None:
        if event.stage == "started":
            self._log.appendPlainText(
                f"Computing solution for MaxSubteamSize={event.max_subteam_size} (up to {event.num_teams} Teams)"
            )
        else:
            self._log.appendPlainText(
                f"> Computed in {event.solution.time:.2f}s. Result: {event.solution.status.name}"
            )
            self._progress_bar.setValue(event.combo_index + 1)

    def _on_finished(self, solutions: list[Solution]) -> None:
        self._run_button.setEnabled(True)
        self._load_button.setEnabled(True)
        self._cancel_button.setEnabled(False)

        try:
            self._state.set_solutions(solutions)
        except OSError as error:
            QMessageBox.warning(
                self,
                "Could not save distributions",
                f"The computed distributions could not be saved to disk: {error}\n"
                "They are still available for this session, but re-running the "
                "solver later won't be avoidable if the app is closed first.",
            )

        self._populate_results(solutions)

    def _populate_results(self, solutions: list[Solution]) -> None:
        self._results_list.clear()
        self._grid.clear()
        self._export_button.setEnabled(False)
        self._banner.setText("")

        feasible = [s for s in solutions if s.status in (SolutionStatus.FEASIBLE, SolutionStatus.OPTIMAL)]
        feasible.sort(key=lambda s: s.score())

        for solution in feasible:
            label = (
                f"#Teams={len(solution.distribution)}, MaxSubteamSize={solution.max_subteam_size}, "
                f"{solution.status.name}, {solution.time:.1f}s, Score={solution.score()}"
            )
            item = QListWidgetItem(label)
            item.setData(SOLUTION_ROLE, solution)
            self._results_list.addItem(item)

        if self._results_list.count() == 0:
            self._banner.setText("No feasible distribution was found for any configured combination.")

    def _on_error(self, message: str) -> None:
        self._run_button.setEnabled(True)
        self._load_button.setEnabled(True)
        self._cancel_button.setEnabled(False)
        QMessageBox.critical(self, "Solver error", message)

    def _load_saved(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self, "Load saved distributions", self._state.distributions_path, "JSON Files (*.json)"
        )
        if not filename:
            return

        try:
            self._state.load_solutions_from(filename)
        except (OSError, ValueError, KeyError, TypeError) as error:
            QMessageBox.critical(
                self, "Load failed", f"Could not load distributions from {filename}: {error}"
            )
            return

        self._log.clear()
        self._progress_bar.setValue(0)
        self._populate_results(self._state.solutions)

    def _on_result_selected(self) -> None:
        items = self._results_list.selectedItems()
        if not items:
            return
        solution: Solution = items[0].data(SOLUTION_ROLE)
        self._state.select_solution(solution)
        self._grid.load_distribution(solution.distribution, self._state.config, solution.max_subteam_size)
        self._export_button.setEnabled(True)

    def _on_grid_edited(self) -> None:
        self._state.edited_distribution = self._grid.extract_distribution()
        self._state.mark_distribution_edited()

    def _export(self) -> None:
        if self._state.selected_solution is None or self._state.edited_distribution is None:
            return

        filename, _ = QFileDialog.getSaveFileName(
            self, "Export distribution", "distributions.xlsx", "Excel Workbook (*.xlsx)"
        )
        if not filename:
            return

        try:
            export_distribution(
                filename,
                self._state.edited_distribution,
                self._state.config,
                self._state.selected_solution.max_subteam_size,
            )
        except FileWriteError as error:
            QMessageBox.critical(self, "Export failed", str(error))
        else:
            QMessageBox.information(self, "Export complete", f"Distribution exported to {filename}")
