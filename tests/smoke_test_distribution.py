"""Exercises the Distribution tab's threaded solver run end-to-end: starts a run
with a tiny, fast search space, waits for completion, selects a result, edits the
grid, and exports to a temp .xlsx -- the riskiest wiring in the whole app
(QThread + CP-SAT + drag-and-drop grid + export) that the plain compile/import
smoke test doesn't touch."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from survivor_app.ui import paths as ui_paths

_scratch_dir = tempfile.mkdtemp(prefix="survivor_gui_smoke_")
ui_paths.app_dir = lambda: _scratch_dir

from survivor_app.core.config import OborConfig, SubteamConfig
from survivor_app.ui.main_window import MainWindow


def run():
    app = QApplication(sys.argv)
    window = MainWindow()

    # Small, fast-to-solve config so the smoke test doesn't take SOLVER_TIME_LIMIT
    # seconds per combination.
    config = window.state.config
    config.teams_count = 2
    config.possible_teams_counts = [2]
    config.possible_teams_sizes = [6]
    config.teams_names = ["Team A", "Team B"]
    config.subteams = [SubteamConfig("1", "#ffffff"), SubteamConfig("2", "#000000")]
    config.obory = [OborConfig("Fyzika", [11, 12]), OborConfig("Informatika", [21])]
    window.state.counts = {11: 4, 12: 4, 21: 4}

    from survivor_app.core.config import validate as _validate

    validation_errors = _validate(config)
    print("config valid:", not validation_errors, validation_errors, flush=True)

    window.state.configChanged.emit()
    print("configChanged emitted OK", flush=True)

    distribution_tab = window._tabs.widget(1)

    errors = []

    def check_finished():
        try:
            if distribution_tab._results_list.count() == 0:
                errors.append(AssertionError("Solver produced no feasible results"))
                app.quit()
                return

            distribution_tab._results_list.setCurrentRow(0)
            app.processEvents()

            grid = distribution_tab._grid
            assert grid.topLevelItemCount() == 2, "Expected 2 Team rows in the grid"

            extracted = grid.extract_distribution()
            assigned_ids = {kruh.id for team in extracted for subteam in team for kruh in subteam}
            assert assigned_ids == {11, 12, 21}, f"Unexpected assigned ids: {assigned_ids}"

            export_path = os.path.join(_scratch_dir, "distributions.xlsx")
            from survivor_app.core.excel_export import export_distribution

            export_distribution(
                export_path,
                window.state.edited_distribution,
                window.state.config,
                window.state.selected_solution.max_subteam_size,
            )
            assert os.path.exists(export_path) and os.path.getsize(export_path) > 0

            print("Distribution smoke test passed.")
        except Exception as exc:  # noqa: BLE001
            errors.append(exc)
            import traceback
            traceback.print_exc()
        finally:
            app.quit()

    poll_count = [0]

    def poll():
        poll_count[0] += 1
        print(f"poll #{poll_count[0]}: run_button enabled={distribution_tab._run_button.isEnabled()}", flush=True)
        if distribution_tab._run_button.isEnabled():
            check_finished()
        elif poll_count[0] > 50:  # ~10s of polling at 200ms
            errors.append(TimeoutError("Solver run did not finish within ~10s"))
            app.quit()
        else:
            QTimer.singleShot(200, poll)

    def start():
        print("starting solver run...", flush=True)
        distribution_tab._start_run()
        print("_start_run() returned", flush=True)
        QTimer.singleShot(200, poll)

    QTimer.singleShot(100, start)
    app.exec()
    print("app.exec() returned", flush=True)

    if errors:
        sys.exit(1)


if __name__ == "__main__":
    run()
