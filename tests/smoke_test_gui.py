"""Headless-ish smoke test: builds the MainWindow, exercises the four tabs against
the bundled example config, and quits automatically. Not a full UI test suite,
but catches import errors, constructor crashes, and basic wiring problems that
py_compile can't."""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from survivor_app.ui import paths as ui_paths

# Run against an isolated scratch directory so this test never writes
# config.json/counts.json into the real repository root.
_scratch_dir = tempfile.mkdtemp(prefix="survivor_gui_smoke_")
ui_paths.app_dir = lambda: _scratch_dir

from survivor_app.ui.main_window import MainWindow


def run():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()

    errors = []

    def exercise():
        try:
            for i in range(window._tabs.count()):
                window._tabs.setCurrentIndex(i)
                app.processEvents()

            counter_tab = window._tabs.widget(0)
            if counter_tab._count_labels:
                any_kruh_id = next(iter(counter_tab._count_labels))
                counter_tab._increment(any_kruh_id)
                app.processEvents()
                assert window.state.counts[any_kruh_id] >= 1
                counter_tab._decrement(any_kruh_id)
                app.processEvents()

                before = window.state.counts.get(any_kruh_id, 0)
                counter_tab._entry_field.setText(str(any_kruh_id))
                counter_tab._submit_entry()
                app.processEvents()
                assert window.state.counts[any_kruh_id] == before + 1
                assert counter_tab._entry_field.text() == ""
                assert counter_tab._history_list.count() >= 1

                counter_tab._entry_field.setText("999999")
                counter_tab._submit_entry()
                app.processEvents()
                assert window.state.counts.get(999999, 0) == 0, "unknown Kruh must not be counted"

                assert counter_tab._undo_button.isEnabled()
                counter_tab._undo()
                app.processEvents()
                assert window.state.counts.get(any_kruh_id, 0) == before
                assert counter_tab._redo_button.isEnabled()

                counter_tab._redo()
                app.processEvents()
                assert window.state.counts[any_kruh_id] == before + 1

                # Save-as/load/reset go through AppState directly, skipping the OS
                # file picker (which would block waiting for input in a headless run).
                snapshot_path = os.path.join(_scratch_dir, "counts_snapshot.json")
                snapshot_counts = dict(window.state.counts)
                window.state.save_counts_as(snapshot_path)
                assert os.path.exists(snapshot_path)

                window.state.reset_counts()
                app.processEvents()
                assert window.state.counts == {}
                assert not window.state.history.can_undo()

                window.state.load_counts_from(snapshot_path)
                app.processEvents()
                assert window.state.counts == snapshot_counts
                assert not window.state.history.can_undo(), "loading should start a fresh history"

            config_tab = window._tabs.widget(3)
            config_tab._reload()
            built = config_tab._build_config()
            assert built.obory, "Config editor should round-trip the example Obory"

            timesheet_tab = window._tabs.widget(2)
            timesheet_tab.refresh()
            assert timesheet_tab._preview.rowCount() > 0

            print("GUI smoke test passed.")
        except Exception as exc:  # noqa: BLE001
            errors.append(exc)
            import traceback
            traceback.print_exc()
        finally:
            app.quit()

    QTimer.singleShot(200, exercise)
    app.exec()

    if errors:
        sys.exit(1)


if __name__ == "__main__":
    run()
