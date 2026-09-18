"""Entry point for the Survivor desktop app.

Run with `python app.py`, or use the PyInstaller-built executable produced by
survivor.spec / .github/workflows/build.yml.
"""
import sys

from PySide6.QtWidgets import QApplication

from survivor_app.ui.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Survivor")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
