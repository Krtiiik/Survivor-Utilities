from __future__ import annotations

from PySide6.QtCore import QObject, Signal

from ..core.config import Config
from ..core.distribute import compute_distributions
from ..core.models import Solution


class DistributionWorker(QObject):
    """Runs compute_distributions off the UI thread (meant to be moved via
    QThread.moveToThread). A single CP-SAT solve can take up to SOLVER_TIME_LIMIT
    seconds, and a full run tries every Possible Teams size, so this must never
    run on the GUI thread."""

    progress = Signal(object)  # core.models.ProgressEvent
    finished = Signal(list)  # list[Solution]
    error = Signal(str)

    def __init__(self, counts: dict[int, int], config: Config):
        super().__init__()
        self._counts = counts
        self._config = config
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def _should_cancel(self) -> bool:
        return self._cancelled

    def run(self) -> None:
        try:
            solutions = compute_distributions(
                self._counts,
                self._config,
                progress_callback=self.progress.emit,
                should_cancel=self._should_cancel,
            )
        except Exception as exc:  # surfaced to the GUI thread via a signal, not raised
            self.error.emit(str(exc))
            return
        self.finished.emit(solutions)
