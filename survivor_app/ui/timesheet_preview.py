from __future__ import annotations

from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QAbstractItemView, QTableWidget, QTableWidgetItem

from ..core.timesheet import TimetableLayout

_KIND_COLORS = {
    "team_empty": "#cacaca",
}


class TimesheetPreview(QTableWidget):
    """Read-only preview mirroring the exported .xlsx's layout and colors."""

    def __init__(self):
        super().__init__()
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.horizontalHeader().hide()
        self.verticalHeader().hide()

    def load_layout(self, layout: TimetableLayout) -> None:
        self.clearSpans()
        self.setRowCount(layout.num_rows)
        self.setColumnCount(layout.num_cols)

        for row in range(layout.num_rows):
            for col in range(layout.num_cols):
                self.setItem(row, col, QTableWidgetItem(""))

        for cell in layout.cells:
            item = QTableWidgetItem(cell.text or "")
            color = cell.color or _KIND_COLORS.get(cell.kind)
            if color:
                item.setBackground(QBrush(QColor(color)))
            if cell.kind in ("activity", "time_block"):
                font = item.font()
                font.setPointSize(font.pointSize() + 4)
                font.setBold(True)
                item.setFont(font)
            self.setItem(cell.row, cell.col, item)
            if cell.row_span > 1 or cell.col_span > 1:
                self.setSpan(cell.row, cell.col, cell.row_span, cell.col_span)

        self.resizeColumnsToContents()
        self.resizeRowsToContents()
