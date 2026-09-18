from __future__ import annotations

from PySide6.QtGui import QBrush, QColor, QFontMetrics
from PySide6.QtWidgets import QAbstractItemView, QTableWidget, QTableWidgetItem

from ..core.timesheet import TimetableLayout

# resizeColumnsToContents() undersizes a column that holds a row-spanned cell (e.g. the
# Activities column, one merged cell per activity spanning several rows) -- it doesn't
# account for the anchor item's true text width, so long Activity names get elided
# instead of widening the column. Corrected below by measuring those cells ourselves.
_COLUMN_WIDTH_PADDING = 16

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

        min_column_widths: dict[int, int] = {}
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
            if cell.row_span > 1 and cell.text:
                width = QFontMetrics(item.font()).horizontalAdvance(cell.text) + _COLUMN_WIDTH_PADDING
                min_column_widths[cell.col] = max(min_column_widths.get(cell.col, 0), width)

        self.resizeColumnsToContents()
        self.resizeRowsToContents()

        for col, width in min_column_widths.items():
            if width > self.columnWidth(col):
                self.setColumnWidth(col, width)
