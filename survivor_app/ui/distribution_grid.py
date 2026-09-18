from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QAbstractItemView, QStyledItemDelegate, QTreeWidget, QTreeWidgetItem

from ..core.config import Config
from ..core.excel_export import obor_color_map
from ..core.models import Kruh, T_Distribution, format_kruh_label

KRUH_ID_ROLE = Qt.ItemDataRole.UserRole
OVERFLOW_COLOR = QColor("#ffcdd2")


class _ColumnDividerDelegate(QStyledItemDelegate):
    """Paints a vertical divider between columns. This used to be a
    `QTreeWidget::item` stylesheet rule, but styling ::item at all makes Qt's
    Windows style stop painting Qt::BackgroundRole colors -- the Obor row colors
    and the Subteam overflow warning both silently stopped rendering because of
    it. Painting the divider by hand keeps the default (working) background/
    selection rendering intact."""

    def paint(self, painter, option, index) -> None:
        super().paint(painter, option, index)
        painter.save()
        pen = painter.pen()
        pen.setColor(option.palette.mid().color())
        painter.setPen(pen)
        painter.drawLine(option.rect.topRight(), option.rect.bottomRight())
        painter.restore()


class DistributionGrid(QTreeWidget):
    """Editable Team > Subteam > Kruh view. Kruhy can be dragged between Subteams
    (and Teams); Subteam/Team size labels recompute live on every drop, replacing
    the old workflow of cutting and pasting cells in Excel."""

    edited = Signal()

    def __init__(self):
        super().__init__()
        self.setColumnCount(4)
        self.setHeaderLabels(["Team / Subteam / Kruh", "Team size", "Subteam size", "Kruh size"])
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        # A single shared "Size" column used to hold Team totals, Subteam sizes, and
        # individual Kruh counts all at once, distinguished only by indentation --
        # easy to misread when scanning down the column. Splitting each into its own
        # column (each blank on rows it doesn't apply to) means a number's meaning is
        # given by which column it's in, not how far indented its row happens to be.
        self.setAlternatingRowColors(True)
        self.setItemDelegate(_ColumnDividerDelegate(self))
        # Keep the size columns snug against their content instead of stretching
        # across any leftover width, so their numbers sit right after the divider
        # above, rather than off in empty space on the far side of a wide panel.
        self.header().setStretchLastSection(False)

        self._max_subteam_size = 0
        self._kruh_lookup: dict[int, Kruh] = {}

    def load_distribution(self, distribution: T_Distribution, config: Config, max_subteam_size: int) -> None:
        self._max_subteam_size = max_subteam_size
        self._kruh_lookup = {}
        self.clear()

        obor_colors = obor_color_map(config)

        for i_team, team in enumerate(distribution):
            team_name = (
                config.teams_names[i_team] if i_team < len(config.teams_names) else f"Team {i_team + 1}"
            )
            team_item = QTreeWidgetItem([team_name, "", "", ""])
            team_item.setFlags(team_item.flags() & ~Qt.ItemFlag.ItemIsDragEnabled)
            team_item.setTextAlignment(1, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.addTopLevelItem(team_item)

            for i_subteam, subteam in enumerate(team):
                subteam_name = (
                    config.subteams[i_subteam].name if i_subteam < len(config.subteams) else str(i_subteam + 1)
                )
                subteam_item = QTreeWidgetItem([subteam_name, "", "", ""])
                subteam_item.setFlags(
                    (subteam_item.flags() | Qt.ItemFlag.ItemIsDropEnabled) & ~Qt.ItemFlag.ItemIsDragEnabled
                )
                subteam_item.setTextAlignment(2, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                team_item.addChild(subteam_item)

                for kruh in subteam:
                    self._kruh_lookup[kruh.id] = kruh
                    kruh_item = QTreeWidgetItem([format_kruh_label(kruh), "", "", str(kruh.count)])
                    kruh_item.setData(0, KRUH_ID_ROLE, kruh.id)
                    kruh_item.setFlags(
                        (kruh_item.flags() | Qt.ItemFlag.ItemIsDragEnabled) & ~Qt.ItemFlag.ItemIsDropEnabled
                    )
                    kruh_item.setTextAlignment(3, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    color = obor_colors.get(kruh.obor)
                    if color:
                        brush = QBrush(QColor(color))
                        for column in range(self.columnCount()):
                            kruh_item.setBackground(column, brush)
                    subteam_item.addChild(kruh_item)

            team_item.setExpanded(True)
            for i_subteam in range(team_item.childCount()):
                team_item.child(i_subteam).setExpanded(True)

        self._recompute_sizes()
        self._flag_split_friends()
        self.resizeColumnToContents(0)
        self.resizeColumnToContents(1)
        self.resizeColumnToContents(2)
        self.resizeColumnToContents(3)

    def dropEvent(self, event) -> None:
        super().dropEvent(event)
        self._recompute_sizes()
        self._flag_split_friends()
        self.resizeColumnToContents(1)
        self.resizeColumnToContents(2)
        self.edited.emit()

    def extract_distribution(self) -> T_Distribution:
        distribution = []
        for i_team in range(self.topLevelItemCount()):
            team_item = self.topLevelItem(i_team)
            team = []
            for i_subteam in range(team_item.childCount()):
                subteam_item = team_item.child(i_subteam)
                subteam = [
                    self._kruh_lookup[subteam_item.child(i_kruh).data(0, KRUH_ID_ROLE)]
                    for i_kruh in range(subteam_item.childCount())
                ]
                team.append(subteam)
            distribution.append(team)
        return distribution

    def _recompute_sizes(self) -> None:
        for i_team in range(self.topLevelItemCount()):
            team_item = self.topLevelItem(i_team)
            team_total = 0
            for i_subteam in range(team_item.childCount()):
                subteam_item = team_item.child(i_subteam)
                size = sum(
                    self._kruh_lookup[subteam_item.child(i_kruh).data(0, KRUH_ID_ROLE)].count
                    for i_kruh in range(subteam_item.childCount())
                )
                team_total += size
                subteam_item.setText(2, str(size))
                subteam_item.setBackground(2, QBrush(OVERFLOW_COLOR) if size > self._max_subteam_size else QBrush())
            team_item.setText(1, str(team_total))

    def _flag_split_friends(self) -> None:
        """Soft warning (not a hard block) when a split Kruh's parts (e.g. 11[a]
        and 11[b]) end up in different Teams after manual dragging."""
        original_kruh_teams: dict[int, set[int]] = {}
        for i_team in range(self.topLevelItemCount()):
            team_item = self.topLevelItem(i_team)
            for i_subteam in range(team_item.childCount()):
                subteam_item = team_item.child(i_subteam)
                for i_kruh in range(subteam_item.childCount()):
                    kruh_id = subteam_item.child(i_kruh).data(0, KRUH_ID_ROLE)
                    if kruh_id >= 100:
                        original_id = kruh_id // 100
                        original_kruh_teams.setdefault(original_id, set()).add(i_team)

        split_across_teams = {
            original_id for original_id, teams in original_kruh_teams.items() if len(teams) > 1
        }

        for i_team in range(self.topLevelItemCount()):
            team_item = self.topLevelItem(i_team)
            for i_subteam in range(team_item.childCount()):
                subteam_item = team_item.child(i_subteam)
                for i_kruh in range(subteam_item.childCount()):
                    kruh_item = subteam_item.child(i_kruh)
                    kruh_id = kruh_item.data(0, KRUH_ID_ROLE)
                    label = format_kruh_label(self._kruh_lookup[kruh_id])
                    is_split_warning = kruh_id >= 100 and (kruh_id // 100) in split_across_teams
                    kruh_item.setText(0, f"⚠ {label}" if is_split_warning else label)
                    kruh_item.setToolTip(
                        0,
                        "This Kruh was split across multiple Teams by the solver, but its "
                        "parts are no longer in the same Team."
                        if is_split_warning
                        else "",
                    )
