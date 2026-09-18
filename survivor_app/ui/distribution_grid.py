from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QAbstractItemView, QTreeWidget, QTreeWidgetItem

from ..core.config import Config
from ..core.excel_export import generate_obor_colors
from ..core.models import Kruh, T_Distribution, format_kruh_label

KRUH_ID_ROLE = Qt.ItemDataRole.UserRole
OVERFLOW_COLOR = QColor("#ffcdd2")


class DistributionGrid(QTreeWidget):
    """Editable Team > Subteam > Kruh view. Kruhy can be dragged between Subteams
    (and Teams); Subteam/Team size labels recompute live on every drop, replacing
    the old workflow of cutting and pasting cells in Excel."""

    edited = Signal()

    def __init__(self):
        super().__init__()
        self.setColumnCount(2)
        self.setHeaderLabels(["Team / Subteam / Kruh", "Size"])
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        self._max_subteam_size = 0
        self._kruh_lookup: dict[int, Kruh] = {}

    def load_distribution(self, distribution: T_Distribution, config: Config, max_subteam_size: int) -> None:
        self._max_subteam_size = max_subteam_size
        self._kruh_lookup = {}
        self.clear()

        all_kruhy = [kruh for team in distribution for subteam in team for kruh in subteam]
        obor_colors = generate_obor_colors([kruh.obor for kruh in all_kruhy])

        for i_team, team in enumerate(distribution):
            team_name = (
                config.teams_names[i_team] if i_team < len(config.teams_names) else f"Team {i_team + 1}"
            )
            team_item = QTreeWidgetItem([team_name, ""])
            team_item.setFlags(team_item.flags() & ~Qt.ItemFlag.ItemIsDragEnabled)
            self.addTopLevelItem(team_item)

            for i_subteam, subteam in enumerate(team):
                subteam_name = (
                    config.subteams[i_subteam].name if i_subteam < len(config.subteams) else str(i_subteam + 1)
                )
                subteam_item = QTreeWidgetItem([subteam_name, ""])
                subteam_item.setFlags(
                    (subteam_item.flags() | Qt.ItemFlag.ItemIsDropEnabled) & ~Qt.ItemFlag.ItemIsDragEnabled
                )
                team_item.addChild(subteam_item)

                for kruh in subteam:
                    self._kruh_lookup[kruh.id] = kruh
                    kruh_item = QTreeWidgetItem([format_kruh_label(kruh), str(kruh.count)])
                    kruh_item.setData(0, KRUH_ID_ROLE, kruh.id)
                    kruh_item.setFlags(
                        (kruh_item.flags() | Qt.ItemFlag.ItemIsDragEnabled) & ~Qt.ItemFlag.ItemIsDropEnabled
                    )
                    color = obor_colors.get(kruh.obor)
                    if color:
                        kruh_item.setBackground(0, QBrush(QColor(color)))
                    subteam_item.addChild(kruh_item)

            team_item.setExpanded(True)
            for i_subteam in range(team_item.childCount()):
                team_item.child(i_subteam).setExpanded(True)

        self._recompute_sizes()
        self._flag_split_friends()

    def dropEvent(self, event) -> None:
        super().dropEvent(event)
        self._recompute_sizes()
        self._flag_split_friends()
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
                subteam_item.setText(1, str(size))
                subteam_item.setBackground(1, QBrush(OVERFLOW_COLOR) if size > self._max_subteam_size else QBrush())
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
