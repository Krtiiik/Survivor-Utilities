from __future__ import annotations

import datetime
import itertools
from dataclasses import dataclass

from .config import Config

# Layout constants (rows/cols are 0-indexed, matching the original xlsxwriter layout).
ACTIVITIES_JUMP_ROW = 4  # rows reserved per activity block (aligns with up to 4 Subteams)
TIME_BLOCK_START_ROW = 0
ACTIVITIES_START_ROW = 1
ACTIVITIES_START_COL = 0
TEAMS_START_COL = 1


@dataclass
class TimetableCell:
    row: int
    col: int
    row_span: int
    col_span: int
    text: str | None
    kind: str  # "activity" | "time_block" | "team_all" | "team_rest" | "team_split" | "team_empty"
    color: str | None = None
    border_top: bool = False
    border_bottom: bool = False


@dataclass
class TimetableLayout:
    cells: list[TimetableCell]
    num_rows: int
    num_cols: int


def _unique_splits(num_subteams: int) -> list[list[int]]:
    """Ways to pair Subteam 0 with roughly half the others, rotating which Subteams
    share a split activity together across the schedule."""
    splits = []
    all_subteams = set(range(1, num_subteams))
    for head in itertools.combinations(range(1, num_subteams), (num_subteams - 1) // 2):
        head_set = set(head)
        rest = all_subteams - head_set
        splits.append([0] + sorted(head_set) + sorted(rest))
    return splits


def _activity_ordering(activity_types: list[str], num_teams: int) -> tuple[list[int], list[int]]:
    active, rests = [], []
    for i_activity, activity_type in enumerate(activity_types):
        if activity_type in ("all", "split"):
            active.append(i_activity)
        elif activity_type == "rest":
            rests.append(i_activity)
        else:
            raise ValueError(f"Unrecognized activity type [{activity_type}]")

    all_indices = active + rests
    head, tail = all_indices[:num_teams], all_indices[num_teams:]
    return sorted(head), tail


def compute_timetable_layout(config: Config) -> TimetableLayout:
    """Pure layout computation for the timetable grid -- no xlsxwriter dependency,
    so it can feed both the on-screen preview and the .xlsx export identically."""
    num_teams = config.teams_count
    teams_names = config.teams_names[:num_teams]
    subteams = config.subteams
    num_subteams = config.subteams_count
    activities = config.activities
    num_activities = config.activities_count

    activities_height = num_activities * ACTIVITIES_JUMP_ROW
    num_rows = ACTIVITIES_START_ROW + activities_height
    num_cols = TEAMS_START_COL + num_activities

    cells: list[TimetableCell] = []

    # Activities column: one merged block per activity.
    for i_activity, activity in enumerate(activities):
        row = ACTIVITIES_START_ROW + i_activity * ACTIVITIES_JUMP_ROW
        cells.append(
            TimetableCell(row, ACTIVITIES_START_COL, ACTIVITIES_JUMP_ROW, 1, activity.name, "activity")
        )

    # Time block header row.
    time_start = datetime.datetime.strptime(config.time.start, "%H:%M")
    duration_hours, duration_minutes = (int(p) for p in config.time.activity_duration.split(":"))
    activity_duration = datetime.timedelta(hours=duration_hours, minutes=duration_minutes)
    for i_activity in range(num_activities):
        block_time = time_start + i_activity * activity_duration
        cells.append(
            TimetableCell(
                TIME_BLOCK_START_ROW,
                TEAMS_START_COL + i_activity,
                1,
                1,
                block_time.strftime("%H:%M"),
                "time_block",
            )
        )

    # Team grid.
    teams_labels = [[f"{team} {subteam.name}" for subteam in subteams] for team in teams_names]
    activity_types = [activity.type for activity in activities]
    splits = _unique_splits(num_subteams)
    active_indices, rest_indices = _activity_ordering(activity_types, num_teams)

    for i_team in range(num_teams):
        team_offset = active_indices[i_team]
        split_counter = 0
        team_labels = teams_labels[i_team]
        for i_activity in range(num_activities):
            row = ACTIVITIES_START_ROW + (
                ((i_activity + team_offset) * ACTIVITIES_JUMP_ROW) % activities_height
            )
            col = TEAMS_START_COL + i_activity
            activity_type = activity_types[(i_activity + team_offset) % num_activities]

            if activity_type in ("all", "rest"):
                kind = "team_all" if activity_type == "all" else "team_rest"
                cells.append(TimetableCell(row, col, ACTIVITIES_JUMP_ROW, 1, teams_names[i_team], kind))
            elif activity_type == "split":
                split = splits[split_counter % len(splits)]
                for i_subteam in range(num_subteams):
                    subteam_index = split[i_subteam]
                    # Matches the legacy top/bottom/top/bottom border pattern, which
                    # assumes at most 4 Subteams (same limitation as before).
                    border_top = i_subteam % 2 == 0
                    cells.append(
                        TimetableCell(
                            row + i_subteam,
                            col,
                            1,
                            1,
                            team_labels[subteam_index],
                            "team_split",
                            color=subteams[subteam_index].color,
                            border_top=border_top,
                            border_bottom=not border_top,
                        )
                    )
                split_counter += 1
            else:
                raise ValueError(f"Unrecognized activity type [{activity_type}]")

    for rest_index in rest_indices:
        for i_activity in range(num_activities):
            row = ACTIVITIES_START_ROW + (
                ((i_activity + rest_index) * ACTIVITIES_JUMP_ROW) % activities_height
            )
            col = TEAMS_START_COL + i_activity
            cells.append(TimetableCell(row, col, ACTIVITIES_JUMP_ROW, 1, None, "team_empty"))

    return TimetableLayout(cells, num_rows, num_cols)
