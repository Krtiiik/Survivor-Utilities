from __future__ import annotations

import xlsxwriter
import xlsxwriter.exceptions
import xlsxwriter.worksheet

from .config import Config
from .errors import FileWriteError
from .models import T_Distribution, format_kruh_label
from .timesheet import TimetableLayout


def obor_color_map(config: Config) -> dict[str, str]:
    """Per-Obor colors as set on the Config tab -- manually assigned rather than
    derived from a fixed palette, so e.g. every Matematika-variant Obor can share
    one color while Fyzika/Informatika/Ucitelstvi each keep their own. Shared by
    the in-app Distribution grid and the .xlsx export."""
    return {obor.name: obor.color for obor in config.obory}


# ---------------------------------------------------------------------------
# Distribution export
# ---------------------------------------------------------------------------

def export_distribution(
    filename: str,
    distribution: T_Distribution,
    config: Config,
    max_subteam_size: int,
) -> None:
    """Write a single (possibly user-edited) distribution as a static workbook.

    Unlike the legacy output, Subteam sizes are written as plain numbers computed
    in-memory -- there are no more live XLOOKUP formulas, so the workbook can be
    freely edited afterwards without any cut-and-paste caveat.
    """
    obor_colors = obor_color_map(config)

    workbook = xlsxwriter.Workbook(filename)
    try:
        fmt_obor = {
            name: workbook.add_format({"border": 1, "bg_color": color})
            for name, color in obor_colors.items()
        }
        fmt_team = workbook.add_format(
            {"align": "center", "valign": "vcenter", "border": 1, "top": 2, "bottom": 2}
        )
        fmt_count = workbook.add_format({"border": 1})
        fmt_overflow = workbook.add_format({"bold": 1, "font_color": "#ff0000"})

        kruhy_sheet = workbook.add_worksheet("Kruhy")
        _write_kruhy_sheet(kruhy_sheet, distribution)

        teams_sheet = workbook.add_worksheet("Teams")
        _write_teams_sheet(teams_sheet, distribution, config, max_subteam_size, fmt_team, fmt_count, fmt_obor, fmt_overflow)

        workbook.close()
    except xlsxwriter.exceptions.FileCreateError:
        raise FileWriteError(filename)


def _write_kruhy_sheet(worksheet: xlsxwriter.worksheet.Worksheet, distribution: T_Distribution) -> None:
    kruhy = [kruh for team in distribution for subteam in team for kruh in subteam]

    worksheet.write(0, 0, "Kruh")
    worksheet.write(0, 1, "Size")
    for i_kruh, kruh in enumerate(sorted(kruhy, key=lambda k: k.id)):
        worksheet.write_string(1 + i_kruh, 0, format_kruh_label(kruh))
        worksheet.write_number(1 + i_kruh, 1, kruh.count)


def _write_teams_sheet(
    worksheet: xlsxwriter.worksheet.Worksheet,
    distribution: T_Distribution,
    config: Config,
    max_subteam_size: int,
    fmt_team,
    fmt_count,
    fmt_obor: dict[str, object],
    fmt_overflow,
) -> None:
    team_names = config.teams_names
    num_teams = len(distribution)
    num_subteams = config.subteams_count

    for i_team, team in enumerate(distribution):
        row_team = i_team * num_subteams

        worksheet.merge_range(
            row_team, 0, row_team + num_subteams - 1, 0, team_names[i_team], cell_format=fmt_team
        )

        for i_subteam in range(num_subteams):
            row_subteam = row_team + i_subteam
            subteam = team[i_subteam] if i_subteam < len(team) else []
            size = sum(kruh.count for kruh in subteam)
            worksheet.write_number(row_subteam, 1, size, fmt_count)
            for i_kruh, kruh in enumerate(subteam):
                worksheet.write_string(
                    row_subteam, 2 + i_kruh, format_kruh_label(kruh), fmt_obor[kruh.obor]
                )

    worksheet.conditional_format(
        0, 1, num_teams * num_subteams - 1, 1,
        options={
            "type": "data_bar", "min_type": "num", "min_value": 0,
            "max_type": "num", "max_value": max_subteam_size,
        },
    )
    worksheet.conditional_format(
        0, 1, num_teams * num_subteams - 1, 1,
        options={
            "type": "cell", "criteria": "greater than",
            "value": max_subteam_size, "format": fmt_overflow,
        },
    )


# ---------------------------------------------------------------------------
# Timesheet export
# ---------------------------------------------------------------------------

_ACTIVITY_ROW_HEIGHT = 15
_ACTIVITY_COL_WIDTH = 30
_TIME_BLOCK_ROW_HEIGHT = 30
_TIME_BLOCK_COL_WIDTH = 15


def render_timetable_xlsx(layout: TimetableLayout, config: Config, filename: str) -> None:
    workbook = xlsxwriter.Workbook(filename)
    try:
        fmt_activity = workbook.add_format(
            {"align": "center", "valign": "vcenter", "border": 2, "font_size": 20}
        )
        fmt_time_block = workbook.add_format(
            {"align": "center", "valign": "bottom", "border": 2, "font_size": 20}
        )
        fmt_team_all = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
        fmt_team_rest = fmt_team_all
        fmt_team_empty = workbook.add_format(
            {"align": "center", "valign": "vcenter", "border": 1, "bg_color": "#cacaca", "font_size": 16}
        )
        fmt_split_cache: dict[tuple[str, bool, bool], object] = {}

        def fmt_split(color: str, border_top: bool, border_bottom: bool):
            key = (color, border_top, border_bottom)
            if key not in fmt_split_cache:
                spec = {"align": "center", "valign": "vcenter", "left": 1, "right": 1, "bg_color": color}
                if border_top:
                    spec["top"] = 1
                if border_bottom:
                    spec["bottom"] = 1
                fmt_split_cache[key] = workbook.add_format(spec)
            return fmt_split_cache[key]

        worksheet = workbook.add_worksheet("Timetable")

        worksheet.set_row(0, _TIME_BLOCK_ROW_HEIGHT)
        for row in range(1, layout.num_rows):
            worksheet.set_row(row, _ACTIVITY_ROW_HEIGHT)
        worksheet.set_column(0, 0, _ACTIVITY_COL_WIDTH)
        worksheet.set_column(1, layout.num_cols - 1, _TIME_BLOCK_COL_WIDTH)

        for cell in layout.cells:
            row_last = cell.row + cell.row_span - 1
            col_last = cell.col + cell.col_span - 1
            merged = row_last > cell.row or col_last > cell.col

            if cell.kind == "activity":
                fmt = fmt_activity
            elif cell.kind == "time_block":
                fmt = fmt_time_block
            elif cell.kind == "team_all":
                fmt = fmt_team_all
            elif cell.kind == "team_rest":
                fmt = fmt_team_rest
            elif cell.kind == "team_empty":
                fmt = fmt_team_empty
            elif cell.kind == "team_split":
                fmt = fmt_split(cell.color, cell.border_top, cell.border_bottom)
            else:
                raise ValueError(f"Unrecognized timetable cell kind [{cell.kind}]")

            if merged:
                worksheet.merge_range(cell.row, cell.col, row_last, col_last, cell.text, cell_format=fmt)
            else:
                worksheet.write(cell.row, cell.col, cell.text, fmt)

        workbook.close()
    except xlsxwriter.exceptions.FileCreateError:
        raise FileWriteError(filename)
