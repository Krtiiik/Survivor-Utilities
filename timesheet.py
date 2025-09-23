import argparse
import datetime
import itertools
import json
import sys

import xlsxwriter
import xlsxwriter.exceptions
import xlsxwriter.worksheet


parser = argparse.ArgumentParser()
parser.add_argument("--config", type=str, default="config.json")
parser.add_argument("--output", type=str, default="timesheet.xlsx")


TEAM_FORMAT = "{team} {subteam}"


class Structure:
    class Activities:
        start_row = 1
        start_col = 0
        jump_row = 4
        jump_col = 0

        row_height = 15
        column_width = 30

    class TimeBlocks:
        start_row = 0
        start_col = 1
        jump_row = 0
        jump_col = 1

        row_height = 30
        column_width = 15

    class TeamAxis:
        start_row = 1
        start_col = 1
        jump_row = 4
        jump_col = 1


class Format:
    _center = {
        "align": "center"
    }
    _vcenter = {
        "valign": "vcenter"
    }
    _activities = _center | _vcenter | {
        "border": 2,
        "font_size": 20,
    }
    _time_blocks = _center | {
        "valign": "bottom",
        "border": 2,
        "font_size": 20,
    }
    _team_all = _center | _vcenter | {
        "border": 1,
    }
    _team_split = _center | _vcenter | {
        "left": 1,
        "right": 1,
    }
    _team_split_top = _team_split | {
        "top": 1,
    }
    _team_split_bottom = _team_split | {
        "bottom": 1,
    }
    _team_rest = _center | _vcenter | {
        "border": 1,
    }
    _team_empty = {
        "border": 1,
        "bg_color": "#cacaca",
    }

    @staticmethod
    def init(workbook: xlsxwriter.Workbook, config):
        Format.activities = workbook.add_format(Format._activities)
        Format.time_blocks = workbook.add_format(Format._time_blocks)
        Format.team_all = workbook.add_format(Format._team_all)
        Format.team_split_subteams = [
            [workbook.add_format(fmt | {"bg_color": subteam["Color"]})
             for fmt in [Format._team_split_top, Format._team_split_bottom, Format._team_split_top, Format._team_split_bottom]]
            for subteam in config["Subteams"]
        ]
        Format.team_rest = workbook.add_format(Format._team_rest)
        Format.team_empty = workbook.add_format(Format._team_empty)


def die(message: str):
    print(message)
    sys.exit(1)


def parse_config(filename: str):
    with open(filename, encoding="utf8") as f:
        config = json.load(f)
    return config


def write_merged_sequence(worksheet: xlsxwriter.worksheet.Worksheet, structure, data, format=None):
    cells = [(structure.start_row + i_row * structure.jump_row,
              structure.start_col + i_col * structure.jump_col)
             for i_row, i_col in zip(range(len(data)), range(len(data)))]
    for datum, (row, col) in zip(data, cells):
        if structure.jump_row <= 1 and structure.jump_col <= 1:
            worksheet.write(row, col, datum, format)
        else:
            worksheet.merge_range(row, col,
                                  row + structure.jump_row - (1 if structure.jump_row > 0 else 0),
                                  col + structure.jump_col - (1 if structure.jump_col > 0 else 0),
                                  datum,
                                  cell_format=format)


def set_timetable_dimensions(timetable, config):
    num_activities = config["Activities count"]

    timetable.set_row(Structure.TimeBlocks.start_row, Structure.TimeBlocks.row_height)
    for i_row in range(Structure.Activities.start_row, Structure.Activities.start_row + num_activities * Structure.Activities.jump_row):
        timetable.set_row(i_row, Structure.Activities.row_height)

    timetable.set_column(Structure.Activities.start_col, Structure.Activities.start_col, Structure.Activities.column_width)
    timetable.set_column(Structure.TimeBlocks.start_col,
                         Structure.TimeBlocks.start_col + num_activities * Structure.TimeBlocks.jump_col - 1,
                         Structure.TimeBlocks.column_width)


def build_activites(worksheet, config):
    num_activities = config["Activities count"]
    activities = config["Activities"][:num_activities]
    activities_names = [activity["Name"] for activity in activities]
    write_merged_sequence(worksheet, Structure.Activities, activities_names, Format.activities)


def build_timeblocks(worksheet, config):
    time_start = datetime.datetime.strptime(config["Time"]["Start"], "%H:%M")
    activity_duration_str = config["Time"]["Activity duration"]
    activity_duration_m, activity_duration_s = map(int, activity_duration_str.split(':'))
    activity_duration = datetime.timedelta(hours=activity_duration_m, minutes=activity_duration_s)
    num_activities = config["Activities count"]

    time_blocks = [time_start + i * activity_duration for i in range(num_activities)]
    time_blocks = [block.strftime("%H:%M") for block in time_blocks]
    write_merged_sequence(worksheet, Structure.TimeBlocks, time_blocks, Format.time_blocks)


def build_teams(worksheet: xlsxwriter.worksheet.Worksheet, config):
    def unique_splits():
        splits = []
        all = set(range(1, num_subteams))
        for head in itertools.combinations(range(1, num_subteams), (num_subteams-1) // 2):
            head = set(head)
            rest = all - head
            splits.append([0] + sorted(head) + sorted(rest))
        return splits

    def activity_ordering():
        active, rests = [], []
        for i_activity in range(num_activities):
            activity_type = activity_types[i_activity]
            if activity_type == "all" or activity_type == "split":
                active.append(i_activity)
            elif activity_type == "rest":
                rests.append(i_activity)
            else:
                raise ValueError(f"Unrecognized activity type [{activity_type}]")

        all = active + rests
        head, tail = all[:num_teams], all[num_teams:]
        return sorted(head), tail

    num_teams = config["Teams count"]
    teams_names = config["Teams names"][:num_teams]
    subteams = config["Subteams"]
    num_activities = config["Activities count"]
    num_subteams = config["Subteams count"]

    activities_height = num_activities * Structure.Activities.jump_row

    teams = [[TEAM_FORMAT.format(team=team, subteam=subteam["Name"]) for subteam in subteams] for team in teams_names]
    activity_types = [activity["Type"] for activity in config["Activities"]]
    splits = unique_splits()
    active_indices, rests_indices = activity_ordering()
    for i_team in range(num_teams):
        team_offset = active_indices[i_team]
        split_counter = 0
        team = teams[i_team]
        for i_activity in range(num_activities):
            row = Structure.Activities.start_row \
                  + (((i_activity + team_offset) * Structure.Activities.jump_row) % activities_height)
            col = Structure.TimeBlocks.start_col + i_activity * Structure.TimeBlocks.jump_col

            activity_type = activity_types[(i_activity + team_offset) % num_activities]
            if activity_type == "all" or activity_type == "rest":
                row_max = row + Structure.Activities.jump_row - (1 if Structure.Activities.jump_row > 0 else 0)
                col_max = col + Structure.TimeBlocks.jump_col - (1 if Structure.TimeBlocks.jump_col > 0 else 0)
                fmt = {
                    "all": Format.team_all,
                    "rest": Format.team_rest,
                }[activity_type]
                worksheet.merge_range(row, col, row_max, col_max, teams_names[i_team], cell_format=fmt)
            elif activity_type == "split":
                split = splits[split_counter % len(splits)]
                for i_subteam in range(num_subteams):
                    worksheet.write(row + i_subteam, col, team[split[i_subteam]], Format.team_split_subteams[split[i_subteam]][i_subteam])
                split_counter += 1
            else:
                raise ValueError(f"Unrecognized activity type [{activity_type}]")

    for rest_index in rests_indices:
        for i_activity in range(num_activities):
            row = Structure.Activities.start_row \
                  + (((i_activity + rest_index) * Structure.Activities.jump_row) % activities_height)
            col = Structure.TimeBlocks.start_col + i_activity * Structure.TimeBlocks.jump_col
            row_max = row + Structure.Activities.jump_row - (1 if Structure.Activities.jump_row > 0 else 0)
            col_max = col + Structure.TimeBlocks.jump_col - (1 if Structure.TimeBlocks.jump_col > 0 else 0)
            worksheet.merge_range(row, col, row_max, col_max, None, Format.team_empty)


def construct_timetable(workbook_file: str, config: dict):
    workbook = xlsxwriter.Workbook(workbook_file)
    Format.init(workbook, config)

    timetable = workbook.add_worksheet("Timetable")

    set_timetable_dimensions(timetable, config)

    build_activites(timetable, config)
    build_timeblocks(timetable, config)
    build_teams(timetable, config)

    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError:
        die(f"[{workbook_file}] cannot be written. It is probably open in another program.")


def main(args: argparse.Namespace):
    config = parse_config(args.config)
    construct_timetable(args.output, config)


if __name__ == "__main__":
    args = parser.parse_args()
    main(args)
