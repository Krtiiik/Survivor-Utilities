"""CLI wrapper for the timetable/timesheet builder, backed by survivor_app.core."""
import argparse
import sys

from survivor_app.core.config import load_config
from survivor_app.core.errors import FileWriteError
from survivor_app.core.excel_export import render_timetable_xlsx
from survivor_app.core.timesheet import compute_timetable_layout

parser = argparse.ArgumentParser()
parser.add_argument("--config", type=str, default="config.json")
parser.add_argument("--output", type=str, default="timesheet.xlsx")


def main(args: argparse.Namespace) -> None:
    config = load_config(args.config)
    layout = compute_timetable_layout(config)

    try:
        render_timetable_xlsx(layout, config, args.output)
    except FileWriteError as error:
        print(str(error))
        sys.exit(1)


if __name__ == "__main__":
    main(parser.parse_args())
