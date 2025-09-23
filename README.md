# Survivor utilities

## Installing

All utilities require Python to run.
Required packages are listed in `requirements.txt` and can by installed using

```
python -m pip install -r requirements.txt
```

preferably into a virtual environment.

## Attendance counter (`counter.py`)

Script used to count attendance of "Kruhy" students.
Run using

```
python counter.py [FILE]
```

The script runs interactively, asking for a new group number.
It keeps count of each entered number.
The total count of all group counts is shown in an self-updating table.
A history of recent increments is shown under the table.
To undo previous count increments, type `-` instead of a number.

The script automatically saves the counts into a JSON `FILE` (default `counts.json`).
The script can be exited using the `Ctrl+C` combination.
Upon running the script with an existing `FILE`, the existing group counts are loaded
and are incremented upon.

## Timesheet builder (`timesheet.py`)

Script used to build an Excel timesheet for the event, describing which team should be at which activity at which time.
Run using

```
python timesheet.py [--config CONFIG] [--output OUTPUT]
```

where

- `CONFIG` holds data about the teams, activities and time settings.
  Default `config.json`.
- `OUTPUT` is the output `.xlsx` Excel file to write to. If this file exists, it is overwritten.
  Default `timesheet.xlsx`.

### Configuration

The configuration is a JSON file, specifying important data about the event.
The following keys need to be present:

- `"Teams count"` (`int`) - number of teams to use.
- `"Teams names"` (`list[string]`) - names of teams.
  Must contain at least as many names as given by `"Teams count"`.
- `"Subteams count"` (`int`) - number of sub-teams per team.
- `"Subteams"` (`list[object]`) - definitions of sub-teams.
  Must contain as many definitions as given by `"Subteams count"`.
  Each is an object with
  - `"Name"` (`str`) - name of the sub-team.
  - `"Color"` (`str`) - background color for the sub-team in split activities.
    Can be an HTML color code (`#rrggbb`) or a simple color name (e.g. `red` or `blue`).
- `"Activities count"` (`int`) - number of activities.
- `"Activities"` (`list[object]`) - definitions of activities.
  Must contain as many definitions as given by `"Activities count"`.
  Each is an object with
  - `"Name"` (`str`) - name of the activity.
  - `"Type"` (`str`) - activity type.
    One of `all`, `split` or `rest`.
- `"Time"` (`object`) - definitions of time constants.
  Contains
  - `"Start"` (`str`) - start of the event in `hh:mm` format.
  - `"Activity duration"` (`str`) - duration of activities in `hh:mm` format.

The resulting timesheet (stored in `OUTPUT`) can then be "printed" (exported) to a pdf file for printing.
Any changes made to the resulting timesheet file are overwritten with each script run.
