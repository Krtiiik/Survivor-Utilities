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
python counter.py [file]
```

The script runs interactively, asking for a new group number.
It keeps count of each entered number.
The total count of all group counts is shown in an self-updating table.
A history of recent increments is shown under the table.
To undo previous count increments, type `-` instead of a number.

The script automatically saves the counts into a JSON `file` (default `counts.json`).
The script can be exited using the `Ctrl+C` combination.
Upon running the script with an existing `file`, the existing group counts are loaded
and are incremented upon.
