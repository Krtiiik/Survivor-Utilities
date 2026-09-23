# Survivor utilities

A desktop app for running a Kruh-based team-building event ("Survivor"): headcount
counting, team/subteam distribution, and timetable generation, all in one place.

> Specific words -- such as *Kruhy* or *Obory* -- are not translated. It may prove
> difficult to read such a combination of languages... for that I apologize.

For a non-technical, step-by-step walkthrough of the app in Czech (aimed at
whoever actually runs the event), see [`docs/manual.md`](docs/manual.md).

## Running the app

### Option A: download a prebuilt executable

Every push to `main` and every tagged release builds a self-contained Windows
and Linux executable (see `.github/workflows/build.yml`). Tagged releases
(`vX.Y.Z`) publish both to the repository's **Releases** page -- download the
one for your OS and double-click it. No Python installation needed.

### Option B: run from source

Requires Python 3.11+.

```
python -m pip install -r requirements.txt
python app.py
```

preferably from within a virtual environment.

On first launch, the app looks for `config.json`/`counts.json` next to itself
(next to the `.exe` when using a prebuilt binary, or next to `app.py` when
running from source) and seeds a starting `config.json` from the bundled
example if none exists.

## The app

The app has four tabs, usable in any order (not a strict wizard, since the real
workflow often isn't linear -- you might recount mid-event, or regenerate the
timesheet without re-running the solver):

1. **Counter** -- click `+`/`-` next to each Kruh (grouped by Obor), or type a
   Kruh number into the entry field and press Enter, to record attendance as
   people arrive. Counts autosave to `counts.json` after every change. Every
   increment (button or keyboard) appears in the "Recent increments" list and
   can be undone/redone. "Save counts as..." exports the current counts to a
   file of your choice (an OS file picker), "Load counts..." replaces the
   current counts with those from a chosen file, and "Reset counts" clears
   them back to zero -- all three ask for confirmation where destructive, and
   none of them change where autosave writes: it always keeps saving to the
   default `counts.json`, even right after a save-as or a load.
2. **Distribution** -- runs a solver that assigns Kruhy into Teams and
   Subteams (see "Team distribution algorithm" below), lets you pick from the
   resulting candidates, then drag Kruhy between Teams/Subteams directly in the
   app to fine-tune the result (Subteam sizes recalculate live). Export writes
   a static `.xlsx` -- unlike the old workflow, there's nothing fragile about
   editing it afterwards. Every solver run (the computation can take a while)
   auto-saves all of its candidates -- feasible or not -- to `distributions.json`,
   which is *not* reloaded automatically on the next launch; use "Load saved
   distributions..." to bring an earlier run's candidates back without having
   to recompute them.
3. **Timesheet** -- previews the event timetable (which Team/Subteam is doing
   which Activity, when) and exports it to `.xlsx` for printing.
4. **Config** -- edit Obory, Teams, Subteams, Activities, event timing, and the
   solver's search space, without ever hand-editing `config.json`.

The Distribution and Timesheet tabs are disabled whenever the Config is
invalid; fix it on the Config tab first.

### Team distribution algorithm

For each value of "Possible Teams sizes" (the max Subteam size, edited on the
Config tab), the solver builds a constraint model and assigns the Kruhy into
at most as many Teams as there are "Teams names", and into their Subteams. It
minimizes, in order:

- Number of Teams used (so it picks the Team count itself).
- A weighted blend of the number of different Obory within each Team and how
  uneven the Team and Subteam sizes are.

A Kruh larger than the max Subteam size is spread across the fewest possible
Subteams of its Team, with the solver choosing each part's size.

Every size that finds a feasible or optimal solution shows up as a candidate
you can select and then edit.

## Command-line scripts

The GUI's logic lives in `survivor_app/core/` (no GUI dependency), reused by
three thin CLI wrappers in `cli/` for scripting/automation:

```
python cli/counter_cli.py [FILE]
python cli/distribute_cli.py [--config CONFIG] [--counts COUNTS] [--output OUTPUT]
python cli/timesheet_cli.py [--config CONFIG] [--output OUTPUT]
```

These mirror the original standalone scripts' interfaces and defaults
(`config.json`, `counts.json`, `distributions.xlsx`/`timesheet.xlsx`).

## Configuration

`config.json` is a single file read by all three parts of the app/CLI. Edit it
through the app's Config tab, or by hand using the schema below.

- `"Possible Teams sizes"` (`list[int]`) -- candidate max Subteam sizes for the
  solver to try.
- `"Min split part size"` (`int`, optional, default `3`) -- a Kruh larger than
  the max Subteam size is spread across the fewest possible Subteams of one
  Team, with the solver choosing how many people go into each part. This is
  the fewest people any part may have. It is lowered automatically for a Kruh
  that can't be split that evenly.
- `"Solver time limit"` (`int`, optional, default `30`) -- seconds the solver
  may spend on each Possible Teams size. When it runs out, the best
  distribution found so far is used. Longer usually gives more even Teams and
  Subteams.
- `"Teams names"` (`list[string]`) -- one name per Team. The Timesheet renders
  every Team, and the Distribution solver uses at most this many (fewer if
  that works better). Older configs' `"Teams count"` and `"Possible Teams
  counts"` keys are ignored.
- `"Subteams"` (`list[object]`) -- one entry per Subteam:
  - `"Name"` (`str`)
  - `"Color"` (`str`) -- background color for the Subteam in split Activities,
    as `#rrggbb`.
- `"Activities"` (`list[object]`) -- one entry per Activity, in schedule order:
  - `"Name"` (`str`)
  - `"Type"` (`str`) -- one of `all`, `split`, or `rest`.
- `"Time"` (`object`):
  - `"Start"` (`str`) -- event start, `HH:MM`.
  - `"Activity duration"` (`str`) -- length of each Activity block, `HH:MM`.
- `"Obory"` (`list[object]`) -- one entry per Obor:
  - `"Name"` (`str`)
  - `"Kruhy"` (`list[int]`) -- Kruh ids belonging to this Obor.
  - `"Color"` (`str`) -- background color for this Obor's Kruhy in the
    Distribution grid and exported workbook, as `#rrggbb`. Several Obory can
    share the same color (e.g. multiple Matematika variants); defaults to
    white if omitted (configs saved before this field existed).

Number of Subteams/Activities is simply the length of the `"Subteams"`/
`"Activities"` lists -- there's no separate count field to keep in sync.

## Development

```
python -m pip install -r requirements.txt
python -m pip install pyinstaller pytest  # dev extras
python tests/test_core.py                 # headless core logic tests
python app.py                             # run the app
pyinstaller survivor.spec                 # build a standalone executable locally
```

`survivor_app/core/` has no GUI dependency and can be tested without a display.
`survivor_app/ui/` holds the PySide6 screens.
