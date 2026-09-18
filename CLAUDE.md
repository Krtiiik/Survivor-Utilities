# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Standalone Python scripts for running a Kruh-based team-building event ("Survivor"). No package structure, no tests, no build step — each `.py` file is run directly.

## Setup and running

```
python -m pip install -r requirements.txt
python counter.py [FILE]                                        # default FILE = counts.json
python distribute.py [--config CONFIG] [--counts COUNTS] [--output OUTPUT]  # defaults: config.json, counts.json, distributions.xlsx
python timesheet.py [--config CONFIG] [--output OUTPUT]         # defaults: config.json, timesheet.xlsx
```

There is no test suite or linter configured.

## Domain glossary

Czech terms are used as-is (not translated) throughout code, config, and output — stick to these names in code, comments, and discussion:

- **Obor** — a field of study (e.g. "Fyzika", "Informatika"). Plural: **Obory**.
- **Kruh** — a study group ("circle") within an Obor, identified by a numeric ID (e.g. `11`, `35`). Plural: **Kruhy**. This is the atomic unit of attendees being distributed.
- **Team** — top-level event team (e.g. "α [Alfa]"). Kruhy are assigned into Teams.
- **Subteam** — one of a fixed number of subdivisions within every Team, used for `split`-type activities. Kruhy within a Team are further assigned into Subteams.
- **Activity** — a timetable slot, one of three `Type`s: `all` (whole Team together), `split` (Team divided across its Subteams), `rest` (Team idle).

## Data flow between the three scripts

1. `counter.py` — interactive headcount tool run at the door. Produces `counts.json`: `{kruh_id: attendee_count}`.
2. `distribute.py` — reads `counts.json` + `config.json`, uses OR-Tools CP-SAT to assign Kruhy to Teams/Subteams, writes `distributions.xlsx`.
3. `timesheet.py` — reads `config.json` only (not the distribution output), writes `timesheet.xlsx` describing which Team/Subteam is at which Activity at which time.

All three share one `config.json`, but each script only reads the keys it needs (see `README.md` for the full per-script key list). `example/` holds a sample `config.json` + `counts.json` pair.

## `distribute.py` architecture

This is the only script with real algorithmic complexity — a constraint-satisfaction model, not a heuristic.

- `Obor` is a `StrEnum` **hardcoded** to match the `"Obory"` names in `config.json` — if the config's Obor names change, this enum must be updated to match.
- The solver searches every combination of `Possible Teams counts` × `Possible Team sizes` from the config (`compute_distributions`), building and solving a fresh CP-SAT model per combination (`compute_teams_distribution`), and keeps all resulting `Solution`s (feasible or optimal only get written).
- Objective minimizes, in order of the summed expression: number of Teams used, then number of distinct Obory per Team.
- Kruhy larger than the candidate Subteam size are pre-split into synthetic sub-Kruhy (`compute_kruhy_split`) with IDs of the form `100*kruh_id + part_index`; these are tied back together as "friends" that the model constrains to land in the same Team (though they may land in different Subteams). `Format.format_kruh` renders these split IDs back as `"{id}[{letter}]"` for output.
- Output: one `.xlsx` workbook with a `Kruhy-{numTeams}_{maxSize}` + `Teams-{numTeams}_{maxSize}` worksheet pair per feasible solution, so multiple candidate distributions can be compared side by side before picking one. The Teams sheet uses live `XLOOKUP` formulas against the paired Kruhy sheet to compute Subteam sizes, so rearranging Kruhy in Excel must be done by **cut-and-paste, not by editing cell contents** (see README for why — xlsxwriter/Excel numeric coercion of Kruh names).
