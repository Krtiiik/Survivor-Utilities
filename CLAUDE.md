# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A PySide6 desktop app for running a Kruh-based team-building event ("Survivor"): headcount counting, team/subteam distribution, and timetable generation, walked through in one GUI. The original CLI-only version (three standalone scripts) was reworked into this app so a non-technical successor could run the whole event without touching JSON or Excel formulas; the CLI scripts still exist as thin wrappers for scripting/automation.

## Setup and running

```
python -m pip install -r requirements.txt
python app.py                                                    # the GUI app
python cli/counter_cli.py [FILE]                                     # default FILE = counts.json
python cli/distribute_cli.py [--config CONFIG] [--counts COUNTS] [--output OUTPUT]  # defaults: config.json, counts.json, distributions.xlsx
python cli/timesheet_cli.py [--config CONFIG] [--output OUTPUT]      # defaults: config.json, timesheet.xlsx
```

`tests/test_core.py` and `tests/test_paths.py` are headless (no Qt/display) tests over `survivor_app/core/` and `survivor_app/ui/paths.py`; run them directly, e.g. `python tests/test_core.py`. `tests/smoke_test_gui.py` and `tests/smoke_test_distribution.py` exercise the PySide6 layer (including the threaded solver run) with `QT_QPA_PLATFORM=offscreen`, useful when there's no display attached -- but note they monkeypatch `paths.app_dir()` to an isolated scratch directory, so they won't catch bugs in the real path-resolution logic itself (that's what `test_paths.py` is for; a real `app_dir()` bug once slipped past both offscreen smoke tests for exactly this reason). There is no linter configured.

## Domain glossary

Czech terms are used as-is (not translated) throughout code, config, and output — stick to these names in code, comments, and discussion:

- **Obor** — a field of study (e.g. "Fyzika", "Informatika"). Plural: **Obory**.
- **Kruh** — a study group ("circle") within an Obor, identified by a numeric ID (e.g. `11`, `35`). Plural: **Kruhy**. This is the atomic unit of attendees being distributed.
- **Team** — top-level event team (e.g. "α [Alfa]"). Kruhy are assigned into Teams.
- **Subteam** — one of a fixed number of subdivisions within every Team, used for `split`-type activities. Kruhy within a Team are further assigned into Subteams.
- **Activity** — a timetable slot, one of three `Type`s: `all` (whole Team together), `split` (Team divided across its Subteams), `rest` (Team idle).

## Module layout

- `survivor_app/core/` — all business logic, **zero PySide6 imports**, testable headless. `config.py` (schema + validation), `counts.py` (headcount load/save/increment/decrement), `models.py` (`Kruh`, `Solution`, `format_kruh_label`), `distribute.py` (the CP-SAT solver), `timesheet.py` (pure timetable layout math), `excel_export.py` (all xlsxwriter code, for both distribution and timesheet output), `errors.py`.
- `survivor_app/ui/` — the PySide6 app: `main_window.py` (tabbed nav: Counter/Distribution/Timesheet/Config), `state.py` (`AppState`, the single source of truth shared across tabs, with Qt signals), `workers.py` (`DistributionWorker`, runs the solver on a `QThread` so the GUI thread never blocks), one screen module per tab, `config_editor/` (one sub-editor per config section), `widgets/` (small reusable controls).
- `app.py` — GUI entry point. `cli/counter_cli.py` / `cli/distribute_cli.py` / `cli/timesheet_cli.py` — thin argparse wrappers over `core/`, kept for scripting/automation and so `core/` stays testable without PySide6.
- `example/` holds a sample `config.json` + `counts.json` pair; `survivor_app/resources/example_config.json` is the same config bundled into the app/executable as the first-run default.

## Data flow

1. **Counter tab** (`counter_screen.py`) — click +/- per Kruh (grouped by Obor); autosaves `counts.json` via `core.counts` after every click, same cadence as the old REPL.
2. **Distribution tab** (`distribution_screen.py` + `distribution_grid.py`) — reads `counts.json` + `config.json`, runs `core.distribute.compute_distributions` (OR-Tools CP-SAT) on a background thread, lets the user pick a candidate and drag-edit Kruhy between Teams/Subteams, then exports a single static `.xlsx` via `core.excel_export.export_distribution`.
3. **Timesheet tab** (`timesheet_screen.py` + `timesheet_preview.py`) — reads `config.json` only (not the distribution output); `core.timesheet.compute_timetable_layout` is pure layout math shared by both the on-screen preview and `core.excel_export.render_timetable_xlsx`.
4. **Config tab** (`config_editor/`) — edits an in-memory `Config` and only writes `config.json` (via `core.config.save_config`) once `core.config.validate` passes.

All four tabs share one `AppState` instance (`ui/state.py`); changing the Config invalidates any previously computed distribution (the Distribution tab shows a "please re-run the solver" banner).

## `core/distribute.py` architecture

This is the only module with real algorithmic complexity — a constraint-satisfaction model, not a heuristic. The CP-SAT model itself (variables, constraints, objective, symmetry-breaking) is untouched from the original script; only its Obor-domain source and cancellation support changed.

- Obor names are **config-driven**, not a hardcoded enum — `compute_teams_distribution` derives the Obor domain from `sorted(set(kruh.obor for kruh in kruhy))`. If you're tempted to reintroduce a fixed enum for Obor, don't — the whole point of the Config tab is that Obor names are freely editable.
- The solver searches every combination of `Possible Teams counts` × `Possible Teams sizes` from the config (`compute_distributions`), building and solving a fresh CP-SAT model per combination (`compute_teams_distribution`), and keeps all resulting `Solution`s (feasible or optimal ones are shown as selectable candidates in the Distribution tab).
- Objective minimizes, in order of the summed expression: number of Teams used, then number of distinct Obory per Team.
- Kruhy larger than the candidate Subteam size are pre-split into synthetic sub-Kruhy (`compute_kruhy_split`) with IDs of the form `100*kruh_id + part_index`; these are tied back together as "friends" that the model constrains to land in the same Team (though they may land in different Subteams). `core.models.format_kruh_label` renders these split IDs back as `"{id}[{letter}]"` for both the GUI grid and Excel export. The `DistributionGrid` shows a soft ⚠ warning (not a hard block) if a user manually drags split-Kruh siblings into different Teams — full manual override is intentional.
- A single 30s (`SOLVER_TIME_LIMIT`) solve can be cancelled early via a `CpSolverSolutionCallback` (`_CancelCallback`) driven by the GUI's Cancel button; the outer per-combination loop also checks the same `should_cancel` between combinations.
- Export (`core.excel_export.export_distribution`) writes only the single selected/edited distribution as a static workbook — Subteam sizes are plain numbers computed in-memory, **not** live `XLOOKUP` formulas, so (unlike the original script's output) the exported file has no cut-and-paste-only editing caveat.

## Versioning

The project follows [Semantic Versioning](https://semver.org/) starting at
`1.0.0` (the PySide6 rework). `CHANGELOG.md` follows
[Keep a Changelog](https://keepachangelog.com/en/1.0.0/): every user-facing
change (new feature, fix, behavior change) gets an entry under
`## [Unreleased]` at the top, added in the same commit/session as the change
itself, not batched later.

When a batch of `[Unreleased]` changes is substantial enough to be worth
shipping as a release (judgment call — a meaningful feature or fix, not every
single commit), move that section under a new `## [X.Y.Z] - YYYY-MM-DD`
heading, bump the version (patch for fixes, minor for backwards-compatible
features, major for breaking changes to `config.json`'s schema or the CLI
interfaces), update the compare links at the bottom of the file, and tag the
commit: `git tag -a vX.Y.Z -m "vX.Y.Z"`. Pushing that tag (`git push origin
vX.Y.Z`) triggers `.github/workflows/build.yml`'s `release` job, which builds
the Windows/Linux executables and publishes them to a GitHub Release — so
confirm with the user before pushing a tag, the same way you would before any
other action visible to others.

## Config schema notes

`"Subteams count"`/`"Activities count"` were removed from `config.json` — those counts are now derived from `len(config["Subteams"])`/`len(config["Activities"])` (see `Config.subteams_count`/`activities_count` in `core/config.py`). `"Teams count"` is *not* redundant with `len(Teams names)` and was kept: it's the actual number of Teams the Timesheet renders, while `"Teams names"` just needs to be at least that long (and at least as long as the largest value in `"Possible Teams counts"`, for the Distribution solver). `"Activity duration"` is HH:MM (matches the README and the code's actual runtime behavior).
