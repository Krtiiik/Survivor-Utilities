# Changelog

All notable changes to this project are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and versioning follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [1.5.0] - 2026-09-23

### Changed

- The Distribution solver now decides itself how to split a Kruh that is
  larger than the max Subteam size. It still uses the fewest possible parts,
  all in one Team, but it picks the size of each part. A Kruh of 13 with max
  Subteam size 10 can now become e.g. 7 + 6 instead of always 10 + 3, which
  gives more even Team and Subteam sizes.
- Size balance now trades off against Obory in the solver. It used to only
  break ties: a slightly more mixed Team is now accepted when it makes Team
  and Subteam sizes noticeably more even. Empty Subteams in a Team count as
  imbalanced.
- The Distribution tab shows every Subteam of every Team, including empty
  ones. Candidates are ranked by a `Score` (lower is better) that matches
  the solver's objective, replacing `BalanceScore`.

### Added

- `"Min split part size"` config key (Config tab → Solver search space → "Min
  people per split Kruh part", default 3): the fewest people any part of a
  split Kruh may have. Older configs without the key load with the default.

## [1.4.0] - 2026-09-20

### Changed

- The CLI scripts moved into a `cli/` directory: `counter_cli.py` ->
  `cli/counter_cli.py`, `distribute_cli.py` -> `cli/distribute_cli.py`,
  `timesheet_cli.py` -> `cli/timesheet_cli.py`. Invoke them as e.g.
  `python cli/counter_cli.py` instead of `python counter_cli.py`; arguments
  and defaults are unchanged.

### Added

- `docs/manual.md`: a Czech, non-technical, step-by-step user manual covering
  installation and every tab (Counter, Distribution, Timesheet, Config),
  linked from the README.

### Fixed

- Default Subteam scheme (`example/config.json`, the bundled
  `example_config.json`): the 4th Subteam's color was a pale yellow
  (`#fff491`) instead of the intended purple, and all four Subteams were
  named `"1"`-`"4"` instead of the Č/Z/M/F (red/green/blue/purple) color
  letters used in the event's own planning materials. Renamed and recolored
  to match.

## [1.3.0] - 2026-09-18

### Fixed

- CI: the release job now explicitly requests `contents: write` permission,
  since the default `GITHUB_TOKEN` for this repo is read-only without it --
  `softprops/action-gh-release` was failing with a 403 on tag pushes and no
  GitHub Release was being created.

## [1.2.0] - 2026-09-18

### Changed

- Distribution solver: as a tie-breaker below the existing Team-count and
  Obory-per-Team objectives, the CP-SAT model now also minimizes the spread
  between the largest and smallest Team size, and separately between the
  largest and smallest Subteam size, so Teams and Subteams come out as close
  to equally sized as the other constraints allow.
- Distribution tab: the results list (one candidate per Possible-Teams-count x
  Possible-Teams-size combination) is now sorted by that same size-balance
  score, so the most evenly-sized candidate is listed first instead of just
  following config combination order. Each entry now also shows its score.
- Distribution tab: a Kruh's Obor color now fills its whole row in the grid,
  instead of only the "Team / Subteam / Kruh" label cell.

### Fixed

- Distribution tab: the Team/Subteam/Kruh grid's per-column divider was drawn
  via a `QTreeWidget::item` stylesheet rule, which on Windows' native style
  breaks Qt's normal per-item background painting entirely -- this silently
  blanked out selected rows (until an unrelated repaint) *and* prevented Obor
  row colors from ever showing. The divider is now painted by a small item
  delegate instead, leaving normal background/selection rendering intact.

## [1.1.0] - 2026-09-18

### Fixed

- Timesheet tab: the Activities column no longer ellipsis-truncates long
  Activity names -- `resizeColumnsToContents()` undersizes a column that
  holds a row-spanned cell, so its width is now corrected against the
  actual text width.
- Distribution tab: the Size column now has a visible vertical divider,
  right-aligned numbers, and a snug (non-stretched) width, instead of
  butting up against the Team/Subteam/Kruh labels with no clear boundary.
- Distribution tab: Team, Subteam, and Kruh sizes are now each in their own
  column ("Team size", "Subteam size", "Kruh size", each filled only on the
  rows it applies to) instead of sharing one "Size" column distinguished only
  by indentation.

### Changed

- Obor colors (in the Distribution grid and the exported workbook) are now a
  manually assigned `"Color"` field per Obor in `config.json`, editable on the
  Config tab's Obory editor, instead of being derived from a fixed cycling
  palette. Several Obory can share a color, e.g. every Matematika variant.

### Added

- Counter tab: a keyboard-entry field to record an attendee by typing their
  Kruh number and pressing Enter, an in-session history of recent increments,
  and Undo/Redo for them.
- Counter tab: "Save counts as..." / "Load counts..." (via an OS file picker)
  and "Reset counts", on top of the existing autosave-to-`counts.json` -- a
  save-as or a load never repoints where autosave writes.
- Distribution tab: every solver run now auto-saves all of its computed
  Solutions (not just the selected one) to `distributions.json`, since a solve
  can be expensive to redo. This file is deliberately *not* loaded on startup;
  a new "Load saved distributions..." button (via an OS file picker) restores
  a previous run's candidates on demand instead.

## [1.0.0] - 2026-09-18

### Added

- PySide6 desktop app (`app.py`) wrapping the original three CLI scripts in a
  single tabbed GUI: Counter, Distribution, Timesheet, and Config.
- `survivor_app/core/` — headless business logic (config schema/validation,
  counts, the CP-SAT distribution solver, timesheet layout, Excel export).
- `survivor_app/ui/` — the PySide6 screens, shared `AppState`, and a
  background-thread solver worker so the GUI never blocks.
- GitHub Actions workflow (`.github/workflows/build.yml`) building Windows
  and Linux executables on every push, and publishing them to a GitHub
  Release when a commit is tagged `vX.Y.Z`.
- Headless tests (`tests/test_core.py`, `tests/test_paths.py`) and offscreen
  Qt smoke tests (`tests/smoke_test_gui.py`, `tests/smoke_test_distribution.py`).

### Changed

- The original CLI-only scripts (`counter_cli.py`, `distribute_cli.py`,
  `timesheet_cli.py`) are now thin wrappers over `survivor_app/core/`, kept
  for scripting/automation.

[Unreleased]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.5.0...HEAD
[1.5.0]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/Krtiiik/Survivor-Utilities/releases/tag/v1.0.0
