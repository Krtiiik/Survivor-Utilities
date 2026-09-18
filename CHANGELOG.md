# Changelog

All notable changes to this project are documented in this file.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and versioning follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Fixed

- Timesheet tab: the Activities column no longer ellipsis-truncates long
  Activity names -- `resizeColumnsToContents()` undersizes a column that
  holds a row-spanned cell, so its width is now corrected against the
  actual text width.
- Distribution tab: the Size column now has a visible vertical divider,
  right-aligned numbers, and a snug (non-stretched) width, instead of
  butting up against the Team/Subteam/Kruh labels with no clear boundary.

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

[Unreleased]: https://github.com/Krtiiik/Survivor-Utilities/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/Krtiiik/Survivor-Utilities/releases/tag/v1.0.0
