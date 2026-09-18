"""Headless tests for survivor_app.core -- no PySide6/display required."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from survivor_app.core.config import Config, load_config, validate
from survivor_app.core.counts import decrement, increment, load_counts, summarize
from survivor_app.core.distribute import compute_distributions, compute_kruhy_split
from survivor_app.core.models import Kruh, SolutionStatus, format_kruh_label
from survivor_app.core.timesheet import compute_timetable_layout

EXAMPLE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "example")


def test_load_example_config_and_validate():
    config = load_config(os.path.join(EXAMPLE_DIR, "config.json"))
    assert isinstance(config, Config)
    # Old-schema example config still has "Subteams count"/"Activities count" keys;
    # they're simply ignored now that counts are derived from list length.
    assert config.subteams_count == len(config.subteams)
    assert config.activities_count == len(config.activities)
    assert validate(config) == []


def test_counts_increment_decrement_summarize():
    counts = load_counts(os.path.join(EXAMPLE_DIR, "counts.json"))
    assert counts[11] == 8

    increment(11, counts)
    assert counts[11] == 9

    rows, total = summarize(counts)
    assert rows == sorted(counts.items())
    assert total == sum(counts.values())

    fresh: dict[int, int] = {}
    increment(99, fresh)
    assert fresh[99] == 1
    decrement(99, fresh)
    assert 99 not in fresh  # dropped to 0 -> removed, matching legacy behavior
    decrement(99, fresh)  # decrementing an absent Kruh is a no-op
    assert fresh == {}


def test_format_kruh_label_handles_splits():
    assert format_kruh_label(Kruh(11, 5, "Fyzika")) == "11"
    assert format_kruh_label(Kruh(1102, 5, "Fyzika")) == "11[c]"


def test_compute_kruhy_split_produces_friends():
    kruhy = [Kruh(11, 25, "Fyzika")]
    split, friends = compute_kruhy_split(kruhy, team_size=10)
    assert [k.id for k in split] == [1100, 1101, 1102]
    assert [k.count for k in split] == [10, 10, 5]
    assert friends == [split]


def test_compute_distributions_small_synthetic_case():
    # A tiny, fast-to-solve scenario -- not the full example config, which can take
    # up to SOLVER_TIME_LIMIT seconds per (count, size) combination.
    from survivor_app.core.config import OborConfig, SubteamConfig

    config = Config.empty()
    config.possible_teams_counts = [2]
    config.possible_teams_sizes = [5]
    config.teams_names = ["Team A", "Team B"]
    config.subteams = [SubteamConfig("1", "#ffffff"), SubteamConfig("2", "#000000")]
    config.obory = [OborConfig("Fyzika", [11, 12]), OborConfig("Informatika", [21])]

    counts = {11: 4, 12: 4, 21: 4}

    solutions = compute_distributions(counts, config)
    assert len(solutions) == 1
    solution = solutions[0]
    assert solution.status in (SolutionStatus.FEASIBLE, SolutionStatus.OPTIMAL)
    assigned_ids = {kruh.id for team in solution.distribution for subteam in team for kruh in subteam}
    assert assigned_ids == {11, 12, 21}


def test_compute_timetable_layout_matches_example_shape():
    config = load_config(os.path.join(EXAMPLE_DIR, "config.json"))
    layout = compute_timetable_layout(config)

    assert layout.num_rows == 1 + config.activities_count * 4
    assert layout.num_cols == 1 + config.activities_count

    activity_cells = [c for c in layout.cells if c.kind == "activity"]
    assert len(activity_cells) == config.activities_count

    time_block_cells = [c for c in layout.cells if c.kind == "time_block"]
    assert [c.text for c in time_block_cells] == ["13:00", "13:15", "13:30", "13:45", "14:00",
                                                    "14:15", "14:30", "14:45", "15:00", "15:15"]


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"ok  {name}")
    print("All core tests passed.")
