"""CLI wrapper for the distribution solver, backed by survivor_app.core."""
import argparse
import sys

from survivor_app.core.config import load_config
from survivor_app.core.counts import load_counts
from survivor_app.core.distribute import compute_distributions
from survivor_app.core.errors import FileWriteError
from survivor_app.core.excel_export import export_distribution
from survivor_app.core.models import ProgressEvent, SolutionStatus

parser = argparse.ArgumentParser()
parser.add_argument("--config", type=str, default="config.json")
parser.add_argument("--counts", type=str, default="counts.json")
parser.add_argument("--output", type=str, default="distributions.xlsx")


def _print_progress(event: ProgressEvent) -> None:
    if event.stage == "started":
        print(f"Computing solution for #Teams={event.num_teams}, MaxSubteamSize={event.max_subteam_size}")
    else:
        print(f"> Computed in {event.solution.time:.2f}s. Result: {event.solution.status.name}")


def main(args: argparse.Namespace) -> None:
    config = load_config(args.config)
    counts = load_counts(args.counts)

    solutions = compute_distributions(counts, config, progress_callback=_print_progress)

    best = next(
        (s for s in solutions if s.status in (SolutionStatus.FEASIBLE, SolutionStatus.OPTIMAL)),
        None,
    )
    if best is None:
        print("No feasible distribution was found for any configured combination.")
        sys.exit(1)

    try:
        export_distribution(args.output, best.distribution, config, best.max_subteam_size)
    except FileWriteError as error:
        print(str(error))
        sys.exit(1)


if __name__ == "__main__":
    main(parser.parse_args())
