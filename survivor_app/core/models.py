from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

# Kruh ids >= 100 are synthetic sub-Kruhy created when a Kruh is too large for a
# Subteam and the solver spreads it across several (see compute_teams_distribution).
# The suffix identifies which part it is.
SPLIT_SUFFIXES = ["a", "b", "c", "d", "e", "f", "g", "h"]

# Weights of the solver's objective blend (see compute_teams_distribution), shared
# with Solution.score() so candidates are ranked the same way the solver judged them.
# One extra Obor in a Team costs as much as OBORY_WEIGHT people of Subteam-size spread.
OBORY_WEIGHT = 4
TEAM_SPREAD_WEIGHT = 2
SUBTEAM_SPREAD_WEIGHT = 1


@dataclass
class Kruh:
    id: int
    count: int
    obor: str


class SolutionStatus(Enum):
    # Values match ortools.sat.python.cp_model's solver status ints, so
    # SolutionStatus(solver.solve(model)) can wrap the raw return value directly.
    UNKNOWN = 0
    INFEASIBLE = 3
    FEASIBLE = 2
    OPTIMAL = 4


# Team > Subteam > Kruhy assigned to that subteam.
T_Distribution = list[list[list[Kruh]]]


@dataclass
class Solution:
    num_teams: int
    max_subteam_size: int
    status: SolutionStatus
    distribution: T_Distribution
    time: float | None = None

    def score(self) -> int:
        """Lower is better. Mirrors the solver's objective blend (Obory per Team plus
        Team- and Subteam-size spread, weighted), recomputed from the actual assigned
        Kruhy so it stays correct after manual edits to the distribution. Empty
        Subteams of a Team that has any Kruhy count as size 0."""
        teams = [team for team in self.distribution if any(team)]
        team_sizes = [sum(kruh.count for subteam in team for kruh in subteam) for team in teams]
        subteam_sizes = [sum(kruh.count for kruh in subteam) for team in teams for subteam in team]
        team_spread = (max(team_sizes) - min(team_sizes)) if team_sizes else 0
        subteam_spread = (max(subteam_sizes) - min(subteam_sizes)) if subteam_sizes else 0
        obory = sum(len({kruh.obor for subteam in team for kruh in subteam}) for team in teams)
        return (
            OBORY_WEIGHT * obory
            + TEAM_SPREAD_WEIGHT * team_spread
            + SUBTEAM_SPREAD_WEIGHT * subteam_spread
        )


@dataclass
class ProgressEvent:
    combo_index: int
    combo_total: int
    num_teams: int
    max_subteam_size: int
    stage: str  # "started" | "finished"
    solution: Solution | None = None


def format_kruh_label(kruh: Kruh) -> str:
    """Render a Kruh id back to its human label, decoding split ids like 1104 -> "11[e]"."""
    if kruh.id < 100:
        return str(kruh.id)

    kruh_id, kruh_part = divmod(kruh.id, 100)
    if kruh_part >= len(SPLIT_SUFFIXES):
        raise ValueError(f"Kruh {kruh_id} was split into more than {len(SPLIT_SUFFIXES)} parts.")
    return f"{kruh_id}[{SPLIT_SUFFIXES[kruh_part]}]"
