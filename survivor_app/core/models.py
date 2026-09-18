from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

# Kruh ids >= 100 are synthetic sub-Kruhy created when a Kruh is too large for a
# Subteam (see compute_kruhy_split). The suffix identifies which part it is.
SPLIT_SUFFIXES = ["a", "b", "c", "d", "e", "f", "g", "h"]


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

    def size_balance_score(self) -> int:
        """Lower is more evenly sized. Mirrors the solver's own balance tie-breaker
        (2 * Team-size spread + Subteam-size spread), recomputed from the actual
        assigned Kruhy so it stays correct after manual edits to the distribution."""
        team_sizes = [sum(kruh.count for subteam in team for kruh in subteam) for team in self.distribution]
        subteam_sizes = [sum(kruh.count for kruh in subteam) for team in self.distribution for subteam in team]
        team_spread = (max(team_sizes) - min(team_sizes)) if team_sizes else 0
        subteam_spread = (max(subteam_sizes) - min(subteam_sizes)) if subteam_sizes else 0
        return 2 * team_spread + subteam_spread


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
