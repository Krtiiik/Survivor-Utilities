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
