from __future__ import annotations

import json

from .models import Kruh, Solution, SolutionStatus, T_Distribution


def _kruh_to_dict(kruh: Kruh) -> dict:
    return {"id": kruh.id, "count": kruh.count, "obor": kruh.obor}


def _kruh_from_dict(d: dict) -> Kruh:
    return Kruh(id=int(d["id"]), count=int(d["count"]), obor=d["obor"])


def _distribution_to_dict(distribution: T_Distribution) -> list:
    return [[[_kruh_to_dict(kruh) for kruh in subteam] for subteam in team] for team in distribution]


def _distribution_from_dict(d: list) -> T_Distribution:
    return [[[_kruh_from_dict(kruh) for kruh in subteam] for subteam in team] for team in d]


def _solution_to_dict(solution: Solution) -> dict:
    return {
        "num_teams": solution.num_teams,
        "max_subteam_size": solution.max_subteam_size,
        "status": solution.status.name,
        "distribution": _distribution_to_dict(solution.distribution),
        "time": solution.time,
    }


def _solution_from_dict(d: dict) -> Solution:
    return Solution(
        num_teams=int(d["num_teams"]),
        max_subteam_size=int(d["max_subteam_size"]),
        status=SolutionStatus[d["status"]],
        distribution=_distribution_from_dict(d["distribution"]),
        time=d.get("time"),
    )


def save_solutions(solutions: list[Solution], filename: str) -> None:
    """Persist every computed Solution, not just the one the user has selected --
    re-running the solver is expensive (up to SOLVER_TIME_LIMIT seconds per
    combination), so a later session can load this file and jump straight to
    picking a candidate instead of recomputing from scratch."""
    data = [_solution_to_dict(solution) for solution in solutions]
    with open(filename, "w", encoding="utf8") as file:
        json.dump(data, file, indent=2)


def load_solutions(filename: str) -> list[Solution]:
    with open(filename, "r", encoding="utf8") as file:
        data = json.load(file)
    return [_solution_from_dict(d) for d in data]
