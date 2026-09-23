from __future__ import annotations

import time
from typing import Callable

from ortools.sat.python import cp_model

from .config import Config
from .models import (
    OBORY_WEIGHT,
    SUBTEAM_SPREAD_WEIGHT,
    TEAM_SPREAD_WEIGHT,
    Kruh,
    ProgressEvent,
    Solution,
    SolutionStatus,
)

SOLVER_TIME_LIMIT = 30  # seconds, per Possible Teams size


class _CancelCallback(cp_model.CpSolverSolutionCallback):
    """Lets a single CP-SAT solve be interrupted early via a should_cancel() poll."""

    def __init__(self, should_cancel: Callable[[], bool]):
        super().__init__()
        self._should_cancel = should_cancel

    def on_solution_callback(self) -> None:
        if self._should_cancel():
            self.stop_search()


def compute_distributions(
    counts: dict[int, int],
    config: Config,
    progress_callback: Callable[[ProgressEvent], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> list[Solution]:
    """Solve once per Possible Teams size from config, each time with up to
    config.teams_count Teams (the solver uses the fewest that work).

    Returns one Solution per size attempted (feasible, optimal, or not).
    """
    kruhy_all = [
        Kruh(kruh_id, counts[kruh_id], obor.name)
        for obor in config.obory
        for kruh_id in obor.kruhy
        if kruh_id in counts
    ]

    num_teams = config.teams_count
    combos = list(config.possible_teams_sizes)
    solutions: list[Solution] = []

    for combo_index, max_subteam_size in enumerate(combos):
        if should_cancel is not None and should_cancel():
            break

        if progress_callback is not None:
            progress_callback(
                ProgressEvent(combo_index, len(combos), num_teams, max_subteam_size, "started")
            )

        t_start = time.time()
        solution = compute_teams_distribution(
            num_teams,
            max_subteam_size,
            kruhy_all,
            config.subteams_count,
            config.min_split_part_size,
            should_cancel=should_cancel,
        )
        solution.time = time.time() - t_start

        if progress_callback is not None:
            progress_callback(
                ProgressEvent(
                    combo_index, len(combos), num_teams, max_subteam_size, "finished", solution
                )
            )

        solutions.append(solution)

    return solutions


def split_parts_count(kruh: Kruh, max_subteam_size: int) -> int:
    """Number of Subteams a Kruh must be spread across: 1 if it fits into one,
    otherwise the fewest parts of at most max_subteam_size people."""
    return max(1, -(-kruh.count // max_subteam_size))


def compute_teams_distribution(
    num_teams: int,
    max_subteam_size: int,
    kruhy: list[Kruh],
    num_subteams: int,
    min_split_part_size: int,
    should_cancel: Callable[[], bool] | None = None,
) -> Solution:
    """Assign whole Kruhy to Teams and their people to Subteams.

    A Kruh that fits into one Subteam lands in exactly one. A Kruh larger than
    max_subteam_size is spread across the fewest possible Subteams of a single Team,
    with the solver choosing how many people go into each part (each part at least
    min_split_part_size, clamped down when the Kruh can't be split that evenly).
    Parts are returned as synthetic Kruhy with ids 100*kruh_id + part_index.
    """
    # Obor domain is derived from the Kruhy actually being solved for, rather than
    # a fixed enum, so Obor names are entirely config-driven.
    obor_names = sorted(set(kruh.obor for kruh in kruhy))

    lst_teams = list(range(num_teams))
    lst_subteams = list(range(num_subteams))

    total_count = sum(kruh.count for kruh in kruhy)
    parts_count = {kruh.id: split_parts_count(kruh, max_subteam_size) for kruh in kruhy}

    # Build model ----------------------------------------------------------------
    model = cp_model.CpModel()

    # Variables
    # - @KruhTeam: Kruh is in Team
    as_kruh_team = {
        (kruh.id, team): model.new_bool_var(f"@KruhTeam[{kruh.id},{team}]")
        for kruh in kruhy
        for team in lst_teams
    }
    # - @KruhPart: Kruh has a part (possibly all of it) in Subteam of Team
    as_kruh_part = {
        (kruh.id, team, subteam): model.new_bool_var(f"@KruhPart[{kruh.id},{team},{subteam}]")
        for kruh in kruhy
        for team in lst_teams
        for subteam in lst_subteams
    }
    # - KruhPartSize: people of Kruh in Subteam of Team. A plain expression for Kruhy
    #   that fit into one Subteam, an int variable only for Kruhy that get split.
    es_kruh_part_size = {}
    for kruh in kruhy:
        for team in lst_teams:
            for subteam in lst_subteams:
                a_kruh_part = as_kruh_part[kruh.id, team, subteam]
                if parts_count[kruh.id] == 1:
                    es_kruh_part_size[kruh.id, team, subteam] = kruh.count * a_kruh_part
                else:
                    es_kruh_part_size[kruh.id, team, subteam] = model.new_int_var(
                        0, max_subteam_size, f"KruhPartSize[{kruh.id},{team},{subteam}]"
                    )

    vs_team_used = {team: model.new_bool_var(f"TeamUsed[{team}]") for team in lst_teams}
    vs_team_size = {team: model.new_int_var(0, total_count, f"TeamSize[{team}]") for team in lst_teams}
    as_team_obor = {
        (team, obor): model.new_bool_var(f"TeamObor[{team},{obor}]")
        for team in lst_teams
        for obor in obor_names
    }
    vs_team_subteam_size = {
        (team, subteam): model.new_int_var(0, max_subteam_size, f"SubteamSize[{team},{subteam}]")
        for team in lst_teams
        for subteam in lst_subteams
    }

    # Variables Constraints
    for kruh in kruhy:
        # - Exactly one @KruhTeam
        model.add_exactly_one([as_kruh_team[kruh.id, team] for team in lst_teams])

        # - Kruh has exactly as many parts as it needs, all within its own Team
        #   (at most one part per Subteam, since @KruhPart is a bool per Subteam)
        model.add(
            cp_model.LinearExpr.sum(
                [as_kruh_part[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams]
            )
            == parts_count[kruh.id]
        )
        for team in lst_teams:
            for subteam in lst_subteams:
                model.add_implication(as_kruh_part[kruh.id, team, subteam], as_kruh_team[kruh.id, team])

        # - Split Kruh: parts add up to the whole Kruh, each part at least
        #   min_split_part_size (clamped so an even-enough split always exists)
        if parts_count[kruh.id] > 1:
            min_part_size = min(min_split_part_size, kruh.count // parts_count[kruh.id])
            for team in lst_teams:
                for subteam in lst_subteams:
                    v_part_size = es_kruh_part_size[kruh.id, team, subteam]
                    a_kruh_part = as_kruh_part[kruh.id, team, subteam]
                    model.add(v_part_size >= min_part_size).only_enforce_if(a_kruh_part)
                    model.add(v_part_size == 0).only_enforce_if(a_kruh_part.Not())
            model.add(
                cp_model.LinearExpr.sum(
                    [es_kruh_part_size[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams]
                )
                == kruh.count
            )

    for team in lst_teams:
        # - TeamUsed when Team is used
        model.add_max_equality(vs_team_used[team], [as_kruh_team[kruh.id, team] for kruh in kruhy])

        # - TeamSize definition
        model.add(
            vs_team_size[team]
            == cp_model.LinearExpr.sum([kruh.count * as_kruh_team[kruh.id, team] for kruh in kruhy])
        )

        # - Team has its Obory
        for obor in obor_names:
            model.add_max_equality(
                as_team_obor[team, obor],
                [as_kruh_team[kruh.id, team] for kruh in kruhy if kruh.obor == obor] + [0],
            )

        # - SubteamSize definition (its domain enforces max_subteam_size)
        for subteam in lst_subteams:
            model.add(
                vs_team_subteam_size[team, subteam]
                == cp_model.LinearExpr.sum([es_kruh_part_size[kruh.id, team, subteam] for kruh in kruhy])
            )

    # - Symmetry breaking: Teams are interchangeable, so order them by size (which
    #   also keeps used Teams consecutive); likewise Subteams within a Team.
    for team1, team2 in zip(lst_teams, lst_teams[1:]):
        model.add(vs_team_size[team1] >= vs_team_size[team2])
        model.add_implication(vs_team_used[team2], vs_team_used[team1])
    for team in lst_teams:
        for subteam1, subteam2 in zip(lst_subteams, lst_subteams[1:]):
            model.add(vs_team_subteam_size[team, subteam1] >= vs_team_subteam_size[team, subteam2])

    # - Balance Team sizes: track the spread (max - min) of sizes across used Teams.
    # Unused Teams (size 0) are excluded from the min via a sentinel value, so an
    # unused Team never masquerades as the smallest (and therefore "best balanced") one.
    sentinel_team_size = total_count + 1
    vs_team_size_for_min = {}
    for team in lst_teams:
        v_team_size_for_min = model.new_int_var(0, sentinel_team_size, f"TeamSizeForMin[{team}]")
        vs_team_size_for_min[team] = v_team_size_for_min
        model.add(v_team_size_for_min == vs_team_size[team]).only_enforce_if(vs_team_used[team])
        model.add(v_team_size_for_min == sentinel_team_size).only_enforce_if(vs_team_used[team].Not())

    v_team_size_max = model.new_int_var(0, total_count, "TeamSizeMax")
    model.add_max_equality(v_team_size_max, [vs_team_size[team] for team in lst_teams])
    v_team_size_min = model.new_int_var(0, sentinel_team_size, "TeamSizeMin")
    model.add_min_equality(v_team_size_min, [vs_team_size_for_min[team] for team in lst_teams])

    # - Balance Subteam sizes: same spread trick, over every Subteam of every used
    # Team. An empty Subteam of a used Team counts as 0 -- it's the worst imbalance
    # there is, since that Team would have nobody there for a split Activity.
    sentinel_subteam_size = max_subteam_size + 1
    vs_team_subteam_size_for_min = {}
    for team in lst_teams:
        for subteam in lst_subteams:
            v_subteam_size_for_min = model.new_int_var(
                0, sentinel_subteam_size, f"SubteamSizeForMin[{team},{subteam}]"
            )
            vs_team_subteam_size_for_min[team, subteam] = v_subteam_size_for_min
            model.add(v_subteam_size_for_min == vs_team_subteam_size[team, subteam]).only_enforce_if(
                vs_team_used[team]
            )
            model.add(v_subteam_size_for_min == sentinel_subteam_size).only_enforce_if(
                vs_team_used[team].Not()
            )

    v_subteam_size_max = model.new_int_var(0, max_subteam_size, "SubteamSizeMax")
    model.add_max_equality(
        v_subteam_size_max, [vs_team_subteam_size[team, subteam] for team in lst_teams for subteam in lst_subteams]
    )
    v_subteam_size_min = model.new_int_var(0, sentinel_subteam_size, "SubteamSizeMin")
    model.add_min_equality(
        v_subteam_size_min,
        [vs_team_subteam_size_for_min[team, subteam] for team in lst_teams for subteam in lst_subteams],
    )

    # Objective
    # - Minimize used number of teams (strictly dominant)
    expression_used_team_count = cp_model.LinearExpr.sum([vs_team_used[team] for team in lst_teams])

    # - Then a weighted blend of team Obory against Team/Subteam size spread, so a
    #   mixed-Obor Team can be traded for noticeably more even sizes and vice versa
    #   (weights live in models.py, shared with Solution.score()).
    expression_team_obory_sum = cp_model.LinearExpr.sum(
        [as_team_obor[team, obor] for team in lst_teams for obor in obor_names]
    )
    expression_blend = (
        OBORY_WEIGHT * expression_team_obory_sum
        + TEAM_SPREAD_WEIGHT * (v_team_size_max - v_team_size_min)
        + SUBTEAM_SPREAD_WEIGHT * (v_subteam_size_max - v_subteam_size_min)
    )

    # Must exceed the largest possible value of the blend, so using one Team fewer
    # always wins over anything the blend could gain.
    blend_bound = (
        OBORY_WEIGHT * num_teams * len(obor_names)
        + TEAM_SPREAD_WEIGHT * total_count
        + SUBTEAM_SPREAD_WEIGHT * max_subteam_size
        + 1
    )
    model.minimize(expression_used_team_count * blend_bound + expression_blend)

    # Solve ------------------------------------------------------------------
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT

    if should_cancel is not None:
        status = solver.solve(model, _CancelCallback(should_cancel))
    else:
        status = solver.solve(model)

    # Solution -----------------------------------------------------------------
    distribution = []
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        teams: dict[int, list[list[Kruh]]] = {}
        for kruh in kruhy:
            team = next(team for team in lst_teams if solver.boolean_value(as_kruh_team[kruh.id, team]))
            subteams = teams.setdefault(team, [[] for _ in lst_subteams])
            part_subteams = [
                subteam for subteam in lst_subteams if solver.boolean_value(as_kruh_part[kruh.id, team, subteam])
            ]
            if len(part_subteams) == 1:
                subteams[part_subteams[0]].append(kruh)
                continue
            for part_index, subteam in enumerate(part_subteams):
                part_size = solver.value(es_kruh_part_size[kruh.id, team, subteam])
                subteams[subteam].append(Kruh(100 * kruh.id + part_index, part_size, kruh.obor))

        distribution = [subteams for team, subteams in sorted(teams.items())]

    return Solution(num_teams, max_subteam_size, SolutionStatus(status), distribution)
