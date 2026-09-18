from __future__ import annotations

import itertools
import time
from collections import defaultdict
from typing import Callable

from ortools.sat.python import cp_model

from .config import Config
from .models import Kruh, ProgressEvent, Solution, SolutionStatus

SOLVER_TIME_LIMIT = 30  # seconds, per (num_teams, max_subteam_size) combination


class _CancelCallback(cp_model.CpSolverSolutionCallback):
    """Lets a single CP-SAT solve be interrupted early via a should_cancel() poll."""

    def __init__(self, should_cancel: Callable[[], bool]):
        super().__init__()
        self._should_cancel = should_cancel

    def on_solution_callback(self) -> None:
        if self._should_cancel():
            self.stop_search()


def compute_kruhy_split(kruhy: list[Kruh], team_size: int) -> tuple[list[Kruh], list[list[Kruh]]]:
    """Split any Kruh bigger than team_size into team_size-sized synthetic sub-Kruhy.

    Sub-Kruh ids are 100*original_id + part_index. Returns the full (split) Kruh
    list plus groups of sub-Kruhy ("friends") that must land in the same Team.
    """
    kruhy_split = []
    friends = []
    for kruh in kruhy:
        if kruh.count <= team_size:
            kruhy_split.append(kruh)
            continue

        full_count, remainder = divmod(kruh.count, team_size)
        splits = [Kruh(100 * kruh.id + i, team_size, kruh.obor) for i in range(full_count)]
        if remainder > 0:
            splits.append(Kruh(100 * kruh.id + full_count, remainder, kruh.obor))
        kruhy_split.extend(splits)
        friends.append(splits)

    return kruhy_split, friends


def compute_distributions(
    counts: dict[int, int],
    config: Config,
    progress_callback: Callable[[ProgressEvent], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
) -> list[Solution]:
    """Try every (Possible Teams count) x (Possible Teams size) combination from config.

    Returns one Solution per combination attempted (feasible, optimal, or not).
    """
    kruhy_all = [
        Kruh(kruh_id, counts[kruh_id], obor.name)
        for obor in config.obory
        for kruh_id in obor.kruhy
        if kruh_id in counts
    ]

    combos = list(itertools.product(config.possible_teams_counts, config.possible_teams_sizes))
    solutions: list[Solution] = []

    for combo_index, (num_teams, max_subteam_size) in enumerate(combos):
        if should_cancel is not None and should_cancel():
            break

        kruhy_split, kruhy_friends = compute_kruhy_split(kruhy_all, max_subteam_size)

        if progress_callback is not None:
            progress_callback(
                ProgressEvent(combo_index, len(combos), num_teams, max_subteam_size, "started")
            )

        t_start = time.time()
        solution = compute_teams_distribution(
            num_teams,
            max_subteam_size,
            kruhy_split,
            kruhy_friends,
            config.subteams_count,
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


def compute_teams_distribution(
    num_teams: int,
    max_subteam_size: int,
    kruhy: list[Kruh],
    kruhy_friends: list[list[Kruh]],
    num_subteams: int,
    should_cancel: Callable[[], bool] | None = None,
) -> Solution:
    # Obor domain is derived from the Kruhy actually being solved for, rather than
    # a fixed enum, so Obor names are entirely config-driven.
    obor_names = sorted(set(kruh.obor for kruh in kruhy))
    obor_mapping = {i: obor for i, obor in enumerate(obor_names)} | {
        obor: i for i, obor in enumerate(obor_names)
    }

    # Build model ----------------------------------------------------------------
    model = cp_model.CpModel()

    dom_teams = cp_model.Domain.from_values(list(range(num_teams)))
    dom_subteams = cp_model.Domain.from_values(list(range(num_subteams)))

    lst_kruhy = list(range(len(kruhy)))
    lst_teams = list(range(num_teams))
    lst_subteams = list(range(num_subteams))
    lst_obory = list(range(len(obor_names)))

    # Variables
    vs_kruh_team = {}
    vs_kruh_subteam = {}
    as_kruh_team = {}
    as_kruh_subteam = {}
    as_kruh_team_subteam = {}
    vs_team_used = {}
    vs_team_subteam_used = {}
    vs_kruh_order = {}
    as_team_obor = {}
    for kruh in kruhy:
        v_kruh_team = model.new_int_var_from_domain(dom_teams, f"KruhTeam[{kruh.id}]")
        vs_kruh_team[kruh.id] = v_kruh_team
        v_kruh_subteam = model.new_int_var_from_domain(dom_subteams, f"KruhSubteam[{kruh.id}]")
        vs_kruh_subteam[kruh.id] = v_kruh_subteam

        v_kruh_order = model.new_int_var(0, num_teams * num_subteams - 1, f"I@KruhTeamSubteam[{kruh.id}]")
        vs_kruh_order[kruh.id] = v_kruh_order

        for subteam in lst_subteams:
            a_kruh_subteam = model.new_bool_var(f"@KruhSubteam[{kruh.id},{subteam}]")
            as_kruh_subteam[kruh.id, subteam] = a_kruh_subteam

        for team in lst_teams:
            a_kruh_team = model.new_bool_var(f"@KruhTeam[{kruh.id},{team}]")
            as_kruh_team[kruh.id, team] = a_kruh_team

        for team in lst_teams:
            for subteam in lst_subteams:
                a_kruh_team_subteam = model.new_bool_var(f"@KruhTeamSubteam[{kruh.id},{team},{subteam}]")
                as_kruh_team_subteam[kruh.id, team, subteam] = a_kruh_team_subteam

    for team in lst_teams:
        v_team_used = model.new_bool_var(f"TeamUsed[{team}]")
        vs_team_used[team] = v_team_used

        for obor in lst_obory:
            a_team_obor = model.new_bool_var(f"TeamObor[{team},{obor}]")
            as_team_obor[team, obor] = a_team_obor

        for subteam in lst_subteams:
            v_team_subteam_used = model.new_bool_var(f"TeamSubteamUsed[{team},{subteam}]")
            vs_team_subteam_used[team, subteam] = v_team_subteam_used

    # Variables Constraints
    for kruh in kruhy:
        # - KruhTeam sets @KruhTeam
        model.add_element(
            vs_kruh_team[kruh.id],
            [as_kruh_team[kruh.id, team] for team in lst_teams],
            1,
        )
        # - Exactly one @KruhTeam
        model.add_exactly_one([as_kruh_team[kruh.id, team] for team in lst_teams])

        # - KruhSubteam sets @KruhSubteam
        model.add_element(
            vs_kruh_subteam[kruh.id],
            [as_kruh_subteam[kruh.id, subteam] for subteam in lst_subteams],
            1,
        )
        # - Exactly one @KruhSubteam
        model.add_exactly_one([as_kruh_subteam[kruh.id, subteam] for subteam in lst_subteams])

        # - KruhOrder definition
        model.add(vs_kruh_order[kruh.id] == ((vs_kruh_team[kruh.id] * num_subteams) + vs_kruh_subteam[kruh.id]))

        # - (KruhTeam, KruhSubTeam) sets @KruhTeamSubteam
        model.add_element(
            cp_model.LinearExpr.affine(vs_kruh_order[kruh.id], 1, 0),
            [as_kruh_team_subteam[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams],
            1,
        )
        # - Exactly one @KruhTeamSubteam
        model.add_exactly_one(
            [as_kruh_team_subteam[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams]
        )

    for team in lst_teams:
        # - TeamUsed when Team is used
        model.add_max_equality(vs_team_used[team], [as_kruh_team[kruh.id, team] for kruh in kruhy])

        # - Team has its Obory
        for obor in lst_obory:
            model.add_max_equality(
                as_team_obor[team, obor],
                [as_kruh_team[kruh.id, team] for kruh in kruhy if obor_mapping[kruh.obor] == obor] + [0],
            )

        # TeamSubteamUsed when Team-Subteam is used
        for subteam in lst_subteams:
            model.add_max_equality(
                vs_team_subteam_used[team, subteam],
                [as_kruh_team_subteam[kruh.id, team, subteam] for kruh in kruhy],
            )

    # Constraints
    # - Team size must not exceed max_team_size
    for team in lst_teams:
        for subteam in lst_subteams:
            expr_subteam_size = cp_model.LinearExpr.sum(
                [kruh.count * as_kruh_team_subteam[kruh.id, team, subteam] for kruh in kruhy]
            )
            model.add(expr_subteam_size <= max_subteam_size)

    # - Friends must be in a same Team
    for friends in kruhy_friends:
        for friend1, friend2 in zip(friends, friends[1:]):
            model.add(vs_kruh_team[friend1.id] == vs_kruh_team[friend2.id])

    # - Symmetry breaking teams used consecutively
    for team1, team2 in zip(lst_teams, lst_teams[1:]):
        model.add_implication(vs_team_used[team2], vs_team_used[team1])

    # - Symmetry breaking ordering of Kruhy Subteams
    for team in lst_teams:
        for i_kruh1 in lst_kruhy:
            for i_kruh2 in lst_kruhy[i_kruh1 + 1:]:
                kruh1, kruh2 = kruhy[i_kruh1], kruhy[i_kruh2]
                model.add(vs_kruh_subteam[kruh1.id] <= vs_kruh_subteam[kruh2.id]).only_enforce_if(
                    as_kruh_team[kruh1.id, team], as_kruh_team[kruh2.id, team]
                )

    # Objective
    # - Minimize used number of teams
    expression_used_team_count = cp_model.LinearExpr.sum([vs_team_used[team] for team in lst_teams])

    # - Minimize team Obory
    expression_team_obory_sum = cp_model.LinearExpr.sum(
        [as_team_obor[team, obor] for team in lst_teams for obor in lst_obory]
    )

    model.minimize(expression_used_team_count + expression_team_obory_sum)

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
        teams = defaultdict(lambda: defaultdict(list))
        for kruh in kruhy:
            teams[solver.value(vs_kruh_team[kruh.id])][solver.value(vs_kruh_subteam[kruh.id])].append(kruh)

        distribution = [
            [[kruh for kruh in kruhs] for subteam, kruhs in sorted(subteams.items())]
            for team, subteams in sorted(teams.items())
        ]

    return Solution(num_teams, max_subteam_size, SolutionStatus(status), distribution)
