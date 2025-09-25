from __future__ import annotations
import argparse
from collections import defaultdict
from dataclasses import dataclass
from enum import Enum, StrEnum, auto
import itertools
import json
import os
import sys
import time

from ortools.sat.python import cp_model
import xlsxwriter
import xlsxwriter.worksheet
import xlsxwriter.exceptions


SOLVER_TIME_LIMIT = 30  # seconds


parser = argparse.ArgumentParser()
parser.add_argument("--config", type=str, default="config.json")
parser.add_argument("--counts", type=str, default="counts.json")
parser.add_argument("--output", type=str, default="distributions.xlsx")


class Obor(StrEnum):
    FYZIKA = "Fyzika"
    MATEMATICKE_MODELOVANI = "Matematické Modelování"
    INFORMATIKA = "Informatika"
    OBECNA_MATEMATIKA_MIT = "Obecná Matematika, MIT"
    FINANCNI_MATEMATIKA = "Finanční Matematika"
    UCITELSTVI = "Učitelství"


@dataclass
class Kruh:
    id: int
    count: int
    obor: Obor


@dataclass
class Solution:
    class Status(Enum):
        UNKNOWN = 0
        INFEASIBLE = 3
        FEASIBLE = 2
        OPTIMAL = 4

    num_teams: int
    max_subteam_size: int
    status: Status
    distribution: T_Distribution
    time: float = None

T_Distribution = list[list[list[Kruh]]]

# UTILS ================================================================================================================

def read_config(config_file) -> dict:
    with open(config_file, "r", encoding="utf8") as file:
        config = json.load(file)
    return config


def read_counts(counts_file) -> dict[int, int]:
    with open(counts_file, "r") as file:
        counts = json.load(file)
    counts = {int(k): v for k, v in counts.items()}
    return counts

# SOLVING ==============================================================================================================

def compute_distributions(counts: dict[int, int], config: dict) -> list[Solution]:
    def compute_kruhy_split(team_size: int):
        kruhy_split = []
        friends = []
        for kruh in kruhy:
            if kruh.count <= team_size:
                kruhy_split.append(kruh)
            else:
                full_count, remainder = divmod(kruh.count, team_size)
                splits = []
                for i_full in range(full_count):
                    splits.append(Kruh(100*kruh.id + i_full, team_size, kruh.obor))
                if remainder > 0:
                    splits.append(Kruh(100*kruh.id + i_full + 1, remainder, kruh.obor))
                kruhy_split.extend(splits)
                friends.append(splits)

        return kruhy_split, friends

    kruhy = [Kruh(kruh, counts[kruh], Obor(obor["Name"]))
             for obor in config["Obory"] for kruh in obor["Kruhy"]
             if kruh in counts]

    possible_nums_teams = config["Possible Teams counts"]
    possible_team_sizes = config["Possible Teams sizes"]

    solutions = []
    for properties in itertools.product(possible_nums_teams, possible_team_sizes):
        num_teams, max_subteam_size = properties

        kruhy_split, kruhy_friends = compute_kruhy_split(max_subteam_size)
        print(f"Computing solution for #Teams={num_teams}, MaxSubteamSize={max_subteam_size}")

        t_start = time.time()
        solution = compute_teams_distribution(num_teams, max_subteam_size, kruhy_split, kruhy_friends, config)
        t_end = time.time()
        solution.time = t_end - t_start

        print(f"> Computed in {solution.time:.2f}s. Result: {solution.status.name}")
        solutions.append(solution)

    return solutions


def compute_teams_distribution(num_teams: int, max_subteam_size: int, kruhy: list[Kruh], kruhy_friends: list[list[Kruh]], config: dict) -> Solution:
    num_subteams = config["Subteams count"]

    # Build model ------------------------------------------------------------------------------------------------------
    model = cp_model.CpModel()

    dom_teams = cp_model.Domain.from_values(list(range(num_teams)))
    dom_subteams = cp_model.Domain.from_values(list(range(num_subteams)))
    dom_obory = cp_model.Domain.from_values(list(range(len(Obor))))

    lst_kruhy = list(range(len(kruhy)))
    lst_teams = list(range(num_teams))
    lst_subteams = list(range(num_subteams))
    lst_obory = list(range(len(Obor)))

    obor_mapping = {i: obor for i, obor in enumerate(Obor)} | {obor: i for i, obor in enumerate(Obor)}

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

        v_kruh_order = model.new_int_var(0, num_teams * num_subteams -1, f"I@KruhTeamSubteam[{kruh.id}]")
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
            1
        )
        # - Exactly one @KruhTeam
        model.add_exactly_one(
            [as_kruh_team[kruh.id, team] for team in lst_teams]
        )

        # - KruhSubteam sets @KruhSubteam
        model.add_element(
            vs_kruh_subteam[kruh.id],
            [as_kruh_subteam[kruh.id, subteam] for subteam in lst_subteams],
            1
        )
        # - Exactly one @KruhSubteam
        model.add_exactly_one(
            [as_kruh_subteam[kruh.id, subteam] for subteam in lst_subteams]
        )

        # - KruhOrder definition
        model.add(vs_kruh_order[kruh.id] == ((vs_kruh_team[kruh.id] * num_subteams) + vs_kruh_subteam[kruh.id]))

        # - (KruhTeam, KruhSubTeam) sets @KruhTeamSubteam
        model.add_element(
            cp_model.LinearExpr.affine(vs_kruh_order[kruh.id], 1, 0),
            [as_kruh_team_subteam[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams],
            1
        )
        # - Exactly one @KruhTeamSubteam
        model.add_exactly_one(
            [as_kruh_team_subteam[kruh.id, team, subteam] for team in lst_teams for subteam in lst_subteams]
        )

    for team in lst_teams:
        # - TeamUsed when Team is used
        model.add_max_equality(
            vs_team_used[team],
            [as_kruh_team[kruh.id, team] for kruh in kruhy]
        )

        # - Team has its Obory
        for obor in lst_obory:
            model.add_max_equality(
                as_team_obor[team, obor],
                [as_kruh_team[kruh.id, team] for kruh in kruhy if obor_mapping[kruh.obor] == obor] + [0]
            )

        # TeamSubteamUsed when Team-Subteam is used
        for subteam in lst_subteams:
            model.add_max_equality(
                vs_team_subteam_used[team, subteam],
                [as_kruh_team_subteam[kruh.id, team, subteam] for kruh in kruhy]
            )

    # Constraints
    # - Team size must not exceed max_team_size
    for team in lst_teams:
        for subteam in lst_subteams:
            expr_subteam_size = cp_model.LinearExpr.sum(
                [kruh.count * as_kruh_team_subteam[kruh.id, team, subteam] for kruh in kruhy]
            )
            model.add(
                expr_subteam_size <= max_subteam_size
            )

    # - Friends must be in a same Team
    for friends in kruhy_friends:
        for friend1, friend2 in zip(friends, friends[1:]):
            model.add(
                vs_kruh_team[friend1.id] == vs_kruh_team[friend2.id]
            )

    # - Symmetry breaking teams used consecutively
    for team1, team2 in zip(lst_teams, lst_teams[1:]):
        model.add_implication(vs_team_used[team2], vs_team_used[team1])

    # - Symmetry breaking ordering of Kruhy Subteams
    for team in lst_teams:
        for i_kruh1 in lst_kruhy:
            for i_kruh2 in lst_kruhy[i_kruh1+1:]:
                kruh1, kruh2 = kruhy[i_kruh1], kruhy[i_kruh2]
                model.add(
                    vs_kruh_subteam[kruh1.id] <= vs_kruh_subteam[kruh2.id]
                ).only_enforce_if(as_kruh_team[kruh1.id, team], as_kruh_team[kruh2.id, team])

    # - All Subteams are used unless last team
    # for team1, team2 in zip(lst_teams, lst_teams[1:]):
    #     model.add_bool_and(
    #         *[vs_team_subteam_used[team1, subteam] for subteam in lst_subteams]
    #     ).only_enforce_if(vs_team_used[team2])

    # Objective
    # - Minimize used number of teams
    expression_used_team_count = cp_model.LinearExpr.sum(
        [vs_team_used[team] for team in lst_teams]
    )

    # - Minimize team Obory
    expression_team_obory_sum = cp_model.LinearExpr.sum(
        [as_team_obor[team, obor] for team in lst_teams for obor in lst_obory]
    )

    # - Minimize small Subteams
    # expression_subteam_size_diff = cp_model.LinearExpr.sum([
    #     (max_subteam_size - cp_model.LinearExpr.sum(
    #         [kruh.count * as_kruh_team_subteam[kruh.id, team, subteam] for kruh in kruhy]
    #     )) * subteam
    #     for team in lst_teams for subteam in lst_subteams
    # ])

    model.minimize(
        expression_used_team_count
        + expression_team_obory_sum
        # + expression_subteam_size_diff
    )

    # Solve ------------------------------------------------------------------------------------------------------------

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = SOLVER_TIME_LIMIT

    status = solver.solve(model)

    # Solution ---------------------------------------------------------------------------------------------------------

    distribution = []
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        teams = defaultdict(lambda: defaultdict(list))
        for kruh in kruhy:
            teams[solver.value(vs_kruh_team[kruh.id])][solver.value(vs_kruh_subteam[kruh.id])].append(kruh)

        distribution = [[[kruh for kruh in kruhs]
                         for subteam, kruhs in sorted(subteams.items())]
                        for team, subteams in sorted(teams.items())]

    solution = Solution(num_teams, max_subteam_size, Solution.Status(status), distribution)
    return solution

# OUTPUT ===============================================================================================================

class Format:
    class Obor:
        _common = {
            "border": 1,
        }
        _colors = {
            Obor.FYZIKA: "#37c4e5",
            Obor.MATEMATICKE_MODELOVANI: "#f08baa",
            Obor.INFORMATIKA: "#8ac75a",
            Obor.OBECNA_MATEMATIKA_MIT: "#f08baa",
            Obor.FINANCNI_MATEMATIKA: "#f08baa",
            Obor.UCITELSTVI: "#f5bf69",
        }

    _team = {
        "align": "center",
        "valign": "vcenter",
        "border": 1,
        "top": 2,
        "bottom": 2,
    }

    _count = {
        "border": 1,
    }

    @staticmethod
    def init(workbook: xlsxwriter.Workbook):
        def make_obor(obor: Obor):
            return Format.Obor._common | {"bg_color": Format.Obor._colors[obor]}

        Format.Obor.dictionary = {obor: workbook.add_format(make_obor(obor)) for obor in Obor}
        Format.team = workbook.add_format(Format._team)
        Format.count = workbook.add_format(Format._count)

    def format_kruh(kruh):
        if kruh.id < 100:
            return str(kruh.id)
        else:
            kruh_id, kruh_part = divmod(kruh.id, 100)
            kruh_part = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'][kruh_part]
            return f"{kruh_id}[{kruh_part}]"


def write_kruhy_table(worksheet: xlsxwriter.worksheet.Worksheet, solution: Solution):
    kruhy = [kruh for team in solution.distribution for subteam in team for kruh in subteam]

    worksheet.write(0, 0, "Kruh")
    worksheet.write(0, 1, "Size")
    kruhy_sorted = sorted(kruhy, key=lambda k: k.id)
    for i_kruh, kruh in enumerate(kruhy_sorted):
        worksheet.write_string(1 + i_kruh, 0,
                               Format.format_kruh(kruh))
        worksheet.write_number(1 + i_kruh, 1,
                               kruh.count)
    worksheet.write_number(1 + i_kruh + 1, 1, 0)


def write_solution(solution: Solution, worksheet: xlsxwriter.worksheet.Worksheet, config: dict):
    team_names = config["Teams names"]

    distribution = solution.distribution
    num_subteams = config["Subteams count"]
    num_kruhy = sum(1 for team in distribution for subteam in team for kruh in subteam)

    for i_team, team in enumerate(distribution):
        row_team = i_team*num_subteams

        worksheet.merge_range(row_team, 0,
                              row_team + num_subteams -1, 0,
                              team_names[i_team],
                              cell_format=Format.team)

        for i_subteam, subteam in enumerate(team):
            row_subteam = row_team + i_subteam
            worksheet.write(
                row_subteam, 1,
                f"=SUM(XLOOKUP(C{1+row_subteam}:Z{1+row_subteam}, \'Kruhy-{solution.num_teams}_{solution.max_subteam_size}\'!A1:A{1+num_kruhy+1}, \'Kruhy-{solution.num_teams}_{solution.max_subteam_size}\'!B1:B{1+num_kruhy+1}))",
                Format.count
            )
            for i_kruh, kruh in enumerate(subteam):
                worksheet.write_string(row_subteam, 2 + i_kruh,
                                       Format.format_kruh(kruh),
                                       Format.Obor.dictionary[kruh.obor])

        for i_subteam in range(i_subteam+1, num_subteams):
            row_subteam = row_team + i_subteam
            worksheet.write(
                row_subteam, 1,
                f"=SUM(XLOOKUP(C{1+row_subteam}:Z{1+row_subteam}, \'Kruhy-{solution.num_teams}_{solution.max_subteam_size}\'!A1:A{1+num_kruhy+1}, \'Kruhy-{solution.num_teams}_{solution.max_subteam_size}\'!B1:B{1+num_kruhy+1}))",
                Format.count
            )


def write_solutions(filename: str, solutions: list[Solution], config: dict):
    workbook = xlsxwriter.Workbook(filename)
    Format.init(workbook)

    for solution in solutions:
        if not solution.status in {Solution.Status.FEASIBLE, Solution.Status.OPTIMAL}:
            continue

        kruhy_table = workbook.add_worksheet(f"Kruhy-{solution.num_teams}_{solution.max_subteam_size}")
        write_kruhy_table(kruhy_table, solution)

        worksheet = workbook.add_worksheet(f"Teams-{solution.num_teams}_{solution.max_subteam_size}")
        write_solution(solution, worksheet, config)

    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError:
        print(f"[{filename}] cannot be written. It is probably open in another program.")
        sys.exit(1)

# MAIN =================================================================================================================

def main(args: argparse.Namespace):
    config = read_config(args.config)
    counts = read_counts(args.counts)

    solutions = compute_distributions(counts, config)
    write_solutions(args.output, solutions, config)


if __name__ == "__main__":
    args = parser.parse_args()
    # args = parser.parse_args(["--counts", "test_counts.json"])
    main(args)
