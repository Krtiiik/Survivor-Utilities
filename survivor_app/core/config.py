from __future__ import annotations

import json
from dataclasses import dataclass, field

ACTIVITY_TYPES = ("all", "split", "rest")

# Default for "Min split part size", used when a config predates that key.
DEFAULT_MIN_SPLIT_PART_SIZE = 3
# Default for "Solver time limit" (seconds per solve), same.
DEFAULT_SOLVER_TIME_LIMIT = 30


@dataclass
class OborConfig:
    name: str
    kruhy: list[int] = field(default_factory=list)
    color: str = "#ffffff"

    @staticmethod
    def from_dict(d: dict) -> "OborConfig":
        return OborConfig(
            name=d["Name"],
            kruhy=[int(k) for k in d["Kruhy"]],
            # "Color" is a newer field -- default to white for configs saved before
            # it existed, rather than failing to load them.
            color=d.get("Color", "#ffffff"),
        )

    def to_dict(self) -> dict:
        return {"Name": self.name, "Kruhy": list(self.kruhy), "Color": self.color}


@dataclass
class SubteamConfig:
    name: str
    color: str

    @staticmethod
    def from_dict(d: dict) -> "SubteamConfig":
        return SubteamConfig(name=d["Name"], color=d["Color"])

    def to_dict(self) -> dict:
        return {"Name": self.name, "Color": self.color}


@dataclass
class ActivityConfig:
    name: str
    type: str  # one of ACTIVITY_TYPES

    @staticmethod
    def from_dict(d: dict) -> "ActivityConfig":
        return ActivityConfig(name=d["Name"], type=d["Type"])

    def to_dict(self) -> dict:
        return {"Name": self.name, "Type": self.type}


@dataclass
class TimeConfig:
    start: str  # "HH:MM"
    activity_duration: str  # "HH:MM"

    @staticmethod
    def from_dict(d: dict) -> "TimeConfig":
        return TimeConfig(start=d["Start"], activity_duration=d["Activity duration"])

    def to_dict(self) -> dict:
        return {"Start": self.start, "Activity duration": self.activity_duration}


@dataclass
class Config:
    possible_teams_sizes: list[int]
    teams_names: list[str]
    subteams: list[SubteamConfig]
    activities: list[ActivityConfig]
    time: TimeConfig
    obory: list[OborConfig]
    # Number of Teams the Timesheet renders (the first names of teams_names). Only
    # used by the Timesheet -- the Distribution solver uses teams_count instead.
    timesheet_teams_count: int
    # Smallest number of people a Kruh too large for one Subteam may be split into
    # per Subteam (the solver clamps it down if a Kruh can't be split that evenly).
    min_split_part_size: int = DEFAULT_MIN_SPLIT_PART_SIZE
    # Seconds a single solve (one Possible Teams size) may take before the solver
    # stops and returns the best distribution found so far.
    solver_time_limit: int = DEFAULT_SOLVER_TIME_LIMIT

    @property
    def teams_count(self) -> int:
        """Most Teams the Distribution solver may use (it picks the fewest that work
        on its own). The Timesheet renders timesheet_teams_count Teams instead."""
        return len(self.teams_names)

    @property
    def subteams_count(self) -> int:
        return len(self.subteams)

    @property
    def activities_count(self) -> int:
        return len(self.activities)

    @staticmethod
    def from_dict(d: dict) -> "Config":
        # "Teams count"/"Possible Teams counts" from older configs are ignored: the
        # solver's Team count is now len("Teams names").
        teams_names = list(d["Teams names"])
        activities = [ActivityConfig.from_dict(a) for a in d["Activities"]]
        return Config(
            possible_teams_sizes=[int(x) for x in d["Possible Teams sizes"]],
            teams_names=teams_names,
            subteams=[SubteamConfig.from_dict(s) for s in d["Subteams"]],
            activities=activities,
            time=TimeConfig.from_dict(d["Time"]),
            obory=[OborConfig.from_dict(o) for o in d["Obory"]],
            # Newer key -- older configs default to as many Teams as the Timesheet
            # can render (one per name, at most one per Activity).
            timesheet_teams_count=int(
                d.get("Timesheet Teams count", min(len(teams_names), len(activities)))
            ),
            # Newer, optional key -- older configs fall back to the default.
            min_split_part_size=int(d.get("Min split part size", DEFAULT_MIN_SPLIT_PART_SIZE)),
            solver_time_limit=int(d.get("Solver time limit", DEFAULT_SOLVER_TIME_LIMIT)),
        )

    def to_dict(self) -> dict:
        return {
            "Possible Teams sizes": list(self.possible_teams_sizes),
            "Min split part size": self.min_split_part_size,
            "Solver time limit": self.solver_time_limit,
            "Teams names": list(self.teams_names),
            "Subteams": [s.to_dict() for s in self.subteams],
            "Timesheet Teams count": self.timesheet_teams_count,
            "Activities": [a.to_dict() for a in self.activities],
            "Time": self.time.to_dict(),
            "Obory": [o.to_dict() for o in self.obory],
        }

    @staticmethod
    def empty() -> "Config":
        return Config(
            possible_teams_sizes=[],
            teams_names=[],
            subteams=[],
            activities=[],
            time=TimeConfig(start="00:00", activity_duration="00:15"),
            obory=[],
            timesheet_teams_count=0,
        )


def load_config(filename: str) -> Config:
    with open(filename, "r", encoding="utf8") as file:
        data = json.load(file)
    return Config.from_dict(data)


def save_config(config: Config, filename: str) -> None:
    with open(filename, "w", encoding="utf8") as file:
        json.dump(config.to_dict(), file, indent=4, ensure_ascii=False)


def _is_hhmm(value: str) -> bool:
    parts = value.split(":")
    if len(parts) != 2 or not all(p.isdigit() for p in parts):
        return False
    hours, minutes = int(parts[0]), int(parts[1])
    return hours >= 0 and 0 <= minutes < 60


def validate(config: Config) -> list[str]:
    """Return a list of human-readable problems with config, empty if it's usable."""
    errors: list[str] = []

    if not config.teams_names:
        errors.append("Teams names must not be empty.")

    if not config.possible_teams_sizes:
        errors.append("Possible Teams sizes must not be empty.")
    if any(n <= 0 for n in config.possible_teams_sizes):
        errors.append("Possible Teams sizes must all be positive.")

    if config.min_split_part_size < 1:
        errors.append("Min split part size must be at least 1.")

    if config.solver_time_limit < 1:
        errors.append("Solver time limit must be at least 1 second.")

    if not config.subteams:
        errors.append("At least one Subteam must be defined.")
    subteam_names = [s.name for s in config.subteams]
    if len(subteam_names) != len(set(subteam_names)):
        errors.append("Subteam names must be unique.")

    if not config.activities:
        errors.append("At least one Activity must be defined.")

    if config.timesheet_teams_count < 1:
        errors.append("Timesheet Teams count must be at least 1.")
    if config.timesheet_teams_count > len(config.teams_names):
        errors.append(
            f"Timesheet Teams count ({config.timesheet_teams_count}) must not exceed the "
            f"number of Teams names ({len(config.teams_names)})."
        )
    if config.timesheet_teams_count > config.activities_count:
        errors.append(
            f"Timesheet Teams count ({config.timesheet_teams_count}) must not exceed the "
            f"number of Activities ({config.activities_count}): every Team starts the "
            f"event on a different Activity."
        )
    for activity in config.activities:
        if activity.type not in ACTIVITY_TYPES:
            errors.append(
                f"Activity '{activity.name}' has unrecognized Type '{activity.type}' "
                f"(must be one of {', '.join(ACTIVITY_TYPES)})."
            )

    if not _is_hhmm(config.time.start):
        errors.append(f"Time.Start must be in HH:MM format, got '{config.time.start}'.")
    if not _is_hhmm(config.time.activity_duration):
        errors.append(
            f"Time.Activity duration must be in HH:MM format, got '{config.time.activity_duration}'."
        )

    if not config.obory:
        errors.append("At least one Obor must be defined.")
    seen_kruhy: dict[int, str] = {}
    obor_names = [o.name for o in config.obory]
    if len(obor_names) != len(set(obor_names)):
        errors.append("Obor names must be unique.")
    for obor in config.obory:
        if not obor.name:
            errors.append("An Obor is missing a name.")
        for kruh_id in obor.kruhy:
            if kruh_id in seen_kruhy:
                errors.append(
                    f"Kruh {kruh_id} is assigned to both Obor '{seen_kruhy[kruh_id]}' and '{obor.name}'."
                )
            else:
                seen_kruhy[kruh_id] = obor.name

    return errors
