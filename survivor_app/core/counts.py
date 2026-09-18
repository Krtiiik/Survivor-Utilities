from __future__ import annotations

import json
import os


def load_counts(filename: str) -> dict[int, int]:
    if not os.path.exists(filename):
        return {}

    with open(filename, "r") as file:
        data = json.load(file)
    return {int(k): v for k, v in data.items()}


def save_counts(counts: dict[int, int], filename: str) -> None:
    with open(filename, "w") as file:
        json.dump(counts, file, indent=4)


def increment(kruh_id: int, counts: dict[int, int]) -> None:
    counts[kruh_id] = counts.get(kruh_id, 0) + 1


def decrement(kruh_id: int, counts: dict[int, int]) -> None:
    """Decrement a specific Kruh's count, removing the key once it reaches 0."""
    if kruh_id not in counts:
        return
    if counts[kruh_id] > 1:
        counts[kruh_id] -= 1
    else:
        del counts[kruh_id]


def summarize(counts: dict[int, int]) -> tuple[list[tuple[int, int]], int]:
    """Return (rows sorted by Kruh id, total across all Kruhy)."""
    rows = sorted(counts.items())
    total = sum(counts.values())
    return rows, total
