"""Headless re-implementation of the original counter.py terminal loop, now backed
by survivor_app.core.counts so the same increment/decrement/save logic is shared
with the GUI's Counter screen."""
import argparse
import os

from survivor_app.core.counts import decrement, increment, load_counts, save_counts, summarize

try:
    import tabulate
except ImportError:
    tabulate = None

parser = argparse.ArgumentParser()
parser.add_argument("file", type=str, nargs="?", default="counts.json")


def print_data(counts: dict[int, int], history: list[int]) -> None:
    os.system("cls" if os.name == "nt" else "clear")

    rows, total = summarize(counts)
    table_data = [list(row) for row in rows] + [["Total", total]]
    if tabulate is not None:
        print(tabulate.tabulate(table_data, headers=["Group Number", "Visitor Count"], tablefmt="simple"))
    else:
        for group, count in table_data:
            print(f"{group}\t{count}")

    print("\nRecent Changes:")
    print(" ".join(map(str, history[-15:])))


def undo_last_increment(counts: dict[int, int], history: list[int]) -> None:
    if history:
        decrement(history.pop(), counts)


def input_loop(counts: dict[int, int], history: list[int]) -> bool:
    print_data(counts, history)

    user_input = input("Enter group number to increment or '-' to undo: ").strip()

    if user_input == "-":
        undo_last_increment(counts, history)
        return True
    elif user_input.isdigit():
        group_num = int(user_input)
        increment(group_num, counts)
        history.append(group_num)
        return True

    return False


def main(args: argparse.Namespace) -> None:
    counts = load_counts(args.file)
    history: list[int] = []

    try:
        while True:
            if input_loop(counts, history):
                save_counts(counts, args.file)
    except (KeyboardInterrupt, EOFError):
        save_counts(counts, args.file)


if __name__ == "__main__":
    main(parser.parse_args())
