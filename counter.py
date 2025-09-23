import argparse
import json
import os

import tabulate


parser = argparse.ArgumentParser()
parser.add_argument("file", type=str, nargs="?", default="counts.json")


def load_data(filename: str) -> dict[int, int]:
    if not os.path.exists(filename):
        return {}

    with open(filename, 'r') as file:
        data = json.load(file)
        return {int(k): v for k, v in data.items()}


def save_data(data, filename: str):
    with open(filename, "w") as file:
        json.dump(data, file, indent=4)


def print_data(data: dict[int, int], history: list[int]):
    os.system('cls' if os.name == 'nt' else 'clear')

    table_headers = ["Group Number", "Visitor Count"]
    table_data = sorted([[group, count] for group, count in data.items()])
    table_data += [["Total", sum(data.values())]]

    print(tabulate.tabulate(table_data, headers=table_headers, tablefmt="simple"))

    print("\nRecent Changes:")
    print(" ".join(map(str, history[-15:])))


def increment(group_num: int, data: dict[int, int], history: list[int]):
    if group_num in data:
        data[group_num] += 1
    else:
        data[group_num] = 1

    history.append(group_num)


def undo_last_increment(data: dict[int, int], history: list[int]):
    if history:
        last_group = history.pop()
        if data[last_group] > 1:
            data[last_group] -= 1
        else:
            del data[last_group]


def input_loop(data: dict[int, int], history: list[int]):
    print_data(data, history)

    user_input = input("Enter group number to increment or '-' to undo: ").strip()

    if user_input == "-":
        undo_last_increment(data, history)
        return True
    elif user_input.isdigit():
        group_num = int(user_input)
        increment(group_num, data, history)
        return True

    return False


def main(args: argparse.Namespace):
    data = load_data(args.file)
    history = []

    try:
        while True:
            did_change = input_loop(data, history)
            if did_change:
                save_data(data, args.file)
    except:
        save_data(data, args.file)


if __name__ == "__main__":
    args = parser.parse_args()
    main(args)
