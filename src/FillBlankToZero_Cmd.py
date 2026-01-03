import os
import re
import sys
from typing import List, Optional, Tuple


def print_usage() -> None:
    usage = "Usage: python FillBlankToZero_Cmd.py <input.tsv> [--header-lines N]"
    print(usage)


def is_blank(value: Optional[str]) -> bool:
    if value is None:
        return True
    if value == "":
        return True
    if value.strip() == "":
        return True
    return False


def is_time_value(value: str) -> bool:
    normalized = value.strip()
    return re.fullmatch(r"[0-9]+:[0-9]{2}:[0-9]{2}", normalized) is not None


def is_float_value(value: str) -> bool:
    normalized = value.strip()
    if "." not in normalized:
        return False
    try:
        float(normalized)
    except ValueError:
        return False
    return True


def is_int_value(value: str) -> bool:
    normalized = value.strip()
    try:
        int(normalized)
    except ValueError:
        return False
    return True


def parse_arguments(argv: List[str]):
    header_lines = 2
    positional: List[str] = []
    i = 1
    while i < len(argv):
        arg = argv[i]
        if arg == "--header-lines":
            if i + 1 >= len(argv):
                print_usage()
                return None
            try:
                header_lines = int(argv[i + 1])
            except ValueError:
                print_usage()
                return None
            if header_lines < 0:
                print_usage()
                return None
            i += 2
            continue
        positional.append(arg)
        i += 1

    if len(positional) != 1:
        print_usage()
        return None

    input_path = positional[0]
    return input_path, header_lines


def build_output_path(input_path: str) -> str:
    directory, filename = os.path.split(input_path)
    stem, ext = os.path.splitext(filename)
    if ext == "":
        ext = ".tsv"
    output_filename = f"{stem}_AppendZero{ext}"
    return os.path.join(directory, output_filename)


def determine_column_types(rows: List[List[str]], header_lines: int, max_cols: int) -> List[str]:
    data_rows = rows[header_lines:]
    types: List[str] = []

    for col_idx in range(max_cols):
        has_time = False
        has_float = False
        has_int = False

        for row in data_rows:
            value = row[col_idx] if col_idx < len(row) else None
            if is_blank(value):
                continue
            if is_time_value(value):
                has_time = True
                break
            if is_float_value(value):
                has_float = True
            elif is_int_value(value):
                has_int = True

        if has_time:
            types.append("time")
        elif has_float:
            types.append("float")
        elif has_int:
            types.append("int")
        else:
            types.append("other")

    return types


def fill_row(row: List[str], column_types: List[str]) -> List[str]:
    max_cols = len(column_types)
    filled: List[str] = []
    for idx in range(max_cols):
        value = row[idx] if idx < len(row) else None
        if is_blank(value):
            col_type = column_types[idx]
            if col_type == "time":
                filled.append("0:00:00")
            elif col_type == "float":
                filled.append("0.0")
            else:
                filled.append("0")
        else:
            filled.append(value)
    return filled


def load_rows(input_path: str) -> Tuple[List[List[str]], List[str]]:
    rows: List[List[str]] = []
    raw_lines: List[str] = []
    with open(input_path, "r", encoding="utf-8", newline="") as infile:
        for line in infile:
            raw_lines.append(line)
            stripped = line.rstrip("\n").rstrip("\r")
            rows.append(stripped.split("\t"))
    return rows, raw_lines


def write_output(output_path: str, header_lines: int, raw_lines: List[str], filled_rows: List[List[str]]) -> None:
    with open(output_path, "w", encoding="utf-8", newline="") as outfile:
        for i in range(min(header_lines, len(raw_lines))):
            outfile.write(raw_lines[i])
        for row in filled_rows:
            outfile.write("\t".join(row) + "\n")


def main(argv: List[str]) -> int:
    parsed = parse_arguments(argv)
    if parsed is None:
        return 1

    input_path, header_lines = parsed
    output_path = build_output_path(input_path)

    if not os.path.exists(input_path):
        print(f"Input file not found: {input_path}")
        return 1

    try:
        rows, raw_lines = load_rows(input_path)
        if not rows:
            max_cols = 0
        else:
            max_cols = max(len(row) for row in rows)

        column_types = determine_column_types(rows, header_lines, max_cols)

        data_rows = rows[header_lines:]
        filled_data_rows = [fill_row(row, column_types) for row in data_rows]

        write_output(output_path, header_lines, raw_lines, filled_data_rows)
    except Exception as exc:  # noqa: BLE001
        print(exc)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
