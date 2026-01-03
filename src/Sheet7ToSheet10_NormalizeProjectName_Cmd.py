###############################################################
#
# Sheet7ToSheet10_NormalizeProjectName_Cmd.py
#
# Purpose:
#   Sheet7.tsv を読み込み、
#   プロジェクト名の正規化処理を行い、
#   Sheet10.tsv を出力するためのコマンドラインスクリプト。
#
# Note:
#   This is an initial placeholder file for GitHub / Codex startup.
#   Actual logic will be implemented step by step.
#
###############################################################

import re
import sys


iProjectNameColumnIndex: int = 0
iRemoveColumnIndex: int = 1
iManhourColumnIndex: int = 2


def is_blank(value: str | None) -> bool:
    if value is None:
        return True
    if value == "":
        return True
    if value.strip() == "":
        return True
    if value.lower() == "nan":
        return True
    return False


def normalize_project_name(pszSource: str) -> str:
    if pszSource.startswith("【廃番】"):
        try:
            code: str | None = None
            code_index: int = -1
            for prefix in ["J", "A", "C", "H", "M", "P"]:
                search_from: int = 0
                code_length: int = 4
                if prefix == "P":
                    code_length = 6
                while True:
                    found_index: int = pszSource.find(prefix, search_from)
                    if found_index == -1:
                        break
                    if found_index + code_length <= len(pszSource):
                        code = pszSource[found_index:found_index + code_length]
                        code_index = found_index
                        break
                    search_from = found_index + 1
                if code is not None:
                    break
            if code is None or code_index == -1:
                return pszSource
            head: str = pszSource[:code_index]
            tail: str = pszSource[code_index + len(code):]
            return code + "_" + head + tail
        except Exception:
            return pszSource

    if len(pszSource) >= 1 and pszSource[0] in ["J", "A", "C", "H", "M"]:
        if len(pszSource) >= 5:
            next_char: str = pszSource[4]
            if next_char == "【":
                return pszSource[:4] + "_" + pszSource[4:]
            if next_char == " " or next_char == "　":
                return pszSource[:4] + "_" + pszSource[5:]
        return pszSource

    if len(pszSource) >= 1 and pszSource[0] == "P":
        if len(pszSource) >= 7:
            next_char_p: str = pszSource[6]
            if next_char_p == "【":
                return pszSource[:6] + "_" + pszSource[6:]
            if next_char_p == " " or next_char_p == "　":
                return pszSource[:6] + "_" + pszSource[7:]
        return pszSource

    return pszSource


def preprocess_line_content(line_content: str) -> str:
    line_content = re.sub(
        r'^"([^"]*)\t([^"]*)"([^\r\n]*)',
        r"\1_\2\3",
        line_content,
    )
    line_content = re.sub(r"^(J\d\d\d) +", r"\1_", line_content)
    line_content = re.sub(r"([A-OQ-Z]\d\d\d)[ 　]+", r"\1_", line_content)
    line_content = re.sub(r"(P\d\d\d\d\d)[ 　]+", r"\1_", line_content)
    line_content = re.sub(r"\t[0-9]+\t", "\t", line_content)
    return line_content


def parse_manhour_to_seconds(manhour: str) -> int:
    match = re.match(r"^(\d+):([0-5]\d):([0-5]\d)$", manhour)
    if not match:
        raise ValueError(f"Invalid manhour format: {manhour}")
    hours = int(match.group(1))
    minutes = int(match.group(2))
    seconds = int(match.group(3))
    return hours * 3600 + minutes * 60 + seconds


def format_seconds_to_manhour(total_seconds: int) -> str:
    if total_seconds < 0:
        raise ValueError("Total seconds must not be negative.")
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return f"{hours}:{minutes:02d}:{seconds:02d}"


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python Sheet7ToSheet10_NormalizeProjectName_Cmd.py <input.tsv>")
        sys.exit(1)

    input_path: str = sys.argv[1]
    output_path: str = "Sheet10.tsv"
    sheet11_path: str = "Sheet11.tsv"
    sheet12_path: str = "Sheet12.tsv"
    sheet13_path: str = "Sheet13.tsv"
    if input_path == output_path:
        print("Error: input file must not be overwritten.")
        sys.exit(1)

    try:
        with open(input_path, "r", encoding="utf-8") as input_file:
            lines: list[str] = input_file.readlines()
    except FileNotFoundError:
        print("Error: input file not found.")
        sys.exit(1)
    except Exception as exc:
        print(str(exc))
        sys.exit(1)

    try:
        sheet10_rows: list[tuple[str, str]] = []
        with open(output_path, "w", encoding="utf-8") as output_file:
            for line in lines:
                line_content: str = line.rstrip("\n")
                if line_content == "":
                    output_file.write("\t\n")
                    sheet10_rows.append(("", ""))
                    continue
                line_content = preprocess_line_content(line_content)
                columns: list[str] = line_content.split("\t")
                project_name: str = ""
                manhour: str = ""
                if len(columns) > iProjectNameColumnIndex:
                    project_name = columns[iProjectNameColumnIndex]
                if len(columns) > iManhourColumnIndex:
                    manhour = columns[iManhourColumnIndex]
                elif len(columns) > iRemoveColumnIndex:
                    manhour = columns[iRemoveColumnIndex]
                if is_blank(project_name):
                    normalized_name: str = ""
                else:
                    normalized_name = normalize_project_name(project_name)
                output_file.write(normalized_name + "\t" + manhour + "\n")
                sheet10_rows.append((normalized_name, manhour))
        aggregated_seconds: dict[str, int] = {}
        aggregated_order: list[str] = []
        for project_name, manhour in sheet10_rows:
            if project_name == "" and manhour == "":
                continue
            seconds = parse_manhour_to_seconds(manhour)
            if project_name not in aggregated_seconds:
                aggregated_seconds[project_name] = 0
                aggregated_order.append(project_name)
            aggregated_seconds[project_name] += seconds
        sheet11_rows: list[tuple[str, str]] = []
        with open(sheet11_path, "w", encoding="utf-8") as sheet11_file:
            for project_name in aggregated_order:
                total_manhour = format_seconds_to_manhour(aggregated_seconds[project_name])
                sheet11_file.write(project_name + "\t" + total_manhour + "\n")
                sheet11_rows.append((project_name, total_manhour))
        sheet11_rows_sorted = sorted(
            sheet11_rows,
            key=lambda row: row[0].split("_", 1)[0],
        )
        with open(sheet12_path, "w", encoding="utf-8") as sheet12_file:
            for project_name, total_manhour in sheet11_rows_sorted:
                sheet12_file.write(project_name + "\t" + total_manhour + "\n")
        with open(sheet12_path, "r", encoding="utf-8") as sheet12_file:
            sheet12_lines: list[str] = sheet12_file.readlines()
        with open(sheet13_path, "w", encoding="utf-8") as sheet13_file:
            for line in sheet12_lines:
                line_content = line.rstrip("\n")
                if line_content == "":
                    sheet13_file.write("\t\n")
                    continue
                columns = line_content.split("\t")
                project_name = columns[0] if len(columns) > 0 else ""
                if project_name.startswith(("A", "H")):
                    continue
                sheet13_file.write(line_content + "\n")
    except Exception as exc:
        print(str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()
