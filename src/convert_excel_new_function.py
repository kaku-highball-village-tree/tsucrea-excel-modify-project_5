import os
import sys
import re
from typing import List, Tuple

import pandas as pd


objTimePattern: re.Pattern[str] = re.compile(r"^\d+:\d{2}(:\d{2})?$")
fTimeLikeRatioThreshold: float = 0.6


def _output_argument_error() -> None:
    message_lines: List[str] = [
        "Error: input TSV file path is not specified (insufficient arguments).",
        "Usage: python convert_excel_new_function.py <input_tsv_file_path>",
        "Example: python convert_excel_new_function.py C:\\Data\\Staff_List_Formula.tsv",
    ]
    for line in message_lines:
        print(line)
    error_file_path: str = "convert_excel_new_function_error_argument.tsv"
    with open(error_file_path, "w", encoding="utf-8", newline="") as error_file:
        error_file.write("\n".join(message_lines))


def _output_missing_file_error(input_path: str) -> None:
    print(f"Error: input TSV file not found. Path = {input_path}")
    base_name: str = os.path.basename(input_path)
    name_without_ext: str = os.path.splitext(base_name)[0]
    directory: str = os.path.dirname(os.path.abspath(input_path))
    error_file_name: str = f"{name_without_ext}_error.tsv"
    error_file_path: str = os.path.join(directory, error_file_name)
    with open(error_file_path, "w", encoding="utf-8", newline="") as error_file:
        error_file.write(f"Error: input TSV file not found. Path = {input_path}")


def _write_unexpected_error_file(input_path: str, exception_message: str) -> None:
    base_name: str = os.path.basename(input_path)
    name_without_ext: str = os.path.splitext(base_name)[0]
    directory: str = os.path.dirname(os.path.abspath(input_path))
    error_file_name: str = f"{name_without_ext}_improved_error.tsv"
    error_file_path: str = os.path.join(directory, error_file_name)
    with open(error_file_path, "w", encoding="utf-8", newline="") as error_file:
        error_file.write(f"Error: unexpected exception. Detail = {exception_message}")


def _load_input_tsv(input_path: str) -> pd.DataFrame:
    return pd.read_csv(
        input_path,
        sep="\t",
        dtype=str,
        encoding="utf-8",
        engine="python",
        keep_default_na=False,
    )


def _is_time_like_string(pszValue: str) -> bool:
    pszStrippedValue: str = pszValue.strip()
    if pszStrippedValue == "":
        return False
    return objTimePattern.match(pszStrippedValue) is not None


def _calculate_time_like_ratio(objColumnValues: pd.Series) -> float:
    iTimeLikeCount: int = 0
    iNonEmptyCount: int = 0
    for objValue in objColumnValues:
        pszCellValue: str = "" if pd.isna(objValue) else str(objValue)
        pszStrippedValue: str = pszCellValue.strip()
        if pszStrippedValue == "":
            continue
        iNonEmptyCount += 1
        if _is_time_like_string(pszStrippedValue):
            iTimeLikeCount += 1
    if iNonEmptyCount == 0:
        return 0.0
    return iTimeLikeCount / iNonEmptyCount


def _is_time_column(objColumnValues: pd.Series, pszColumnName: str) -> bool:
    fTimeLikeRatio: float = _calculate_time_like_ratio(objColumnValues)
    if fTimeLikeRatio >= fTimeLikeRatioThreshold:
        return True
    return False


def _identify_time_columns(df: pd.DataFrame) -> List[bool]:
    objTimeColumnFlags: List[bool] = []
    for objColumnName in df.columns:
        objColumnValues: pd.Series = df[objColumnName]
        bIsTimeColumn: bool = _is_time_column(objColumnValues, str(objColumnName))
        objTimeColumnFlags.append(bIsTimeColumn)
    return objTimeColumnFlags


def _fill_blank_time_cells(
    df: pd.DataFrame,
    objTimeColumnFlags: List[bool],
    report_lines: List[str],
) -> pd.DataFrame:
    pszZeroTime: str = "0:00:00"
    for row_index in range(len(df)):
        for col_index in range(len(df.columns)):
            if not objTimeColumnFlags[col_index]:
                continue
            raw_value: object = df.iat[row_index, col_index]
            pszCellValue: str = "" if pd.isna(raw_value) else str(raw_value)
            if pszCellValue.strip() == "":
                df.iat[row_index, col_index] = pszZeroTime
                report_line: str = (
                    f"Sheet\tR{row_index + 2} C{col_index + 1}\t"
                    f"BEFORE={pszCellValue}\tAFTER={pszZeroTime}"
                )
                report_lines.append(report_line)
    return df


def _simplify_iferror(formula: str) -> Tuple[str, bool]:
    pattern: re.Pattern[str] = re.compile(
        r"^=IFERROR\(\s*IFERROR\(\s*(.+?)\s*,\s*(\"\"|''?)\s*\)\s*,\s*(.+?)\s*\)\s*$",
        re.IGNORECASE,
    )
    match: re.Match[str] | None = pattern.match(formula)
    if match is None:
        return formula, False
    inner_expression: str = match.group(1)
    inner_error_value: str = match.group(2)
    simplified_formula: str = f"=IFERROR({inner_expression},{inner_error_value})"
    return simplified_formula, simplified_formula != formula


def _improve_formula_cell(cell_value: str) -> Tuple[str, bool]:
    if not cell_value.startswith("="):
        return cell_value, False
    simplified_value: str = cell_value
    changed: bool = False
    simplified_value, iferror_changed = _simplify_iferror(simplified_value)
    if iferror_changed:
        changed = True
    return simplified_value, changed


def _process_dataframe(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    improved_rows: List[List[str]] = []
    report_lines: List[str] = []
    columns: List[str] = list(df.columns)
    for row_index in range(len(df)):
        row_data: List[str] = []
        for col_index in range(len(columns)):
            raw_value: object = df.iat[row_index, col_index]
            cell_value: str = "" if pd.isna(raw_value) else str(raw_value)
            improved_value, changed = _improve_formula_cell(cell_value)
            row_data.append(improved_value)
            if changed:
                report_line: str = (
                    f"Sheet\tR{row_index + 2} C{col_index + 1}\t"
                    f"BEFORE={cell_value}\tAFTER={improved_value}"
                )
                report_lines.append(report_line)
        improved_rows.append(row_data)
    improved_df: pd.DataFrame = pd.DataFrame(improved_rows, columns=columns)
    return improved_df, report_lines


def _restore_blank_column_names(df: pd.DataFrame) -> pd.DataFrame:
    objNewColumns: List[str] = []
    for objColumn in df.columns:
        pszColumn: str = str(objColumn)
        if re.match(r"^Unnamed:\s*\d+$", pszColumn):
            objNewColumns.append("")
        else:
            objNewColumns.append(pszColumn)
    df.columns = objNewColumns
    return df


def _write_output_files(
    input_path: str,
    improved_df: pd.DataFrame,
    report_lines: List[str],
) -> None:
    base_name: str = os.path.basename(input_path)
    name_without_ext: str = os.path.splitext(base_name)[0]
    directory: str = os.path.dirname(os.path.abspath(input_path))
    improved_file_name: str = f"{name_without_ext}_improved.tsv"
    improved_file_path: str = os.path.join(directory, improved_file_name)
    report_file_name: str = f"{name_without_ext}_improved_report.txt"
    report_file_path: str = os.path.join(directory, report_file_name)
    improved_df.to_csv(
        improved_file_path,
        sep="\t",
        index=False,
        encoding="utf-8",
        lineterminator="\n",
    )
    with open(report_file_path, "w", encoding="utf-8", newline="") as report_file:
        report_file.write("\n".join(report_lines))


def main() -> None:
    if len(sys.argv) < 2:
        _output_argument_error()
        return
    input_path: str = sys.argv[1]
    if not os.path.isfile(input_path):
        _output_missing_file_error(input_path)
        return
    try:
        df: pd.DataFrame = _load_input_tsv(input_path)
        improved_df, report_lines = _process_dataframe(df)
        objTimeColumnFlags: List[bool] = _identify_time_columns(improved_df)
        improved_df = _fill_blank_time_cells(improved_df, objTimeColumnFlags, report_lines)
        improved_df = _restore_blank_column_names(improved_df)
        _write_output_files(input_path, improved_df, report_lines)
    except Exception as exc:  # noqa: BLE001
        _write_unexpected_error_file(input_path, str(exc))
        print(f"Error: unexpected exception. Detail = {exc}")


if __name__ == "__main__":
    main()
