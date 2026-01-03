import os
import sys
from typing import Any, List

import pandas as pd


def normalize_value(objValue: Any) -> Any:
    pszText: str
    if objValue is None:
        return 0.0
    if pd.isna(objValue):
        return 0.0
    if isinstance(objValue, str):
        pszText = objValue.strip()
        if pszText == "":
            return 0.0
        try:
            return float(pszText)
        except ValueError:
            return pszText
    try:
        return float(objValue)
    except (TypeError, ValueError):
        return str(objValue)


def to_output_value(objValue: Any) -> str:
    if objValue is None:
        return ""
    if pd.isna(objValue):
        return ""
    return str(objValue)


def write_error_file(pszFileName: str, arrMessages: List[str]) -> None:
    pszJoinedMessage: str = "\n".join(arrMessages)
    with open(pszFileName, "w", encoding="utf-8") as objFile:
        objFile.write(pszJoinedMessage)


def compare_rows(arrLeftRows: List[List[Any]], arrRightRows: List[List[Any]]) -> List[str]:
    arrDiffLines: List[str] = []
    for iRowIndex, arrLeftRow in enumerate(arrLeftRows):
        arrRightRow: List[Any] = arrRightRows[iRowIndex]
        for iColIndex, objLeftValue in enumerate(arrLeftRow):
            objRightValue: Any = arrRightRow[iColIndex]
            objNormLeft: Any = normalize_value(objLeftValue)
            objNormRight: Any = normalize_value(objRightValue)
            bIsStringComparison: bool = isinstance(objNormLeft, str) or isinstance(objNormRight, str)
            bIsMatch: bool
            if bIsStringComparison:
                bIsMatch = str(objNormLeft) == str(objNormRight)
            else:
                bIsMatch = objNormLeft == objNormRight
            if not bIsMatch:
                iReportRow: int = iRowIndex + 1
                iReportCol: int = iColIndex + 1
                pszLeftDisplay: str = to_output_value(objLeftValue)
                pszRightDisplay: str = to_output_value(objRightValue)
                pszNormLeftDisplay: str = str(objNormLeft)
                pszNormRightDisplay: str = str(objNormRight)
                pszLine: str = (
                    f"R{iReportRow} C{iReportCol}\tLEFT={pszLeftDisplay}\tRIGHT={pszRightDisplay}"
                    f"\tNORM_LEFT={pszNormLeftDisplay}\tNORM_RIGHT={pszNormRightDisplay}"
                )
                arrDiffLines.append(pszLine)
    return arrDiffLines


def main() -> int:
    if len(sys.argv) < 3:
        pszError1: str = "Error: input TSV file paths are not specified (insufficient arguments)."
        pszError2: str = "Usage: python compare_tsv_with_blank_zero.py <left_tsv_path> <right_tsv_path>"
        pszError3: str = "Example: python compare_tsv_with_blank_zero.py C:\\Data\\A.tsv C:\\Data\\B.tsv"
        arrErrors: List[str] = [pszError1, pszError2, pszError3]
        for pszLine in arrErrors:
            print(pszLine)
        write_error_file("compare_tsv_with_blank_zero_error_argument.tsv", arrErrors)
        return 2

    pszLeftPath: str = sys.argv[1]
    pszRightPath: str = sys.argv[2]

    arrMissingMessages: List[str] = []
    if not os.path.exists(pszLeftPath):
        arrMissingMessages.append(f"Error: input TSV file not found. Path = {pszLeftPath}")
    if not os.path.exists(pszRightPath):
        arrMissingMessages.append(f"Error: input TSV file not found. Path = {pszRightPath}")

    if len(arrMissingMessages) > 0:
        for pszMessage in arrMissingMessages:
            print(pszMessage)
        write_error_file("compare_tsv_with_blank_zero_error.tsv", arrMissingMessages)
        return 2

    try:
        objLeftDf: pd.DataFrame = pd.read_csv(
            pszLeftPath, sep="\t", dtype=str, encoding="utf-8", engine="python"
        )
        objRightDf: pd.DataFrame = pd.read_csv(
            pszRightPath, sep="\t", dtype=str, encoding="utf-8", engine="python"
        )
    except Exception as objException:  # noqa: BLE001
        pszMessage: str = f"Error: unexpected exception. Detail = {objException}"
        print(pszMessage)
        write_error_file("compare_tsv_with_blank_zero_error.tsv", [pszMessage])
        return 2

    if objLeftDf.shape != objRightDf.shape:
        pszShapeMessage: str = (
            f"Error: TSV shape mismatch. Left = {objLeftDf.shape[0]}x{objLeftDf.shape[1]}, "
            f"Right = {objRightDf.shape[0]}x{objRightDf.shape[1]}"
        )
        print(pszShapeMessage)
        write_error_file("compare_tsv_with_blank_zero_error.tsv", [pszShapeMessage])
        return 2

    arrLeftRows: List[List[Any]] = [list(objLeftDf.columns)] + objLeftDf.astype(object).values.tolist()
    arrRightRows: List[List[Any]] = [list(objRightDf.columns)] + objRightDf.astype(object).values.tolist()

    arrDifferences: List[str] = compare_rows(arrLeftRows, arrRightRows)

    pszReportFile: str = "compare_tsv_with_blank_zero_report.txt"
    if len(arrDifferences) == 0:
        with open(pszReportFile, "w", encoding="utf-8") as objReport:
            objReport.write("OK: no differences\n")
        return 0

    iDiffCount: int = len(arrDifferences)
    with open(pszReportFile, "w", encoding="utf-8") as objReport:
        objReport.write(f"NG: differences found. Count = {iDiffCount}\n")
        for pszLine in arrDifferences:
            objReport.write(f"{pszLine}\n")
    return 1


if __name__ == "__main__":
    sys.exit(main())
