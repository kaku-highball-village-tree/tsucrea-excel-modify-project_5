# -*- coding: utf-8 -*-
"""
Make_PjSummary_PL_Manhour.py

A∪B統合版
- A: make_manhour_to_sheet8_01_0001.py
- B: PL_CsvToTsv_Cmd.py
"""

from __future__ import annotations

import sys


PL_SOURCE = r'''import csv
import os
import re
import sys
from typing import List, Tuple


def get_target_year_month_from_filename(pszInputFilePath: str) -> Tuple[int, int]:
    pszBaseName: str = os.path.basename(pszInputFilePath)
    objMatch: re.Match[str] | None = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        raise ValueError("入力ファイル名から対象年月を取得できません。")
    iYearTwoDigits: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iYear: int = 2000 + iYearTwoDigits
    return iYear, iMonth


def get_target_year_month_from_period_row(pszRowA: str) -> Tuple[int, int]:
    pszNormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
    pszNormalized = pszNormalized.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    objMatch: re.Match[str] | None = re.search(r"(?:自)?(\d{4})年(\d{1,2})月(?:度)?", pszNormalized)
    if objMatch is None:
        objMatch = re.search(r"(\d{4})[./-](\d{1,2})", pszNormalized)
    if objMatch is None:
        raise ValueError("集計期間から対象年月を取得できません。")
    iYear: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    return iYear, iMonth


def read_csv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    try:
        with open(pszInputFilePath, mode="r", encoding="utf-8-sig", errors="strict", newline="") as objFile:
            objReader: csv.reader = csv.reader(objFile)
            for objRow in objReader:
                objRows.append(objRow)
        append_debug_log("input decoded as utf-8-sig")
        return objRows
    except UnicodeDecodeError:
        append_debug_log("utf-8-sig decode failed; retrying with cp932")

    with open(pszInputFilePath, mode="r", encoding="cp932", errors="strict", newline="") as objFile:
        objReader = csv.reader(objFile)
        for objRow in objReader:
            objRows.append(objRow)
    append_debug_log("input decoded as cp932")
    return objRows


def write_tsv_rows(pszOutputFilePath: str, objRows: List[List[str]]) -> None:
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objFile:
        objWriter: csv.writer = csv.writer(objFile, delimiter="\t", lineterminator="\n")
        for objRow in objRows:
            objWriter.writerow(objRow)


def read_tsv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objFile:
        objReader: csv.reader = csv.reader(objFile, delimiter="\t")
        for objRow in objReader:
            objRows.append(objRow)
    return objRows


def build_first_column_rows(objRows: List[List[str]]) -> List[List[str]]:
    return [[objRow[0] if objRow else ""] for objRow in objRows]



def build_unique_subjects(objSubjectRows: List[List[str]]) -> List[str]:
    objSubjects: List[str] = []
    objSeen: set[str] = set()
    for objRow in objSubjectRows:
        pszValue: str = objRow[0] if objRow else ""
        if pszValue == "" or pszValue in objSeen:
            continue
        objSeen.add(pszValue)
        objSubjects.append(pszValue)
    return objSubjects



def build_union_subject_order(objSubjectLists: List[List[str]]) -> List[str]:
    objAppearanceOrder: dict[str, int] = {}
    iCounter: int = 0
    for objSubjectList in objSubjectLists:
        for pszSubject in objSubjectList:
            if pszSubject not in objAppearanceOrder:
                objAppearanceOrder[pszSubject] = iCounter
                iCounter += 1

    objAdjacency: dict[str, set[str]] = {psz: set() for psz in objAppearanceOrder}
    objIndegree: dict[str, int] = {psz: 0 for psz in objAppearanceOrder}
    for objSubjectList in objSubjectLists:
        for iIndex in range(len(objSubjectList) - 1):
            pszBefore: str = objSubjectList[iIndex]
            pszAfter: str = objSubjectList[iIndex + 1]
            if pszAfter not in objAdjacency[pszBefore]:
                objAdjacency[pszBefore].add(pszAfter)
                objIndegree[pszAfter] += 1

    objOrderedSubjects: List[str] = []
    objReady: List[str] = [
        pszSubject for pszSubject, iDegree in objIndegree.items() if iDegree == 0
    ]
    objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    while objReady:
        pszSubject = objReady.pop(0)
        objOrderedSubjects.append(pszSubject)
        for pszNext in sorted(objAdjacency[pszSubject], key=lambda psz: objAppearanceOrder[psz]):
            objIndegree[pszNext] -= 1
            if objIndegree[pszNext] == 0:
                objReady.append(pszNext)
        objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    if len(objOrderedSubjects) != len(objAppearanceOrder):
        return list(objAppearanceOrder.keys())

    return objOrderedSubjects



def build_subject_vertical_rows(objSubjects: List[str]) -> List[List[str]]:
    return [[pszSubject] for pszSubject in objSubjects]



def normalize_project_name(pszProjectName: str) -> str:
    if pszProjectName == "":
        return pszProjectName
    normalized = pszProjectName.replace("\t", "_")
    normalized = re.sub(r"^([A-OQ-Z]\d{3})([ 　]+)", r"\1_", normalized)
    normalized = re.sub(r"^([A-OQ-Z]\d{3})(【)", r"\1_\2", normalized)
    normalized = re.sub(r"^(P\d{5})([ 　]+)", r"\1_", normalized)
    normalized = re.sub(r"^(P\d{5})(【)", r"\1_\2", normalized)
    return normalized


def normalize_project_names_in_row(objRows: List[List[str]], iRowIndex: int) -> None:
    if iRowIndex < 0 or iRowIndex >= len(objRows):
        return
    objTargetRow = objRows[iRowIndex]
    for iIndex, pszProjectName in enumerate(objTargetRow):
        objTargetRow[iIndex] = normalize_project_name(pszProjectName)


def find_row_index_with_subject_tab(objRows: List[List[str]], iStartIndex: int) -> int | None:
    for iRowIndex in range(iStartIndex, len(objRows)):
        objRow = objRows[iRowIndex]
        if any(
            "科目名\t" in pszValue or pszValue.strip() == "科目名"
            for pszValue in objRow
        ):
            return iRowIndex
    return None


def build_pj_name_vertical_rows(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []

    objHeaderRow: List[str] = objRows[0]
    objItemRows: List[List[str]] = objRows[1:]

    objVerticalRows: List[List[str]] = []
    objVerticalHeader: List[str] = ["PJ名称"]
    for objItemRow in objItemRows:
        pszItemName: str = objItemRow[0] if len(objItemRow) > 0 else ""
        objVerticalHeader.append(pszItemName)
    objVerticalRows.append(objVerticalHeader)

    for iColumnIndex in range(1, len(objHeaderRow)):
        pszProjectName: str = objHeaderRow[iColumnIndex]
        objVerticalRow: List[str] = [pszProjectName]
        for objItemRow in objItemRows:
            pszValue: str = objItemRow[iColumnIndex] if len(objItemRow) > iColumnIndex else ""
            objVerticalRow.append(pszValue)
        objVerticalRows.append(objVerticalRow)

    return objVerticalRows


def write_first_row_tabs_to_newlines(pszInputFilePath: str, pszOutputFilePath: str) -> None:
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objInputFile:
        pszFirstLine: str = objInputFile.readline()
    pszConverted: str = pszFirstLine.replace("\t", "\n")
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objOutputFile.write(pszConverted)


def insert_company_expense_columns(objRows: List[List[str]]) -> None:
    if not objRows:
        return
    objHeaderRow: List[str] = objRows[0]
    try:
        iHeadOfficeIndex: int = objHeaderRow.index("本部")
    except ValueError:
        return

    objExpenseColumns: List[str] = [
        "1Cカンパニー販管費",
        "2Cカンパニー販管費",
        "3Cカンパニー販管費",
        "4Cカンパニー販管費",
        "事業開発カンパニー販管費",
        "社長室カンパニー販管費",
        "本部カンパニー販管費",
    ]
    iInsertIndex: int = iHeadOfficeIndex + 1
    objHeaderRow[iInsertIndex:iInsertIndex] = objExpenseColumns
    for objRow in objRows[1:]:
        objRow[iInsertIndex:iInsertIndex] = ["0"] * len(objExpenseColumns)


COMPANY_EXPENSE_REPLACEMENTS: dict[str, str] = {
    "1Cカンパニー販管費": "C001_1Cカンパニー販管費",
    "2Cカンパニー販管費": "C002_2Cカンパニー販管費",
    "3Cカンパニー販管費": "C003_3Cカンパニー販管費",
    "4Cカンパニー販管費": "C004_4Cカンパニー販管費",
    "事業開発カンパニー販管費": "C005_事業開発カンパニー販管費",
    "社長室カンパニー販管費": "C006_社長室カンパニー販管費",
    "本部カンパニー販管費": "C007_本部カンパニー販管費",
}


def replace_company_expense_labels(objRows: List[List[str]], objReplacementMap: dict[str, str]) -> None:
    for objRow in objRows:
        for iIndex, pszValue in enumerate(objRow):
            if pszValue in objReplacementMap:
                objRow[iIndex] = objReplacementMap[pszValue]


def append_debug_log(pszMessage: str, pszDebugFilePath: str = "debug.txt") -> None:
    with open(pszDebugFilePath, mode="a", encoding="utf-8", newline="") as objDebugFile:
        objDebugFile.write(f"{pszMessage}\n")


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python src/PL_CsvToTsv_Cmd.py <csv_file> [<csv_file> ...]")
        return 1

    iExitCode: int = 0
    objCostReportVerticalFilePaths: List[str] = []
    for pszInputFilePath in sys.argv[1:]:
        try:
            append_debug_log("start")
            iFileYear: int
            iFileMonth: int
            iFileYear, iFileMonth = get_target_year_month_from_filename(pszInputFilePath)
            append_debug_log(f"filename parsed: {iFileYear}-{iFileMonth:02d}")

            if not os.path.isfile(pszInputFilePath):
                raise FileNotFoundError(f"入力ファイルが存在しません: {pszInputFilePath}")

            objRows: List[List[str]] = read_csv_rows(pszInputFilePath)
            if len(objRows) < 2:
                raise ValueError("集計期間の取得に必要な行が存在しません。")
            append_debug_log(f"rows read: {len(objRows)}")

            normalize_project_names_in_row(objRows, 7)
            iSubjectRowIndex = find_row_index_with_subject_tab(objRows, 8)
            if iSubjectRowIndex is not None:
                normalize_project_names_in_row(objRows, iSubjectRowIndex)
            append_debug_log("project names normalized")

            pszRowA: str = objRows[1][1] if len(objRows[1]) > 1 else ""
            append_debug_log(f"B2 value: {pszRowA}")
            pszRowANormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
            if "期首振戻" in pszRowANormalized:
                append_debug_log("period parse skipped due to 期首振戻; using filename")
            else:
                iPeriodYear: int
                iPeriodMonth: int
                iPeriodYear, iPeriodMonth = get_target_year_month_from_period_row(pszRowA)
                append_debug_log(f"period parsed: {iPeriodYear}-{iPeriodMonth:02d}")

                if iFileYear != iPeriodYear or iFileMonth != iPeriodMonth:
                    raise ValueError("ファイル名と集計期間の対象年月が一致しません。")
                append_debug_log("period matches filename")

            pszMonth: str = f"{iFileMonth:02d}"
            pszOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月.tsv"
            pszCostReportFilePath: str = f"製造原価報告書_{iFileYear}年{pszMonth}月.tsv"
            objOutputRows: List[List[str]] = []
            objCostReportRows: List[List[str]] = []
            iSplitIndex: int | None = None
            for iRowIndex in range(7, len(objRows) - 1):
                objRow: List[str] = objRows[iRowIndex]
                objNextRow: List[str] = objRows[iRowIndex + 1]
                if objRow and objNextRow and objRow[0] == "当期純利益" and objNextRow[0] == "科目名":
                    iSplitIndex = iRowIndex
                    break

            if iSplitIndex is None:
                for iRowIndex in range(7, len(objRows)):
                    objRow = objRows[iRowIndex]
                    objOutputRows.append(objRow[:])
            else:
                for iRowIndex in range(7, iSplitIndex + 1):
                    objRow = objRows[iRowIndex]
                    objOutputRows.append(objRow[:])
                for iRowIndex in range(iSplitIndex + 1, len(objRows)):
                    objRow = objRows[iRowIndex]
                    objCostReportRows.append(objRow[:])
            append_debug_log(f"output rows prepared: {len(objOutputRows)}")

            if (iFileYear, iFileMonth) <= (2025, 7):
                insert_company_expense_columns(objOutputRows)
                append_debug_log("company expense columns inserted")

            replace_company_expense_labels(
                objOutputRows,
                COMPANY_EXPENSE_REPLACEMENTS,
            )
            append_debug_log("company expense labels replaced")

            write_tsv_rows(pszOutputFilePath, objOutputRows)
            append_debug_log(f"tsv written: {pszOutputFilePath}")

            if objCostReportRows:
                write_tsv_rows(pszCostReportFilePath, objCostReportRows)
                append_debug_log(f"tsv written: {pszCostReportFilePath}")
                objCostReportTsvRows: List[List[str]] = read_tsv_rows(pszCostReportFilePath)
                objCostReportVerticalRows: List[List[str]] = build_first_column_rows(objCostReportTsvRows)
                pszCostReportVerticalFilePath: str = (
                    f"製造原価報告書_{iFileYear}年{pszMonth}月_科目名_vertical.tsv"
                )
                write_tsv_rows(pszCostReportVerticalFilePath, objCostReportVerticalRows)
                append_debug_log(f"vertical tsv written: {pszCostReportVerticalFilePath}")

                objCostReportVerticalFilePaths.append(pszCostReportVerticalFilePath)


            pszVerticalOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月_PJ名称_vertical.tsv"
            objPjNameVerticalRows: List[List[str]] = build_pj_name_vertical_rows(objOutputRows)
            write_tsv_rows(pszVerticalOutputFilePath, objPjNameVerticalRows)
            append_debug_log(f"vertical tsv written: {pszVerticalOutputFilePath}")
        except Exception as objException:
            iExitCode = 1
            append_debug_log(f"error: {objException}")
            print(objException)
            try:
                iErrorYear: int
                iErrorMonth: int
                iErrorYear, iErrorMonth = get_target_year_month_from_filename(pszInputFilePath)
                pszErrorMonth: str = f"{iErrorMonth:02d}"
                pszErrorFilePath: str = f"損益計算書_{iErrorYear}年{pszErrorMonth}月_error.txt"
            except Exception:
                pszBaseName: str = os.path.basename(pszInputFilePath)
                pszErrorFilePath = f"{pszBaseName}_error.txt"
            with open(pszErrorFilePath, mode="w", encoding="utf-8", newline="") as objErrorFile:
                objErrorFile.write(str(objException))

    create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths)
    return iExitCode


def create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths: List[str]) -> None:
    if not objCostReportVerticalFilePaths:
        return

    objSubjectLists: List[List[str]] = []
    for pszFilePath in objCostReportVerticalFilePaths:
        objRows: List[List[str]] = read_tsv_rows(pszFilePath)
        objSubjectLists.append(build_unique_subjects(objRows))
    objUnionSubjects: List[str] = build_union_subject_order(objSubjectLists)
    objUnionRows: List[List[str]] = build_subject_vertical_rows(objUnionSubjects)

    for pszFilePath in objCostReportVerticalFilePaths:
        pszUnionFilePath: str = pszFilePath.replace("_科目名_vertical.tsv", "_科目名_A∪B_vertical.tsv")
        write_tsv_rows(pszUnionFilePath, objUnionRows)
        append_debug_log(f"union vertical tsv written: {pszUnionFilePath}")


if __name__ == "__main__":
    sys.exit(main())
'''

MANHOUR_SOURCE = r'''# //////////////////////////////
# Main editable file
#
# All implementation must be done in this file.
# Other Python files under src/ are reference-only
# and must not be modified.
# //////////////////////////////

# -*- coding: utf-8 -*-
# ///////////////////////////////////////////////////////////////
#
# make_manhour_to_sheet8_01_0001.py
#
# 役割:
#   単一のジョブカン工数 CSV を入力として、
#   正解スクリプト群を (1)〜(7) の順で直列実行し、
#   Sheet7.tsv / Sheet8.tsv / Sheet9.tsv までを
#   同一フォルダに生成する。
#
# 実行例:
#   python make_manhour_to_sheet8_01_0001.py manhour_xxxxxx.csv
#
# ///////////////////////////////////////////////////////////////

from __future__ import annotations

import argparse
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple
import pandas as pd


def get_target_year_month_from_filename(pszInputFilePath: str) -> Tuple[int, int]:
    pszBaseName: str = os.path.basename(pszInputFilePath)
    objMatch: re.Match[str] | None = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        raise ValueError("入力ファイル名から対象年月を取得できません。")
    iYearTwoDigits: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iYear: int = 2000 + iYearTwoDigits
    return iYear, iMonth

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: csv_to_tsv_h_mm_ss.py
# ///////////////////////////////////////////////////////////////
pszSource_csv_to_tsv_h_mm_ss_py: str = r'''# -*- coding: utf-8 -*-
"""
csv_to_tsv_h_mm_ss.py

ジョブカン工数 CSV を読み込み、TSV に変換する。
工数列(F列)が「h:mm」形式の場合、「h:mm:00」を付与して「h:mm:ss」にそろえる。
また、K列も同様に「h:mm」形式の場合、「h:mm:00」を付与する。

入力:  manhour_*.csv
出力:  manhour_*.tsv

注意:
  - 入力 CSV は UTF-8 を想定。
  - 出力 TSV は UTF-8。
"""

import sys
import os
import csv
from typing import List, Dict


# //////////////////////////////
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
# //////////////////////////////
def build_output_file_full_path(
    pszInputFileFullPath: str, 
    pszOutputSuffix: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)

    pszStem: str = os.path.splitext(pszBaseName)[0]
    pszOutputFileName: str = pszStem + pszOutputSuffix

    pszOutputFileFullPath: str = os.path.join(pszDirectory, pszOutputFileName)
    return pszOutputFileFullPath


# //////////////////////////////
#
# 工数文字列 "h:mm" を "h:mm:ss" に揃える関数。
# 例: "7:30" -> "7:30:00"
#
# //////////////////////////////
def normalize_time_h_mm_to_h_mm_ss(
    pszTimeText: str,
) -> str:
    pszText: str = (pszTimeText or "").strip()
    if pszText == "":
        return ""

    # "h:mm:ss" はそのまま
    if pszText.count(":") == 2:
        return pszText

    # "h:mm" の場合のみ ":00" を付与
    if pszText.count(":") == 1:
        return pszText + ":00"

    return pszText


# //////////////////////////////
#
# 入力CSVを読み込み、TSVに変換する関数。
#
# //////////////////////////////
def convert_csv_to_tsv_file(
    pszInputCsvPath: str,
) -> str:
    if not os.path.exists(pszInputCsvPath):
        raise FileNotFoundError(f"Input CSV not found: {pszInputCsvPath}")

    pszOutputTsvPath: str = build_output_file_full_path(pszInputCsvPath, ".tsv")

    objRows: List[List[str]] = []

    # まず UTF-8(BOM あり) で読み込み、失敗したら cp932 で再トライする。
    arrEncodings: List[str] = [
        "utf-8-sig",
        "cp932",
    ]

    objLastDecodeError: Exception | None = None

    for pszEncoding in arrEncodings:
        try:
            with open(
                pszInputCsvPath,
                mode="r",
                encoding=pszEncoding,
                newline="",
            ) as objInputFile:
                objReader: csv.reader = csv.reader(objInputFile)
                for objRow in objReader:
                    objRows.append(list(objRow))
            objLastDecodeError = None
            break
        except UnicodeDecodeError as objError:
            objLastDecodeError = objError
            objRows = []

    if objLastDecodeError is not None:
        raise objLastDecodeError

    # 行数が少ない場合はそのまま出す
    if len(objRows) <= 1:
        with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            for objRow in objRows:
                objWriter.writerow(objRow)
        return pszOutputTsvPath

    # ヘッダ行(0行目)を含めて処理
    #   ※ 本バージョンでは、F列・K列の値が「h:mm」形式の場合、
    #      自動的に「h:mm:00」を付与する (例: 7:30 → 7:30:00)。

    # F列 = index 5
    iTimeColumnIndexF: int = 5
    # K列 = index 10
    iTimeColumnIndexK: int = 10

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]

        # 列数チェック
        if iTimeColumnIndexF < len(objRow):
            objRow[iTimeColumnIndexF] = normalize_time_h_mm_to_h_mm_ss(objRow[iTimeColumnIndexF])

        if iTimeColumnIndexK < len(objRow):
            objRow[iTimeColumnIndexK] = normalize_time_h_mm_to_h_mm_ss(objRow[iTimeColumnIndexK])

        objRows[iRowIndex] = objRow

    # ------------------------------------------------------------
    # ヘッダ先頭セルだけを正規化する。
    # 目的:
    #   先頭セルに UTF-8 BOM(\ufeff) やダブルクォートが残っていると、
    #   writer が " を "" にエスケープして見た目が崩れるため、
    #   "日時" を 日時 にそろえる。
    # 対象:
    #   0行目(ヘッダ行)の0列目のみ。
    # ------------------------------------------------------------
    if len(objRows) >= 1 and len(objRows[0]) >= 1:
        pszHeaderFirstCell: str = objRows[0][0]

        # BOM を除去
        if pszHeaderFirstCell.startswith("\ufeff"):
            pszHeaderFirstCell = pszHeaderFirstCell.lstrip("\ufeff")

        # 外側の " を 1回だけ除去し、"" を " に戻す
        if len(pszHeaderFirstCell) >= 2 and pszHeaderFirstCell.startswith('"') and pszHeaderFirstCell.endswith('"'):
            pszHeaderFirstCell = pszHeaderFirstCell[1:-1]
            pszHeaderFirstCell = pszHeaderFirstCell.replace('""', '"')

        # まだ "日時" の形なら再度外側だけ落とす
        if len(pszHeaderFirstCell) >= 2 and pszHeaderFirstCell.startswith('"') and pszHeaderFirstCell.endswith('"'):
            pszHeaderFirstCell = pszHeaderFirstCell[1:-1]

        objRows[0][0] = pszHeaderFirstCell

    with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objRows:
            objWriter.writerow(objRow)

    return pszOutputTsvPath


# //////////////////////////////
#
# main
#
# //////////////////////////////
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python csv_to_tsv_h_mm_ss.py <input_manhour_csv>")
        return 1

    pszInputCsvPath: str = sys.argv[1]

    pszOutputTsvPath: str = convert_csv_to_tsv_file(pszInputCsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: manhour_remove_uninput_rows.py
# ///////////////////////////////////////////////////////////////
pszSource_manhour_remove_uninput_rows_py: str = r'''###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
# //////////////////////////////

# -*- coding: utf-8 -*-
# ///////////////////////////////////////////////////////////////
#  manhour_remove_uninput_rows.py
#  （G,H,I,J 列に「未入力」が含まれる行を削除する）
# ///////////////////////////////////////////////////////////////

import os
import sys
from typing import List

import pandas as pd
from pandas import DataFrame


# ---------------------------------------------------------------
# エラー内容を TSV ファイルとして書き出す
# ---------------------------------------------------------------
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)
        if not pszErrorMessage.endswith("\n"):
            objFile.write("\n")


# ---------------------------------------------------------------
# 出力ファイルパスを構築する
#   例:
#     入力: C:\Data\foo.tsv
#     出力: C:\Data\foo_removed_uninput.tsv
# ---------------------------------------------------------------
def build_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszOutputBaseName: str = pszRootName + "_removed_uninput.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# メイン処理本体
#   ・入力 TSV を読み込み
#   ・G,H,I,J 列に「未入力」がある行を削除
#   ・*_removed_uninput.tsv として出力
# ---------------------------------------------------------------
def make_removed_uninput_tsv_from_manhour_tsv(
    pszInputFileFullPath: str,
) -> None:
    # 入力ファイル存在チェック
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectory: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)
        pszRootName: str
        pszExt: str
        pszRootName, pszExt = os.path.splitext(pszBaseName)
        pszErrorFileFullPath: str = os.path.join(
            pszDirectory,
            pszRootName + "_error.tsv",
        )

        write_error_tsv(
            pszErrorFileFullPath,
            "Error: input TSV file not found. Path = {0}".format(
                pszInputFileFullPath
            ),
        )
        return

    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for removing '未入力'. "
            "Detail = {0}".format(objException),
        )
        return

    # 列数チェック（G,H,I,J 列があるか: 少なくとも 10 列必要）
    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 10:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required columns G-J do not exist (need at least 10 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    # 対象列名（G,H,I,J 列）を取得
    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszColumnG: str = objColumnNameList[6]   # G 列
    pszColumnH: str = objColumnNameList[7]   # H 列
    pszColumnI: str = objColumnNameList[8]   # I 列
    pszColumnJ: str = objColumnNameList[9]   # J 列

    # 「未入力」を含む行の判定
    try:
        objSeriesHasUninputG = (
            objDataFrame[pszColumnG].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputH = (
            objDataFrame[pszColumnH].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputI = (
            objDataFrame[pszColumnI].fillna("").astype(str).str.strip() == "未入力"
        )
        objSeriesHasUninputJ = (
            objDataFrame[pszColumnJ].fillna("").astype(str).str.strip() == "未入力"
        )

        objSeriesHasUninputAny = (
            objSeriesHasUninputG
            | objSeriesHasUninputH
            | objSeriesHasUninputI
            | objSeriesHasUninputJ
        )

        # 「未入力」を含まない行だけを残す
        objDataFrameFiltered: DataFrame = objDataFrame.loc[~objSeriesHasUninputAny].copy()

    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while filtering rows with '未入力'. "
            "Detail = {0}".format(objException),
        )
        return

    # TSV として書き出し
    try:
        objDataFrameFiltered.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing TSV without '未入力' rows. "
            "Detail = {0}".format(objException),
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足時は専用のエラー TSV を出力
    if iArgCount < 2:
        pszProgramName: str = os.path.basename(sys.argv[0])
        pszErrorFileFullPath: str = "manhour_remove_uninput_rows_error_argument.tsv"

        pszErrorMessage: str = (
            "Error: input TSV file path is not specified (insufficient arguments).\n"
            "Usage: python {0} <input_tsv_file_path>\n"
            "Example: python {0} C:\\Data\\manhour_202511181454691c0a3179197_sorted_staff_code.tsv"
        ).format(pszProgramName)

        write_error_tsv(
            pszErrorFileFullPath,
            pszErrorMessage,
        )
        return

    pszInputFileFullPath: str = sys.argv[1]

    make_removed_uninput_tsv_from_manhour_tsv(pszInputFileFullPath)


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: sort_manhour_by_staff_code.py
# ///////////////////////////////////////////////////////////////
pszSource_sort_manhour_by_staff_code_py: str = r'''# -*- coding: utf-8 -*-
"""
sort_manhour_by_staff_code.py

ジョブカン工数 TSV を読み込み、
スタッフコード(列B)の数値順に並び替えて TSV を出力する。

入力:  manhour_*_removed_uninput.tsv
出力:  manhour_*_removed_uninput_sorted_staff_code.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import os
import sys
from typing import List

import pandas as pd
from pandas import DataFrame


# //////////////////////////////
#
# 指定したパスにエラーTSVを書き込む（既存スクリプト踏襲）
#
# //////////////////////////////
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage + "\n")


# //////////////////////////////
#
# 出力ファイルパスを構築（拡張子 .tsv → _sorted_staff_code.tsv）
#
# //////////////////////////////
def build_output_file_full_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRoot, pszExt = os.path.splitext(pszBaseName)

    pszOutputBase: str = pszRoot + "_sorted_staff_code.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBase
    return os.path.join(pszDirectory, pszOutputBase)


# //////////////////////////////
#
# メイン処理（manhour_*.tsv → *_sorted_staff_code.tsv）
#
# //////////////////////////////
def make_sorted_staff_code_tsv_from_manhour_tsv(
    pszInputFileFullPath: str,
) -> str:

    # 入力ファイル存在チェック
    if not os.path.isfile(pszInputFileFullPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputFileFullPath}")

    # 出力ファイルパスを構築
    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading manhour TSV for staff code sort. Detail = {0}".format(objException),
        )
        return pszOutputFileFullPath

    # 列数チェック（2列目が必要）
    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: staff code column (2nd column) does not exist. ColumnCount = {0}".format(iColumnCount),
        )
        return pszOutputFileFullPath

    # ソート対象列（第2列）
    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszSortColumnName: str = objColumnNameList[1]

    # -----------------------------------------------------------
    # スタッフコード列を数値に変換して、数値順にソートする
    # -----------------------------------------------------------
    try:
        # 元データをコピーして、一時列 __sort_staff_code__ を追加
        objSorted: DataFrame = objDataFrame.copy()
        objSorted["__sort_staff_code__"] = pd.to_numeric(
            objSorted[pszSortColumnName],
            errors="coerce",
        )

        # 一時列をキーに安定ソート（数値順）
        objSorted = objSorted.sort_values(
            by="__sort_staff_code__",
            ascending=True,
            kind="mergesort",
        ).drop(columns=["__sort_staff_code__"])

    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while sorting by staff code. Detail = {0}".format(objException),
        )
        return pszOutputFileFullPath

    # TSV 書き込み
    try:
        objSorted.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing sorted staff-code TSV. Detail = {0}".format(objException),
        )
        return pszOutputFileFullPath

    return pszOutputFileFullPath


# //////////////////////////////
#
# main
#
# //////////////////////////////
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python sort_manhour_by_staff_code.py <input_manhour_tsv>")
        return 1

    pszInputTsvPath: str = sys.argv[1]

    pszOutputTsvPath: str = make_sorted_staff_code_tsv_from_manhour_tsv(pszInputTsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: convert_yyyy_mm_dd.py
# ///////////////////////////////////////////////////////////////
pszSource_convert_yyyy_mm_dd_py: str = r'''# -*- coding: utf-8 -*-
"""
convert_yyyy_mm_dd.py

ジョブカン工数 TSV を読み込み、
日付列を yyyy/mm/dd 形式に揃えて Sheet4.tsv を出力する。

入力:  manhour_*_removed_uninput_sorted_staff_code.tsv
出力:  Sheet4.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import os
import sys
import re
from typing import List

import pandas as pd
from pandas import DataFrame


# //////////////////////////////
#
# 指定パスにエラーメッセージだけを書いた TSV を作成する。
#
# //////////////////////////////
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectoryName: str = os.path.dirname(pszOutputFileFullPath)

    if len(pszDirectoryName) > 0 and (not os.path.exists(pszDirectoryName)):
        os.makedirs(pszDirectoryName, exist_ok=True)

    objLines: List[str] = pszErrorMessage.splitlines()

    with open(pszOutputFileFullPath, mode="w", encoding="utf-8", newline="") as objFile:
        for pszLine in objLines:
            objFile.write(pszLine)
            objFile.write("\n")


# //////////////////////////////
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
# //////////////////////////////
def build_output_file_full_path(
    pszInputFileFullPath: str,
    pszOutputFileName: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszOutputFileFullPath: str = os.path.join(pszDirectory, pszOutputFileName)
    return pszOutputFileFullPath


# //////////////////////////////
#
# 単一セルの文字列を yyyy/mm/dd 形式に正規化する。
#
# //////////////////////////////
def normalize_yyyy_mm_dd_in_value(
    objValue: object,
    objPattern: re.Pattern,
) -> object:
    if not isinstance(objValue, str):
        return objValue

    pszText: str = objValue

    objMatch = objPattern.match(pszText)
    if objMatch is None:
        return pszText

    pszYear: str = objMatch.group(1)
    pszMonthRaw: str = objMatch.group(2)
    pszDayRaw: str = objMatch.group(3)

    try:
        iMonth: int = int(pszMonthRaw)
        iDay: int = int(pszDayRaw)
    except Exception:
        return pszText

    if iMonth < 1 or iMonth > 12:
        return pszText
    if iDay < 1 or iDay > 31:
        return pszText

    pszMonth: str = str(iMonth).zfill(2)
    pszDay: str = str(iDay).zfill(2)

    return pszYear + "/" + pszMonth + "/" + pszDay


# //////////////////////////////
#
# DataFrame 全体に対して日付の正規化を行う。
#
# //////////////////////////////
def normalize_yyyy_mm_dd_in_dataframe(
    objDataFrameInput: DataFrame,
) -> DataFrame:
    objPattern: re.Pattern = re.compile(r"^\s*(\d{4})/(\d{1,2})/(\d{1,2})\s*$")

    def _normalize_wrapper(objValue: object) -> object:
        return normalize_yyyy_mm_dd_in_value(objValue, objPattern)

    try:
        objDataFrameOutput: DataFrame = objDataFrameInput.map(_normalize_wrapper)
    except AttributeError:
        objDataFrameOutput = objDataFrameInput.applymap(_normalize_wrapper)

    return objDataFrameOutput


# //////////////////////////////
#
# 入力TSVを読み込み、日付列を正規化して Sheet4.tsv を出力する関数。
#
# //////////////////////////////
def make_sheet4_tsv_from_input_tsv(
    pszInputTsvPath: str,
    pszOutputSheet4TsvPath: str,
) -> str:
    if not os.path.exists(pszInputTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputTsvPath}")

    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputTsvPath,
            sep="\t",
            header=0,
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while reading TSV. Detail = "
            + str(objException)
        )
        print(pszErrorMessage)
        write_error_tsv(pszOutputSheet4TsvPath, pszErrorMessage)
        return pszOutputSheet4TsvPath

    try:
        objDataFrameOutput: DataFrame = normalize_yyyy_mm_dd_in_dataframe(
            objDataFrameInput
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while converting date format. Detail = "
            + str(objException)
        )
        print(pszErrorMessage)
        write_error_tsv(pszOutputSheet4TsvPath, pszErrorMessage)
        return pszOutputSheet4TsvPath

    try:
        objDataFrameOutput.to_csv(
            pszOutputSheet4TsvPath,
            sep="\t",
            index=False,
            encoding="utf-8",
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while writing normalized TSV. Detail = "
            + str(objException)
        )
        print(pszErrorMessage)
        write_error_tsv(pszOutputSheet4TsvPath, pszErrorMessage)
        return pszOutputSheet4TsvPath

    return pszOutputSheet4TsvPath


# //////////////////////////////
#
# main
#
# //////////////////////////////
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python convert_yyyy_mm_dd.py <input_manhour_tsv>")
        return 1

    pszInputTsvPath: str = sys.argv[1]
    pszOutputSheet4TsvPath: str = build_output_file_full_path(pszInputTsvPath, "Sheet4.tsv")

    pszOutputTsvPath: str = make_sheet4_tsv_from_input_tsv(pszInputTsvPath, pszOutputSheet4TsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: make_unique_staff_code_list.py
# ///////////////////////////////////////////////////////////////
pszSource_make_unique_staff_code_list_py: str = r'''# -*- coding: utf-8 -*-
"""
make_unique_staff_code_list.py
（Sheet1_yyyy_mm_dd.tsv から B列のスタッフコード一覧を作成する）
"""

import os
import sys
from typing import List, Set

import pandas as pd
from pandas import DataFrame, Series


# ---------------------------------------------------------------
# エラーメッセージを TSV として書き出す共通関数
# ---------------------------------------------------------------
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)
        if not pszErrorMessage.endswith("\n"):
            objFile.write("\n")


# ---------------------------------------------------------------
# 出力ファイルパスを構築する
#   例:
#     入力: C:\Data\Sheet1_2025_11_18.tsv
#     出力: C:\Data\Sheet1_2025_11_18_unique_staff_code.tsv
# ---------------------------------------------------------------
def build_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseName)

    pszOutputBaseName: str = pszRootName + "_unique_staff_code.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# 実処理本体
#   ・Sheet1_yyyy_mm_dd.tsv を読み込み
#   ・B列（2列目）の重複を除いた一覧を作成
#   ・*_unique_staff_code.tsv として出力
# ---------------------------------------------------------------
def make_unique_staff_code_tsv_from_sheet1_tsv(
    pszInputFileFullPath: str,
) -> None:
    # 入力ファイル存在チェック
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectory: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)
        pszRootName: str
        pszExt: str
        pszRootName, pszExt = os.path.splitext(pszBaseName)
        pszErrorFileFullPath: str = os.path.join(
            pszDirectory,
            pszRootName + "_error.tsv",
        )

        write_error_tsv(
            pszErrorFileFullPath,
            "Error: input TSV file not found. Path = {0}".format(
                pszInputFileFullPath
            ),
        )
        return

    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for unique staff code list. "
            "Detail = {0}".format(objException),
        )
        return

    # 列数チェック（B列が存在するか: 少なくとも 2 列必要）
    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required column B does not exist (need at least 2 columns). "
            "ColumnCount = {0}".format(iColumnCount),
        )
        return

    # B列のヘッダ名とデータ列を取得
    objColumnNameList: List[str] = list(objDataFrame.columns)
    pszStaffCodeColumnName: str = objColumnNameList[1]

    objSeriesStaffCode: Series = objDataFrame.iloc[:, 1]

    # B列の値から、空白行を除き、出現順を保ったまま重複を除去する
    try:
        objListUniqueStaffCode: List[str] = []
        objSetSeen: Set[str] = set()

        for pszValueRaw in objSeriesStaffCode.tolist():
            pszValue: str = "" if pszValueRaw is None else str(pszValueRaw)
            pszValueStripped: str = pszValue.strip()

            # 空文字列は UNIQUE 対象から除外
            if pszValueStripped == "":
                continue

            if pszValueStripped in objSetSeen:
                continue

            objSetSeen.add(pszValueStripped)
            objListUniqueStaffCode.append(pszValueStripped)

        # 出力用 DataFrame を作成（1列のみ）
        objOutputDataFrame: DataFrame = DataFrame(
            {pszStaffCodeColumnName: objListUniqueStaffCode}
        )

    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while creating unique staff code list. "
            "Detail = {0}".format(objException),
        )
        return

    # TSV として書き出し
    try:
        objOutputDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing unique staff code TSV. "
            "Detail = {0}".format(objException),
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足時は専用のエラー TSV を出力
    if iArgCount < 2:
        pszProgramName: str = os.path.basename(sys.argv[0])
        pszErrorFileFullPath: str = "make_unique_staff_code_list_error_argument.tsv"

        pszErrorMessage: str = (
            "Error: input TSV file path is not specified (insufficient arguments).\n"
            "Usage: python {0} <input_tsv_file_path>\n"
            "Example: python {0} C:\\Data\\Sheet1_2025_11_18.tsv"
        ).format(pszProgramName)

        write_error_tsv(
            pszErrorFileFullPath,
            pszErrorMessage,
        )
        return

    pszInputFileFullPath: str = sys.argv[1]

    make_unique_staff_code_tsv_from_sheet1_tsv(pszInputFileFullPath)


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: make_staff_code_range.py
# ///////////////////////////////////////////////////////////////
pszSource_make_staff_code_range_py: str = r'''###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
# //////////////////////////////

# -*- coding: utf-8 -*-
# ///////////////////////////////////////////////////////////////
#  make_staff_code_range.py
#  （Sheet1.tsv からスタッフコードごとの行範囲一覧を作成する）
# ///////////////////////////////////////////////////////////////

import os
import sys
from typing import Dict, List, Tuple

import pandas as pd
from pandas import DataFrame, Series


# ---------------------------------------------------------------
# エラーメッセージを TSV として書き出す共通関数
# ---------------------------------------------------------------
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    """
    指定パスにエラーメッセージだけを書き込む。
    末尾に改行が無ければ追加する。
    """
    pszDirectoryFullPath: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectoryFullPath) > 0:
        os.makedirs(pszDirectoryFullPath, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)
        if not pszErrorMessage.endswith("\n"):
            objFile.write("\n")


# ---------------------------------------------------------------
# 出力ファイルパスを構築する
#   例:
#     入力: C:\Data\Sheet1_2025_11_18.tsv
#     出力: C:\Data\Sheet1_2025_11_18_staff_code_range.tsv
# ---------------------------------------------------------------
def build_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
    pszBaseFileName: str = os.path.basename(pszInputFileFullPath)
    pszRootName: str
    pszExt: str
    pszRootName, pszExt = os.path.splitext(pszBaseFileName)

    pszOutputFileName: str = pszRootName + "_staff_code_range.tsv"

    if len(pszDirectoryFullPath) == 0:
        return pszOutputFileName

    return os.path.join(pszDirectoryFullPath, pszOutputFileName)


# ---------------------------------------------------------------
# スタッフコード列から
#   ・ユニークなスタッフコード一覧
#   ・各コードの (最初の行インデックス, 最後の行インデックス)
# を求める
# ---------------------------------------------------------------
def analyze_staff_code_column(
    objSeriesStaffCode: Series,
) -> Tuple[List[str], Dict[str, Tuple[int, int]]]:
    """
    B列(スタッフコード列)を解析して、
      ・ユニークなスタッフコード一覧(出現順)
      ・各コードの最初と最後の DataFrame 行インデックス
    を返す。
    空文字列は無視する。
    """
    objListUniqueStaffCode: List[str] = []
    objDictCodeToRange: Dict[str, Tuple[int, int]] = {}

    iRowCount: int = objSeriesStaffCode.shape[0]
    for iRowIndex in range(iRowCount):
        objValue: object = objSeriesStaffCode.iat[iRowIndex]
        pszRaw: str = "" if objValue is None else str(objValue)
        pszCode: str = pszRaw.strip()

        # 空白のみ（または完全な空）のセルはスキップ
        if pszCode == "":
            continue

        # 初めて登場したコードはユニークリストに追加
        if pszCode not in objDictCodeToRange:
            objListUniqueStaffCode.append(pszCode)
            objDictCodeToRange[pszCode] = (iRowIndex, iRowIndex)
        else:
            # 既に辞書にある場合は「最後のインデックス」だけ更新
            iFirstIndex: int
            iLastIndex: int
            iFirstIndex, iLastIndex = objDictCodeToRange[pszCode]
            objDictCodeToRange[pszCode] = (iFirstIndex, iRowIndex)

    return objListUniqueStaffCode, objDictCodeToRange


# ---------------------------------------------------------------
# 実処理本体
#   ・Sheet1.tsv を読み込み
#   ・B列（2列目）のスタッフコードごとに
#       開始行・終了行を求める
#   ・*_staff_code_range.tsv として出力
# ---------------------------------------------------------------
def make_staff_code_range_tsv_from_sheet1_tsv(
    pszInputFileFullPath: str,
) -> None:
    # 入力ファイル存在チェック
    if not os.path.isfile(pszInputFileFullPath):
        pszDirectoryFullPath: str = os.path.dirname(pszInputFileFullPath)
        pszBaseFileName: str = os.path.basename(pszInputFileFullPath)
        pszBase: str
        pszExt: str
        pszBase, pszExt = os.path.splitext(pszBaseFileName)

        pszOutputFileFullPath: str = os.path.join(
            pszDirectoryFullPath,
            pszBase + "_error.tsv",
        )
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: input TSV file not found. Path = " + pszInputFileFullPath + "\n",
        )
        return

    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            encoding="utf-8",
            dtype=str,
            keep_default_na=False,
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading TSV for staff code range. "
            "Detail = " + str(objException) + "\n",
        )
        return

    # 列数チェック（B列が存在するか: 少なくとも 2 列必要）
    iColumnCount: int = objDataFrameInput.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: required column B does not exist (need at least 2 columns). "
            "ColumnCount = {0}\n".format(iColumnCount),
        )
        return

    # B列（スタッフコード列）のヘッダ名とデータ列を取得
    objColumnNameList: List[str] = list(objDataFrameInput.columns)
    pszStaffCodeColumnName: str = objColumnNameList[1]
    objSeriesStaffCode: Series = objDataFrameInput.iloc[:, 1]

    # スタッフコード列を解析して
    #   ・ユニークなスタッフコード一覧
    #   ・各コードの (最初の行インデックス, 最後の行インデックス)
    # を取得
    try:
        objListUniqueStaffCode: List[str]
        objDictCodeToRange: Dict[str, Tuple[int, int]]
        objListUniqueStaffCode, objDictCodeToRange = analyze_staff_code_column(
            objSeriesStaffCode
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while analyzing staff code column. "
            "Detail = " + str(objException) + "\n",
        )
        return

    # 有効なスタッフコードが 1 件も無い場合はエラーとする
    if len(objListUniqueStaffCode) == 0:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: no valid staff code found in column B.\n",
        )
        return

    # 出力用のデータを準備
    objListOutputCode: List[str] = []
    objListOutputStartRow: List[int] = []
    objListOutputEndRow: List[int] = []

    # DataFrame の行インデックス 0 が Excel 行 2 に対応するので
    # Excel 行番号 = インデックス + 2
    for pszCode in objListUniqueStaffCode:
        if pszCode not in objDictCodeToRange:
            write_error_tsv(
                pszOutputFileFullPath,
                "Error: internal inconsistency. Code not found in range map. "
                "Code = " + pszCode + "\n",
            )
            return

        iFirstIndex: int
        iLastIndex: int
        iFirstIndex, iLastIndex = objDictCodeToRange[pszCode]

        iStartRow: int = iFirstIndex + 2
        iEndRow: int = iLastIndex + 2

        objListOutputCode.append(pszCode)
        objListOutputStartRow.append(iStartRow)
        objListOutputEndRow.append(iEndRow)

    # 出力 DataFrame を構築
    pszStartColumnName: str = "開始行"
    pszEndColumnName: str = "終了行"

    objDataFrameOutput: DataFrame = DataFrame(
        {
            pszStaffCodeColumnName: objListOutputCode,
            pszStartColumnName: objListOutputStartRow,
            pszEndColumnName: objListOutputEndRow,
        }
    )

    # TSV として書き出し
    try:
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing staff code range TSV. "
            "Detail = " + str(objException) + "\n",
        )
        return


# ---------------------------------------------------------------
# main（引数チェック）
# ---------------------------------------------------------------
def main() -> None:
    """
    引数チェックを行い、実処理本体を呼び出す。
    """
    iArgCount: int = len(sys.argv)

    # 引数不足時は専用のエラー TSV を出力して終了
    if iArgCount < 2:
        pszProgramName: str = os.path.basename(sys.argv[0])

        pszLine1: str = (
            "Error: input TSV file path is not specified (insufficient arguments)."
        )
        pszLine2: str = "Usage: python {0} <input_tsv_file_path>".format(
            pszProgramName
        )
        pszLine3: str = (
            "Example: python {0} C:\\Data\\Sheet1_2025_11_18.tsv".format(
                pszProgramName
            )
        )

        write_error_tsv(
            "make_staff_code_rang_error_argument.tsv",
            pszLine1 + "\n" + pszLine2 + "\n" + pszLine3 + "\n",
        )
        return

    pszInputFileFullPath: str = sys.argv[1]

    make_staff_code_range_tsv_from_sheet1_tsv(pszInputFileFullPath)


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: make_sheet6_from_sheet4.py
# ///////////////////////////////////////////////////////////////
pszSource_make_sheet6_from_sheet4_py: str = r'''# -*- coding: utf-8 -*-
"""
 make_sheet6_from_sheet4.py
 （Sheet4.tsv と Sheet4_staff_code_range.tsv から Sheet6.tsv を作成する）

 目的:
   ・Sheet4.tsv（ジョブカン工数明細）と
     Sheet4_staff_code_range.tsv（スタッフ毎の行範囲）を読み込み、
     「スタッフごとに使用しているプロジェクト名一覧」を横方向に並べた
     Sheet6.tsv を作成する。

 前提:
   ・Sheet4.tsv は、Excel の Sheet4 を .tsv 化したもの。
     ヘッダー行あり、区切りはタブ、UTF-8。
     少なくとも次の列が存在すること:
       - "スタッフコード"
       - "プロジェクト名"

   ・Sheet4_staff_code_range.tsv は、Excel 上の
     「スタッフコード」「開始行」「終了行」の 3 列を
     そのまま .tsv 化したもの。

     A列: スタッフコード
     B列: 開始行（Excel の行番号。=MATCH(A2, Sheet4!B:B, 0) の結果など）
     C列: 終了行（次のスタッフの開始行-1 等。ROW() 由来の Excel 行番号）

     ※開始行・終了行は「Excel の行番号（1 始まり・ヘッダーを含む）」とする。
       Python 側では、ヘッダーを除いた DataFrame の行番号に変換して用いる。

 出力:
   ・Sheet6.tsv（同じフォルダに作成）
     1行目: 1,2,3,...（スタッフ列の通し番号）
     2行目: スタッフコード（Sheet4_staff_code_range.tsv の 1 列目の値）
     3行目以降: 各列が 1 人のスタッフを表し、そのスタッフの
                 「プロジェクト名」一覧（昇順ソート＋重複削除済み）を縦方向に並べる。

 実行例:
   python make_sheet6_from_sheet4.py Sheet4.tsv Sheet4_staff_code_range.tsv

 ===============================================================
"""

import os
import sys
from typing import List

import pandas as pd
from pandas import DataFrame


# ---------------------------------------------------------------
# 指定したパスにエラーTSVを書き込む（既存スクリプト踏襲）
# ---------------------------------------------------------------
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage + "\n")


# ---------------------------------------------------------------
# 出力ファイルパスを構築（Sheet4.tsv のフォルダに Sheet6.tsv を出力）
# ---------------------------------------------------------------
def build_output_file_full_path(
    pszSheet4FileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet6.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# 0 始まり列インデックス → Excel 列名（A, B, ..., Z, AA, AB, ...）
# （今回は 2 行目に式は出力しないが、既存コードをそのまま残している）
# ---------------------------------------------------------------
def convert_column_index_to_excel_column_name(
    iColumnIndex: int,
) -> str:
    iWorkIndex: int = iColumnIndex
    pszColumnName: str = ""

    while True:
        iQuotient, iRemainder = divmod(iWorkIndex, 26)
        cChar: str = chr(ord("A") + iRemainder)
        pszColumnName = cChar + pszColumnName
        if iQuotient == 0:
            break
        iWorkIndex = iQuotient - 1

    return pszColumnName


# ---------------------------------------------------------------
# Sheet4.tsv と Sheet4_staff_code_range.tsv から Sheet6.tsv を作成する
# ---------------------------------------------------------------
def make_sheet6_from_sheet4(
    pszSheet4FileFullPath: str,
    pszRangeFileFullPath: str,
) -> None:

    # 出力ファイルパス（エラー時もこのフォルダを基準にする）
    pszOutputFileFullPath: str = build_output_file_full_path(pszSheet4FileFullPath)
    pszErrorFileFullPath: str = pszOutputFileFullPath.replace(".tsv", "_error.tsv")

    # 入力ファイル存在チェック（Sheet4）
    if not os.path.isfile(pszSheet4FileFullPath):
        write_error_tsv(
            pszSheet4FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: Sheet4 TSV file not found. Path = {0}".format(pszSheet4FileFullPath),
        )
        return

    # 入力ファイル存在チェック（Sheet4_staff_code_range）
    if not os.path.isfile(pszRangeFileFullPath):
        write_error_tsv(
            pszRangeFileFullPath.replace(".tsv", "_error.tsv"),
            "Error: Sheet4_staff_code_range TSV file not found. Path = {0}".format(pszRangeFileFullPath),
        )
        return

    # Sheet4.tsv 読み込み
    try:
        objDataFrameSheet4: DataFrame = pd.read_csv(
            pszSheet4FileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading Sheet4 TSV. Detail = {0}".format(objException),
        )
        return

    # 必須列チェック（スタッフコード・プロジェクト名）
    objSheet4Columns: List[str] = list(objDataFrameSheet4.columns)
    if ("スタッフコード" not in objSheet4Columns) or ("プロジェクト名" not in objSheet4Columns):
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: required columns not found in Sheet4 TSV. "
            "Required columns: スタッフコード, プロジェクト名. "
            "Columns = {0}".format(", ".join(objSheet4Columns)),
        )
        return

    # Sheet4_staff_code_range.tsv 読み込み
    try:
        objDataFrameRange: DataFrame = pd.read_csv(
            pszRangeFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while reading Sheet4_staff_code_range TSV. Detail = {0}".format(objException),
        )
        return

    # 列数チェック（少なくとも 3 列: スタッフコード / 開始行 / 終了行）
    iRangeColumnCount: int = objDataFrameRange.shape[1]
    if iRangeColumnCount < 3:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: Sheet4_staff_code_range TSV must have at least 3 columns "
            "(staff_code, start_row, end_row). ColumnCount = {0}".format(iRangeColumnCount),
        )
        return

    # 開始行・終了行を数値に変換（Excel の行番号 → 後で DataFrame の行番号に変換）
    try:
        objDataFrameRange = objDataFrameRange.copy()
        # 1列目: スタッフコード
        # 2列目: 開始行（Excel 行番号）
        # 3列目: 終了行（Excel 行番号）
        objDataFrameRange["__start_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 1],
            errors="coerce",
        )
        objDataFrameRange["__end_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 2],
            errors="coerce",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while converting start/end rows to numeric. Detail = {0}".format(
                objException
            ),
        )
        return

    # NaN チェック
    if objDataFrameRange["__start_row_excel__"].isna().any() or objDataFrameRange["__end_row_excel__"].isna().any():
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: start_row or end_row contains non-numeric value in Sheet4_staff_code_range TSV.",
        )
        return

    # スタッフごとのプロジェクト名一覧（列ごと）とスタッフコード一覧を作成
    objListProjectListPerStaff: List[List[str]] = []
    objListStaffCode: List[str] = []

    iSheet4RowCount: int = objDataFrameSheet4.shape[0]

    for iIndex, objRow in objDataFrameRange.iterrows():
        # スタッフコード（そのまま文字列として保持）
        pszStaffCode: str = str(objRow.iloc[0])
        objListStaffCode.append(pszStaffCode)

        # Excel 上の行番号（ヘッダーを含む 1 始まり）
        iStartRowExcel: int = int(objRow["__start_row_excel__"])
        iEndRowExcel: int = int(objRow["__end_row_excel__"])

        # Excel 行番号 → DataFrame の行インデックスに変換
        # Excel: 1 行目 = ヘッダー, 2 行目 = DataFrame index 0
        # → DataFrame index = ExcelRow - 2
        iStartIndex: int = iStartRowExcel - 2
        iEndIndex: int = iEndRowExcel - 2

        # 範囲の妥当性チェック
        if (iStartIndex < 0) or (iEndIndex < 0) or (iStartIndex > iEndIndex) or (iEndIndex >= iSheet4RowCount):
            write_error_tsv(
                pszErrorFileFullPath,
                "Error: invalid row range for staff code {0}. "
                "StartExcel={1}, EndExcel={2}, StartIndex={3}, EndIndex={4}, Sheet4RowCount={5}".format(
                    pszStaffCode,
                    iStartRowExcel,
                    iEndRowExcel,
                    iStartIndex,
                    iEndIndex,
                    iSheet4RowCount,
                ),
            )
            return

        # 該当範囲の行を抽出
        objDataFrameSub: DataFrame = objDataFrameSheet4.iloc[iStartIndex : iEndIndex + 1]

        # 念のためスタッフコードでフィルタ（行範囲が正しければ全て一致するはず）
        objDataFrameSub = objDataFrameSub[objDataFrameSub["スタッフコード"] == pszStaffCode]

        # プロジェクト名列だけを取り出し、NaN・空白を除去
        objSeriesPj: pd.Series = objDataFrameSub["プロジェクト名"].dropna()
        objSeriesPj = objSeriesPj.astype(str)
        objSeriesPj = objSeriesPj[objSeriesPj.str.strip() != ""]

        # 昇順ソート → 重複削除（Excel の UNIQUE(SORT()) と同等の結果を目指す）
        objSeriesPjSorted: pd.Series = objSeriesPj.sort_values()
        objSeriesPjUnique: pd.Series = objSeriesPjSorted.drop_duplicates()

        objProjectList: List[str] = objSeriesPjUnique.tolist()
        objListProjectListPerStaff.append(objProjectList)

    # スタッフ数
    iStaffCount: int = len(objListProjectListPerStaff)

    # スタッフが 0 人なら空ファイルを出力して終了
    if iStaffCount == 0:
        try:
            objEmpty: DataFrame = DataFrame([])
            objEmpty.to_csv(
                pszOutputFileFullPath,
                sep="\t",
                index=False,
                header=False,
                encoding="utf-8",
                lineterminator="\n",
            )
        except Exception as objException:
            write_error_tsv(
                pszErrorFileFullPath,
                "Error: unexpected exception while writing empty Sheet6 TSV. Detail = {0}".format(objException),
            )
        return

    # 各スタッフの最大プロジェクト数
    iMaxProjectCount: int = max(len(objList) for objList in objListProjectListPerStaff)

    # Sheet6 全体行の 2 次元リストを構築
    objRows: List[List[str]] = []

    # 1 行目: 1,2,3,...（スタッフ列の通し番号）
    objRow1: List[str] = [str(iIndex + 1) for iIndex in range(iStaffCount)]
    objRows.append(objRow1)

    # 2 行目: スタッフコード（Sheet4_staff_code_range.tsv の 1 列目の値）
    objRow2: List[str] = [str(pszCode) for pszCode in objListStaffCode]
    objRows.append(objRow2)

    # 3 行目以降: 各列ごとに「プロジェクト名」一覧を縦に並べる
    for iPjIndex in range(iMaxProjectCount):
        objRow: List[str] = []
        for iStaffIndex in range(iStaffCount):
            objProjectList: List[str] = objListProjectListPerStaff[iStaffIndex]
            if iPjIndex < len(objProjectList):
                objRow.append(objProjectList[iPjIndex])
            else:
                objRow.append("")
        objRows.append(objRow)

    # DataFrame に変換して TSV 出力（ヘッダーなし）
    try:
        objDataFrameOutput: DataFrame = DataFrame(objRows)
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszErrorFileFullPath,
            "Error: unexpected exception while writing Sheet6 TSV. Detail = {0}".format(objException),
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足 → _error_argument.tsv を出力
    if iArgCount < 3:
        pszProgram: str = os.path.basename(sys.argv[0])
        pszErrorFile: str = "make_sheet6_from_sheet4_error_argument.tsv"

        write_error_tsv(
            pszErrorFile,
            "Error: input TSV file paths are not specified (insufficient arguments).\n"
            "Usage: python {0} <Sheet4_tsv_file_path> <Sheet4_staff_code_range_tsv_file_path>\n"
            "Example: python {0} Sheet4.tsv Sheet4_staff_code_range.tsv".format(pszProgram),
        )
        return

    pszSheet4FileFullPath: str = sys.argv[1]
    pszRangeFileFullPath: str = sys.argv[2]

    make_sheet6_from_sheet4(
        pszSheet4FileFullPath,
        pszRangeFileFullPath,
    )


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
'''

# ///////////////////////////////////////////////////////////////
# EMBEDDED SOURCE: make_sheet789_from_sheet4.py
# ///////////////////////////////////////////////////////////////
pszSource_make_sheet789_from_sheet4_py: str = r'''# -*- coding: utf-8 -*-
# //////////////////////////////
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
# //////////////////////////////

# -*- coding: utf-8 -*-
# ///////////////////////////////////////////////////////////////
#  make_sheet789_from_sheet4.py
#  （Sheet4.tsv / Sheet4_staff_code_range.tsv / Sheet6.tsv から
#    Sheet7.tsv と Sheet8.tsv と Sheet9.tsv を作成する・Sheet6主体アルゴリズム版）
#
#  目的:
#    ・Sheet6.tsv を「マスタ」として用い、
#      - Sheet6 の 2 行目に並ぶスタッフコードの順番をそのまま採用し、
#      - 各列の 3 行目以降に並ぶプロジェクト名一覧をそのまま利用して、
#      Sheet7.tsv / Sheet8.tsv を作成する。
#
#    ・工数の実データは Sheet4.tsv から読み取り、
#      スタッフコードとプロジェクト名毎に工数を合計する。
#
#    ・Sheet9.tsv には「氏名・スタッフコード」の一覧を重複なくもれなく出力する。
#      1 行目はヘッダー行（氏名, スタッフコード）とする。
#
#  各シートのイメージ:
#
#    Sheet6.tsv（make_sheet6_from_sheet4.py の出力）
#      1 行目: 1, 2, 3, ...                   ← スタッフ通し番号
#      2 行目: スタッフコード(1列目), スタッフコード(2列目), ...
#      3 行目以降: 各列のスタッフが関わったプロジェクト名一覧（縦方向）
#
#    Sheet7.tsv（本スクリプトの出力）
#      1 列目: プロジェクト名（Sheet6 の A,B,C,... 列を縦に並べたもの）
#      2 列目: スタッフコード
#               └ 各スタッフの行すべてに同じスタッフコードを入れる
#      3 列目: 合計工数時間（hh:mm:ss 形式、24 時間超も総時間数）
#
#    Sheet8.tsv（本スクリプトの出力）
#      1 列目: 氏名
#               └ 各スタッフの「最初の行」にだけ値を入れ、それ以外は空文字
#      2 列目: プロジェクト名（Sheet7.tsv の 1 列目と同じ）
#      3 列目: スタッフコード（Sheet7.tsv の 2 列目と同じ。全行に値あり）
#      4 列目: 合計工数時間（Sheet7.tsv の 3 列目と同じ）
#
#    ※Sheet8.tsv の 2～4 列目は、Sheet7.tsv の 1～3 列目と同じ内容になる。
#
#    Sheet9.tsv（本スクリプトの出力）
#      1 行目: 氏名, スタッフコード（ヘッダー）
#      2 行目以降:
#        ・氏名, スタッフコード
#        ・同じスタッフコードは 1 回だけ出力する（重複なし）。
#        ・Sheet6 のスタッフコード一覧と Sheet4 の「スタッフコード→氏名」の
#          両方を元に、可能な限り「もれなく」一覧化する。
#
#  前提:
#    ・Sheet4.tsv は、Excel の Sheet4 を .tsv 化したもの。
#      ヘッダー行あり、区切りはタブ。
#      エンコーディングは UTF-8(BOM付き) または cp932(Shift-JIS) のいずれか。
#      少なくとも次の列が存在すること:
#        - "スタッフコード"
#        - "プロジェクト名"
#        - "工数"
#        - 氏名列（"姓 名" または "氏名" のいずれか）
#
#    ・Sheet4_staff_code_range.tsv は、Excel 上の
#      「スタッフコード」「開始行」「終了行」の 3 列を
#      そのまま .tsv 化したもの。
#
#      A列: スタッフコード
#      B列: 開始行（Excel 行番号。=MATCH(A2, Sheet4!B:B, 0) の結果など）
#      C列: 終了行（次のスタッフの開始行-1 等。ROW() 由来の Excel 行番号）
#
#      ※開始行・終了行は「Excel の行番号（1 始まり・ヘッダーを含む）」とする。
#        Python 側では、ッダーを除いた DataFrame の行番号に変換して用いる。
#
#    ・Sheet6.tsv は、make_sheet6_from_sheet4.py によって作成された形式
#      （1 行目: 通し番号, 2 行目: スタッフコード, 3 行目以降: プロジェクト名一覧）であること。
#
#  実行例:
#    python make_sheet789_from_sheet4.py Sheet4.tsv Sheet4_staff_code_range.tsv Sheet6.tsv
#
#  出力:
#    ・Sheet7.tsv（Sheet4.tsv と同じフォルダ）
#    ・Sheet8.tsv（Sheet4.tsv と同じフォルダ）
#    ・Sheet9.tsv（Sheet4.tsv と同じフォルダ）
#
# ///////////////////////////////////////////////////////////////

import os
import sys
from typing import Dict, List, Tuple, Set

import pandas as pd
from pandas import DataFrame


# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
# //  TSV を UTF-8(BOM付き) 優先・cp932 併用で読み込む関数。
# //  bHasHeader=True のときは 1 行目をヘッダーとして扱う。
# //  bHasHeader=False のときはヘッダー無し（header=None）で読み込む。
# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
def read_tsv_with_encoding_candidates(
    pszInputFileFullPath: str,
    bHasHeader: bool,
) -> DataFrame:
    objEncodingCandidateList: List[str] = [
        "utf-8-sig",
        "cp932",
    ]
    bLoaded: bool = False
    objLastDecodeError: Exception | None = None
    objDataFrameResult: DataFrame | None = None

    for pszEncoding in objEncodingCandidateList:
        try:
            if bHasHeader:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath,
                    sep="\t",
                    dtype=str,
                    encoding=pszEncoding,
                    engine="python",
                )
            else:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath,
                    sep="\t",
                    dtype=str,
                    header=None,
                    encoding=pszEncoding,
                    engine="python",
                )
            bLoaded = True
            break
        except UnicodeDecodeError as objDecodeError:
            objLastDecodeError = objDecodeError
            continue

    if not bLoaded or objDataFrameResult is None:
        if objLastDecodeError is not None:
            raise objLastDecodeError
        raise UnicodeDecodeError(
            "utf-8-sig",
            b"",
            0,
            1,
            "cannot decode TSV with utf-8-sig nor cp932",
        )

    return objDataFrameResult


# ---------------------------------------------------------------
# 指定したパスにエラーTSVを書き込む（既存スタイル踏襲）
# ---------------------------------------------------------------
def write_error_tsv(
    pszOutputFileFullPath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
    if len(pszDirectory) > 0:
        os.makedirs(pszDirectory, exist_ok=True)

    with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage + "\n")


# ---------------------------------------------------------------
# Sheet7.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet7(
    pszSheet4FileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet7.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# Sheet8.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet8(
    pszSheet4FileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet8.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# Sheet9.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet9(
    pszSheet4FileFullPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet9.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
# //  文字列の時間（hh:mm:ss または hh:mm）を「秒数(int)」に変換する関数。
# //  不正な形式や空文字の場合は 0 秒として扱う。
# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
def convert_time_string_to_seconds(
    pszTimeText: str,
) -> int:
    if pszTimeText is None:
        return 0

    pszWork: str = str(pszTimeText).strip()
    if len(pszWork) == 0:
        return 0

    objParts: List[str] = pszWork.split(":")
    try:
        if len(objParts) == 3:
            iHour: int = int(objParts[0])
            iMinute: int = int(objParts[1])
            iSecond: int = int(objParts[2])
        elif len(objParts) == 2:
            iHour: int = int(objParts[0])
            iMinute: int = int(objParts[1])
            iSecond: int = 0
        else:
            return 0
    except ValueError:
        return 0

    iTotalSeconds: int = iHour * 3600 + iMinute * 60 + iSecond
    return iTotalSeconds


# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
# //  秒数(int)を「hh:mm:ss」形式の文字列に変換する関数。
# //  24 時間を超えても、そのまま総時間数として出力する。
# ///////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////
def convert_seconds_to_time_string(
    iTotalSeconds: int,
) -> str:
    if iTotalSeconds <= 0:
        return "0:00:00"

    iHour: int = iTotalSeconds // 3600
    iRemain: int = iTotalSeconds % 3600
    iMinute: int = iRemain // 60
    iSecond: int = iRemain % 60

    pszTimeText: str = "{0}:{1:02d}:{2:02d}".format(iHour, iMinute, iSecond)
    return pszTimeText


# ---------------------------------------------------------------
# Sheet4.tsv, Sheet4_staff_code_range.tsv, Sheet6.tsv から
# Sheet7.tsv と Sheet8.tsv と Sheet9.tsv を作成する（Sheet6 主体アルゴリズム）
# ---------------------------------------------------------------
def make_sheet789_from_sheet4(
    pszSheet4FileFullPath: str,
    pszRangeFileFullPath: str,
    pszSheet6FileFullPath: str,
) -> None:

    # 出力ファイルパス（エラー時もこのフォルダを基準にする）
    pszSheet7FileFullPath: str = build_output_file_full_path_for_sheet7(pszSheet4FileFullPath)
    pszSheet7ErrorFileFullPath: str = pszSheet7FileFullPath.replace(".tsv", "_error.tsv")
    pszSheet8FileFullPath: str = build_output_file_full_path_for_sheet8(pszSheet4FileFullPath)
    pszSheet8ErrorFileFullPath: str = pszSheet8FileFullPath.replace(".tsv", "_error.tsv")
    pszSheet9FileFullPath: str = build_output_file_full_path_for_sheet9(pszSheet4FileFullPath)
    pszSheet9ErrorFileFullPath: str = pszSheet9FileFullPath.replace(".tsv", "_error.tsv")

    # 入力ファイル存在チェック（Sheet4）
    if not os.path.isfile(pszSheet4FileFullPath):
        write_error_tsv(
            pszSheet4FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: Sheet4 TSV file not found. Path = {0}".format(pszSheet4FileFullPath),
        )
        return

    # 入力ファイル存在チェック（Sheet4_staff_code_range）
    if not os.path.isfile(pszRangeFileFullPath):
        write_error_tsv(
            pszRangeFileFullPath.replace(".tsv", "_error.tsv"),
            "Error: Sheet4_staff_code_range TSV file not found. Path = {0}".format(pszRangeFileFullPath),
        )
        return

    # 入力ファイル存在チェック（Sheet6）
    if not os.path.isfile(pszSheet6FileFullPath):
        write_error_tsv(
            pszSheet6FileFullPath.replace(".tsv", "_error.tsv"),
            "Error: Sheet6 TSV file not found. Path = {0}".format(pszSheet6FileFullPath),
        )
        return

    # -----------------------------------------------------------
    # Sheet4.tsv 読み込み（UTF-8-sig → cp932 の順で試す）
    # -----------------------------------------------------------
    try:
        objDataFrameSheet4: DataFrame = read_tsv_with_encoding_candidates(
            pszSheet4FileFullPath,
            True,
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while reading Sheet4 TSV. Detail = {0}".format(objException),
        )
        return

    objSheet4Columns: List[str] = list(objDataFrameSheet4.columns)
    if (
        ("スタッフコード" not in objSheet4Columns)
        or ("プロジェクト名" not in objSheet4Columns)
        or ("工数" not in objSheet4Columns)
    ):
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: required columns not found in Sheet4 TSV. "
            "Required columns: スタッフコード, プロジェクト名, 工数. "
            "Columns = {0}".format(", ".join(objSheet4Columns)),
        )
        return

    # 氏名列の候補を探す（"姓 名" または "氏名"）
    pszNameColumn: str = ""
    if "姓 名" in objSheet4Columns:
        pszNameColumn = "姓 名"
    elif "氏名" in objSheet4Columns:
        pszNameColumn = "氏名"

    # 工数列を秒数に変した補助列を追加
    try:
        objDataFrameSheet4 = objDataFrameSheet4.copy()
        objDataFrameSheet4["__time_seconds__"] = objDataFrameSheet4["工数"].apply(
            lambda pszValue: convert_time_string_to_seconds(pszValue),
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while converting time text to seconds. Detail = {0}".format(objException),
        )
        return

    iSheet4RowCount: int = objDataFrameSheet4.shape[0]

    # -----------------------------------------------------------
    # Sheet4_staff_code_range.tsv 読み込み
    #   → スタッフコード → (開始インデックス, 終了インデックス) の辞書を作る
    # -----------------------------------------------------------
    try:
        objDataFrameRange: DataFrame = read_tsv_with_encoding_candidates(
            pszRangeFileFullPath,
            True,
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while reading Sheet4_staff_code_range TSV. Detail = {0}".format(
                objException,
            ),
        )
        return

    if objDataFrameRange.shape[1] < 3:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: Sheet4_staff_code_range TSV must have at least 3 columns (staff_code, start_row, end_row).",
        )
        return

    try:
        objDataFrameRange = objDataFrameRange.copy()
        objDataFrameRange["__start_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 1], errors="coerce"
        )
        objDataFrameRange["__end_row_excel__"] = pd.to_numeric(
            objDataFrameRange.iloc[:, 2], errors="coerce"
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while converting start/end rows to numeric. Detail = {0}".format(
                objException,
            ),
        )
        return

    if objDataFrameRange["__start_row_excel__"].isna().any() or objDataFrameRange["__end_row_excel__"].isna().any():
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: start_row or end_row contains non-numeric value in Sheet4_staff_code_range TSV.",
        )
        return

    # スタッフコード → (開始Index, 終了Index) の辞書を構築
    objDictStaffCodeToRange: Dict[str, Tuple[int, int]] = {}
    for iIndex, objRow in objDataFrameRange.iterrows():
        pszStaffCodeRange: str = str(objRow.iloc[0]).strip()
        if len(pszStaffCodeRange) == 0:
            continue

        iStartRowExcel: int = int(objRow["__start_row_excel__"])
        iEndRowExcel: int = int(objRow["__end_row_excel__"])

        iStartIndex: int = iStartRowExcel - 2  # 1 行目: ヘッダー, 2 行目: index 0
        iEndIndex: int = iEndRowExcel - 2

        if (iStartIndex < 0) or (iEndIndex < 0) or (iStartIndex > iEndIndex) or (iEndIndex >= iSheet4RowCount):
            write_error_tsv(
                pszSheet7ErrorFileFullPath,
                "Error: invalid row range for staff code {0}. "
                "StartExcel={1}, EndExcel={2}, StartIndex={3}, EndIndex={4}, Sheet4RowCount={5}".format(
                    pszStaffCodeRange,
                    iStartRowExcel,
                    iEndRowExcel,
                    iStartIndex,
                    iEndIndex,
                    iSheet4RowCount,
                ),
            )
            return

        # 同じスタッフコードが複数行に出てきた場合は最初のものを優先
        if pszStaffCodeRange not in objDictStaffCodeToRange:
            objDictStaffCodeToRange[pszStaffCodeRange] = (iStartIndex, iEndIndex)

    # -----------------------------------------------------------
    # Sheet4 から「スタッフコード → 氏名」の辞書を作る（あれば）
    # -----------------------------------------------------------
    objDictStaffCodeToName: Dict[str, str] = {}
    if pszNameColumn != "":
        for iIndex, objRow in objDataFrameSheet4.iterrows():
            pszStaffCodeFromSheet4: str = str(objRow["スタッフコード"]).strip()
            if len(pszStaffCodeFromSheet4) == 0:
                continue
            pszStaffName: str = str(objRow[pszNameColumn]).strip()
            if pszStaffCodeFromSheet4 not in objDictStaffCodeToName:
                objDictStaffCodeToName[pszStaffCodeFromSheet4] = pszStaffName

    # -----------------------------------------------------------
    # Sheet6.tsv 読み込み（スタッフコード順・プロジェクト一覧を利用）
    # -----------------------------------------------------------
    try:
        # ヘッダー行なしとして読み込む（1 行目〜をそのまま扱う）
        objDataFrameSheet6: DataFrame = read_tsv_with_encoding_candidates(
            pszSheet6FileFullPath,
            False,
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while reading Sheet6 TSV. Detail = {0}".format(objException),
        )
        return

    if objDataFrameSheet6.shape[0] < 2:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: Sheet6 TSV must have at least 2 rows (1st: index, 2nd: staff_code).",
        )
        return

    iSheet6RowCount: int = objDataFrameSheet6.shape[0]
    iSheet6ColumnCount: int = objDataFrameSheet6.shape[1]

    # 2 行目(インデックス 1) のスタッフコード一覧を取得
    objListStaffCodeFromSheet6: List[str] = []
    for iColumnIndex in range(iSheet6ColumnCount):
        pszStaffCodeFromSheet6: str = ""
        objValue = objDataFrameSheet6.iat[1, iColumnIndex]
        if objValue is not None:
            pszStaffCodeFromSheet6 = str(objValue).strip()
        if len(pszStaffCodeFromSheet6) == 0:
            # 空列はスキップ
            continue
        objListStaffCodeFromSheet6.append(pszStaffCodeFromSheet6)

    # -----------------------------------------------------------
    # Sheet7 / Sheet8 / Sheet9 の出力行リストを構築
    # -----------------------------------------------------------
    objListOutputRowsSheet7: List[List[str]] = []
    objListOutputRowsSheet8: List[List[str]] = []
    objListOutputRowsSheet9: List[List[str]] = []

    # Sheet9 用: スタッフコード一覧を元に「氏名・スタッフコード」を作成
    objSetAddedStaffCodeForSheet9: Set[str] = set()

    for pszStaffCodeForSheet9 in objListStaffCodeFromSheet6:
        pszStaffNameForSheet9: str = objDictStaffCodeToName.get(pszStaffCodeForSheet9, "")
        objListOutputRowsSheet9.append(
            [pszStaffNameForSheet9, pszStaffCodeForSheet9],
        )
        objSetAddedStaffCodeForSheet9.add(pszStaffCodeForSheet9)

    # Sheet4 由来のスタッフコードで、Sheet6 に無いものがあれば末尾に追加
    for pszStaffCodeExtra, pszStaffNameExtra in objDictStaffCodeToName.items():
        if pszStaffCodeExtra in objSetAddedStaffCodeForSheet9:
            continue
        objListOutputRowsSheet9.append(
            [pszStaffNameExtra, pszStaffCodeExtra],
        )
        objSetAddedStaffCodeForSheet9.add(pszStaffCodeExtra)

    # スタッフコードを Sheet6 の並び順でループ（Sheet7 / Sheet8 用）
    for pszStaffCode in objListStaffCodeFromSheet6:
        # 該当スタッフの行範囲（Sheet4）を辞書から取得
        if pszStaffCode in objDictStaffCodeToRange:
            iStartIndex, iEndIndex = objDictStaffCodeToRange[pszStaffCode]
            objDataFrameSubStaff: DataFrame = objDataFrameSheet4.iloc[iStartIndex : iEndIndex + 1]
            objDataFrameSubStaff = objDataFrameSubStaff[objDataFrameSubStaff["スタッフコード"] == pszStaffCode]
        else:
            # 念のため、辞書に無い場合はSheet4全体からフィルタ
            objDataFrameSubStaff = objDataFrameSheet4[
                objDataFrameSheet4["スタッフコード"] == pszStaffCode
            ]

        if objDataFrameSubStaff.empty:
            # 該当スタッフの工数データが無ければスキップ
            continue

        pszStaffNameForSheet8: str = objDictStaffCodeToName.get(pszStaffCode, "")
        iRowIndexWithinStaff: int = 0

        # Sheet6 の列からプロジェクト名一覧を縦方向に取得
        #  1 行目: 通し番号
        #  2 行目: スタッフコード
        #  3 行目以降: プロジェクト名
        #
        # ここでは、全行を走査し、指定したスタッフコードの列だけを見る。
        for iColumnIndex in range(iSheet6ColumnCount):
            objValueStaffCodeAtColumn = objDataFrameSheet6.iat[1, iColumnIndex]
            pszStaffCodeAtColumn: str = ""
            if objValueStaffCodeAtColumn is not None:
                pszStaffCodeAtColumn = str(objValueStaffCodeAtColumn).strip()
            if pszStaffCodeAtColumn != pszStaffCode:
                # この列は別のスタッフなのでスキップ
                continue

            # この列が対象スタッフの列
            for iRowIndex in range(2, iSheet6RowCount):
                objValueProject = objDataFrameSheet6.iat[iRowIndex, iColumnIndex]
                pszProjectNameFromSheet6: str = ""
                if objValueProject is not None:
                    pszProjectNameFromSheet6 = str(objValueProject).strip()
                if len(pszProjectNameFromSheet6) == 0:
                    # 空セル（これ以降も空の可能性が高いが、全行チェックでも問題なし）
                    continue

                # このスタッフ・このプロジェクトについて、工数秒数を合計
                objDataFrameSubProject: DataFrame = objDataFrameSubStaff[
                    objDataFrameSubStaff["プロジェクト名"] == pszProjectNameFromSheet6
                ]

                if objDataFrameSubProject.empty:
                    # データ不整合で Sheet6 にだけプロジェクト名が存在する場合はスキップ
                    continue

                # __time_seconds__ を合計
                try:
                    iTotalSeconds: int = int(objDataFrameSubProject["__time_seconds__"].sum())
                except Exception:
                    iTotalSeconds = 0

                pszTimeTotal: str = convert_seconds_to_time_string(iTotalSeconds)

                # スタッフコードは各行すべてに出力
                pszStaffCodeForRow: str = pszStaffCode
                # 氏名はスタッフの最初の行だけ出力
                if iRowIndexWithinStaff == 0:
                    pszStaffNameForRow: str = pszStaffNameForSheet8
                else:
                    pszStaffNameForRow = ""

                # Sheet7 用の 1 行
                objListOutputRowsSheet7.append(
                    [pszProjectNameFromSheet6, pszStaffCodeForRow, pszTimeTotal],
                )

                # Sheet8 用の 1 行（Sheet7 の前に氏名列を追加）
                objListOutputRowsSheet8.append(
                    [pszStaffNameForRow, pszProjectNameFromSheet6, pszStaffCodeForRow, pszTimeTotal],
                )

                iRowIndexWithinStaff += 1

            # 同じスタッフコードの列は 1 列だけの想定なので break
            break

    # -----------------------------------------------------------
    # TSV 出力（Sheet7, Sheet8, Sheet9）
    # -----------------------------------------------------------
    try:
        objDataFrameOutputSheet7: DataFrame = DataFrame(objListOutputRowsSheet7)
        objDataFrameOutputSheet7.to_csv(
            pszSheet7FileFullPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath,
            "Error: unexpected exception while writing Sheet7 TSV. Detail = {0}".format(objException),
        )
        return

    try:
        objDataFrameOutputSheet8: DataFrame = DataFrame(objListOutputRowsSheet8)
        objDataFrameOutputSheet8.to_csv(
            pszSheet8FileFullPath,
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet8ErrorFileFullPath,
            "Error: unexpected exception while writing Sheet8 TSV. Detail = {0}".format(objException),
        )
        return

    try:
        # Sheet9 はヘッダー行あり（氏名, スタッフコード）
        objDataFrameOutputSheet9: DataFrame = DataFrame(
            objListOutputRowsSheet9,
            columns=["氏名", "スタッフコード"],
        )
        objDataFrameOutputSheet9.to_csv(
            pszSheet9FileFullPath,
            sep="\t",
            index=False,
            header=True,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet9ErrorFileFullPath,
            "Error: unexpected exception while writing Sheet9 TSV. Detail = {0}".format(objException),
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足 → _error_argument.tsv を出力
    if iArgCount < 4:
        pszProgram: str = os.path.basename(sys.argv[0])
        pszErrorFile: str = "make_sheet789_from_sheet4_error_argument.tsv"

        write_error_tsv(
            pszErrorFile,
            "Error: input TSV file paths are not specified (insufficient arguments).\n"
            "Usage: python {0} <Sheet4_tsv_file_path> <Sheet4_staff_code_range_tsv_file_path> <Sheet6_tsv_file_path>\n"
            "Example: python {0} Sheet4.tsv Sheet4_staff_code_range.tsv Sheet6.tsv".format(pszProgram),
        )
        return

    pszSheet4FileFullPath: str = sys.argv[1]
    pszRangeFileFullPath: str = sys.argv[2]
    pszSheet6FileFullPath: str = sys.argv[3]

    make_sheet789_from_sheet4(
        pszSheet4FileFullPath,
        pszRangeFileFullPath,
        pszSheet6FileFullPath,
    )


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
'''

# ///////////////////////////////////////////////////////////////
#
# 各スクリプトを独立名前空間で実行するためのヘルパー
#
# ///////////////////////////////////////////////////////////////
def create_module_from_source(
    pszModuleName: str,
    pszSourceCode: str,
) -> Dict[str, Any]:
    objGlobals: Dict[str, Any] = {
        "__name__": pszModuleName,
        "__file__": pszModuleName + ".py",
        "__package__": None,
    }
    exec(pszSourceCode, objGlobals)
    return objGlobals


# ///////////////////////////////////////////////////////////////
#
# エラー内容を UTF-8 テキストとして書き出す関数
#
# ///////////////////////////////////////////////////////////////
def write_error_text_utf8(
    pszOutputTextFilePath: str,
    pszErrorMessage: str,
) -> None:
    pszDirectory: str = os.path.dirname(pszOutputTextFilePath)
    if pszDirectory != "":
        os.makedirs(pszDirectory, exist_ok=True)

    with open(
        pszOutputTextFilePath,
        mode="w",
        encoding="utf-8",
        newline="",
    ) as objFile:
        objFile.write(pszErrorMessage)


# ///////////////////////////////////////////////////////////////
#
# Project_List.tsv 生成
#
# ///////////////////////////////////////////////////////////////
def convert_time_text_to_seconds_for_project_list(pszTimeText: str) -> int:
    pszText: str = str(pszTimeText or "").strip()
    if len(pszText) == 0:
        return 0

    objParts: List[str] = pszText.split(":")
    try:
        if len(objParts) == 2:
            iHours: int = int(objParts[0])
            iMinutes: int = int(objParts[1])
            return iHours * 3600 + iMinutes * 60

        if len(objParts) == 3:
            iHours: int = int(objParts[0])
            iMinutes: int = int(objParts[1])
            iSeconds: int = int(objParts[2])
            return iHours * 3600 + iMinutes * 60 + iSeconds
    except ValueError:
        return 0

    return 0


def format_seconds_to_h_mm_ss(iTotalSeconds: int) -> str:
    iSecondsSafe: int = max(int(iTotalSeconds or 0), 0)
    if iSecondsSafe == 0:
        return ""
    iHours: int = iSecondsSafe // 3600
    iMinutes: int = (iSecondsSafe % 3600) // 60
    iSeconds: int = iSecondsSafe % 60
    return f"{iHours}:{iMinutes:02}:{iSeconds:02}"


def _replace_raw_data_column_ranges(
    pszFormula: str,
    iLastRowNumber: int,
) -> str:
    objReplacements: List[Tuple[str, str]] = [
        ("$A:$A", f"$A$1:$A${iLastRowNumber}"),
        ("$D:$D", f"$D$1:$D${iLastRowNumber}"),
        ("$L:$L", f"$L$1:$L${iLastRowNumber}"),
    ]

    pszResult: str = pszFormula
    for pszTarget, pszReplacement in objReplacements:
        if pszTarget in pszResult:
            pszResult = pszResult.replace(pszTarget, pszReplacement)

    return pszResult


def _replace_raw_data_column_ranges_in_dataframe(
    objDataFrame: pd.DataFrame,
    iLastRowNumber: int,
) -> pd.DataFrame:
    if objDataFrame.empty:
        return objDataFrame

    return objDataFrame.applymap(
        lambda pszValue: _replace_raw_data_column_ranges(str(pszValue), iLastRowNumber)
    )


def make_project_list_tsv_from_raw_data(
    pszRawDataTsvPath: str,
    pszProjectListFormulaTsvPath: str,
    pszOutputTsvPath: str,
) -> None:
    try:
        if not os.path.isfile(pszRawDataTsvPath):
            with open(
                pszRawDataTsvPath.replace(".tsv", "_error.tsv"),
                "w",
                encoding="utf-8",
            ) as objFile:
                objFile.write(
                    "Error: input TSV file not found. Path = {0}".format(
                        pszRawDataTsvPath,
                    )
                )
            return

        if not os.path.isfile(pszProjectListFormulaTsvPath):
            with open(
                pszProjectListFormulaTsvPath.replace(".tsv", "_error.tsv"),
                "w",
                encoding="utf-8",
            ) as objFile:
                objFile.write(
                    "Error: input TSV file not found. Path = {0}".format(
                        pszProjectListFormulaTsvPath,
                    )
                )
            return

        objRawDataSheetForProjectList: pd.DataFrame = pd.read_csv(
            pszRawDataTsvPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
        objProjectListFormulaSheet: pd.DataFrame = pd.read_csv(
            pszProjectListFormulaTsvPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            header=None,
            engine="python",
        )

        objOutputProjectList: pd.DataFrame = pd.DataFrame(
            "",
            index=objProjectListFormulaSheet.index,
            columns=objProjectListFormulaSheet.columns,
        )

        objEvaluatedCellsProjectList: Dict[Tuple[str, int, int], Any] = {}
        objEvaluatingCellsProjectList: set[Tuple[str, int, int]] = set()

        class ProjectListCircularReferenceError(RuntimeError):
            pass

        def tokenize_formula_project_list(pszFormula: str) -> List[Tuple[str, str]]:
            objTokens: List[Tuple[str, str]] = []
            iPos: int = 0
            while iPos < len(pszFormula):
                ch: str = pszFormula[iPos]
                if ch.isspace():
                    iPos += 1
                    continue
                if ch.isdigit() or (ch == "." and iPos + 1 < len(pszFormula) and pszFormula[iPos + 1].isdigit()):
                    iStart: int = iPos
                    iPos += 1
                    while iPos < len(pszFormula) and (pszFormula[iPos].isdigit() or pszFormula[iPos] == "."):
                        iPos += 1
                    objTokens.append(("number", pszFormula[iStart:iPos]))
                    continue
                if ch == '"':
                    iPos += 1
                    iStart: int = iPos
                    while iPos < len(pszFormula) and pszFormula[iPos] != '"':
                        iPos += 1
                    objTokens.append(("string", pszFormula[iStart:iPos]))
                    iPos += 1
                    continue
                if ch.isalpha() or ch in ["_", "$", "!"]:
                    iStart: int = iPos
                    iPos += 1
                    while iPos < len(pszFormula) and (pszFormula[iPos].isalnum() or pszFormula[iPos] in ["_", "$", "!"]):
                        iPos += 1
                    objTokens.append(("ident", pszFormula[iStart:iPos]))
                    continue
                if ch in ["+", "-", "*", "/", "&"]:
                    objTokens.append(("op", ch))
                    iPos += 1
                    continue
                if ch in ["=", "<", ">"]:
                    if iPos + 1 < len(pszFormula) and pszFormula[iPos:iPos + 2] in ["<=", ">=", "<>"]:
                        objTokens.append(("cmp", pszFormula[iPos:iPos + 2]))
                        iPos += 2
                    else:
                        objTokens.append(("cmp", ch))
                        iPos += 1
                    continue
                if ch in ["(", ")", ",", ":"]:
                    objTokens.append(("symbol", ch))
                    iPos += 1
                    continue
                iPos += 1
            return objTokens

        def column_label_to_index_project_list(pszLabel: str) -> int:
            iResult: int = 0
            for ch in pszLabel.upper():
                if ch == "$":
                    continue
                if ch < "A" or ch > "Z":
                    return -1
                iResult = iResult * 26 + (ord(ch) - ord("A") + 1)
            return iResult - 1

        def is_cell_reference_project_list(pszText: str) -> bool:
            pszTextUpper: str = pszText.upper()
            if "!" in pszTextUpper:
                pszTextUpper = pszTextUpper.split("!", 1)[1]
            iIndex: int = 0
            while iIndex < len(pszTextUpper) and pszTextUpper[iIndex].isalpha():
                iIndex += 1
            if iIndex == 0:
                return False
            pszRowPart: str = pszTextUpper[iIndex:]
            return pszRowPart.isdigit()

        def parse_cell_reference_project_list(
            pszReference: str,
        ) -> Tuple[str | None, int, int]:
            if "!" in pszReference:
                pszSheetName, pszCell = pszReference.split("!", 1)
            else:
                pszSheetName = None
                pszCell = pszReference
            pszColumnPart: str = ""
            pszRowPart: str = ""
            for ch in pszCell:
                if ch.isalpha() or ch == "$":
                    pszColumnPart += ch
                else:
                    pszRowPart += ch
            pszColumnPart = pszColumnPart.replace("$", "")
            pszRowPart = pszRowPart.replace("$", "")
            iRow: int = int(pszRowPart) - 1 if pszRowPart.isdigit() else -1
            iColumn: int = column_label_to_index_project_list(pszColumnPart) if len(pszColumnPart) > 0 else -1
            return pszSheetName, iRow, iColumn

        def parse_expression_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objValue, iPos = parse_comparison_project_list(objTokens, iStart)
            return objValue, iPos

        def parse_comparison_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_term_project_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "cmp":
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_term_project_list(objTokens, iPos + 1)
                objLeft = ("cmp", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_term_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_factor_project_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["+", "-", "&"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_factor_project_list(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_factor_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_unary_project_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["*", "/"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_unary_project_list(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_unary_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            if iStart < len(objTokens) and objTokens[iStart][0] == "op" and objTokens[iStart][1] in ["+", "-"]:
                pszOp: str = objTokens[iStart][1]
                objValue, iPos = parse_unary_project_list(objTokens, iStart + 1)
                return ("unary", pszOp, objValue), iPos
            return parse_primary_project_list(objTokens, iStart)

        def parse_primary_project_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objTokenType, pszToken = objTokens[iStart]
            if objTokenType == "number":
                return ("number", pszToken), iStart + 1
            if objTokenType == "string":
                return ("string", pszToken), iStart + 1
            if objTokenType == "ident":
                if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", "("):
                    pszFuncName: str = pszToken.upper()
                    objArgs: List[Any] = []
                    iPos: int = iStart + 2
                    if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                        iPos += 1
                    else:
                        while iPos < len(objTokens):
                            objArg, iPos = parse_expression_project_list(objTokens, iPos)
                            objArgs.append(objArg)
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ","):
                                iPos += 1
                                continue
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                                iPos += 1
                                break
                    return ("func", pszFuncName, objArgs), iPos
                if is_cell_reference_project_list(pszToken):
                    if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", ":"):
                        pszSecond: str = objTokens[iStart + 2][1] if (iStart + 2) < len(objTokens) else ""
                        return ("range", pszToken, pszSecond), iStart + 3
                    return ("cell", pszToken), iStart + 1
                return ("string", pszToken), iStart + 1
            if objTokenType == "symbol" and pszToken == "(":
                objValue, iPos = parse_expression_project_list(objTokens, iStart + 1)
                if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                    iPos += 1
                return objValue, iPos
            return ("string", pszToken), iStart + 1

        def evaluate_range_project_list(
            pszSheet: str | None,
            pszStartCell: str,
            pszEndCell: str,
            pszCurrentSheet: str,
        ) -> List[Any]:
            pszStartSheet, iStartRow, iStartColumn = parse_cell_reference_project_list(
                (pszSheet + "!" if pszSheet else "") + pszStartCell,
            )
            pszEndSheet, iEndRow, iEndColumn = parse_cell_reference_project_list(
                (pszSheet + "!" if pszSheet else "") + pszEndCell,
            )
            pszTargetSheet: str | None = pszStartSheet or pszEndSheet or pszCurrentSheet
            objValues: List[Any] = []
            for iRow in range(min(iStartRow, iEndRow), max(iStartRow, iEndRow) + 1):
                for iColumn in range(min(iStartColumn, iEndColumn), max(iStartColumn, iEndColumn) + 1):
                    objValues.append(
                        evaluate_cell_by_sheet_project_list(
                            pszTargetSheet,
                            iRow,
                            iColumn,
                            pszCurrentSheet,
                        ),
                    )
            return objValues

        def evaluate_node_project_list(
            objNode: Any,
            iCurrentRow: int,
            iCurrentColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            if objNode[0] == "number":
                try:
                    if "." in objNode[1]:
                        return float(objNode[1])
                    return int(objNode[1])
                except ValueError:
                    return 0
            if objNode[0] == "string":
                return str(objNode[1])
            if objNode[0] == "unary":
                pszOp: str = objNode[1]
                objValue: Any = evaluate_node_project_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                try:
                    fNumber: float = float(objValue)
                    return -fNumber if pszOp == "-" else fNumber
                except Exception:
                    return objValue
            if objNode[0] == "op":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node_project_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                objRight: Any = evaluate_node_project_list(
                    objNode[3],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                if pszOp == "&":
                    return str(objLeft) + str(objRight)
                try:
                    fLeft: float = float(objLeft)
                    fRight: float = float(objRight)
                    if pszOp == "+":
                        return fLeft + fRight
                    if pszOp == "-":
                        return fLeft - fRight
                    if pszOp == "*":
                        return fLeft * fRight
                    if pszOp == "/":
                        return fLeft / fRight if fRight != 0 else 0
                except Exception:
                    if pszOp == "+":
                        return str(objLeft) + str(objRight)
                return ""
            if objNode[0] == "cmp":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node_project_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                objRight: Any = evaluate_node_project_list(
                    objNode[3],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                if pszOp == "=":
                    return objLeft == objRight
                if pszOp == "<":
                    return objLeft < objRight
                if pszOp == ">":
                    return objLeft > objRight
                if pszOp == "<=":
                    return objLeft <= objRight
                if pszOp == ">=":
                    return objLeft >= objRight
                if pszOp == "<>":
                    return objLeft != objRight
            if objNode[0] == "cell":
                pszSheet, iRow, iColumn = parse_cell_reference_project_list(objNode[1])
                pszTargetSheet: str = pszSheet or pszCurrentSheet
                return evaluate_cell_by_sheet_project_list(
                    pszTargetSheet,
                    iRow,
                    iColumn,
                    pszCurrentSheet,
                )
            if objNode[0] == "range":
                pszSheetStart, _, _ = parse_cell_reference_project_list(objNode[1])
                pszSheetEnd, _, _ = parse_cell_reference_project_list(objNode[2])
                pszSheetNameRange: str | None = pszSheetStart or pszSheetEnd or pszCurrentSheet
                pszStartCell: str = objNode[1].split("!", 1)[-1]
                pszEndCell: str = objNode[2].split("!", 1)[-1]
                return evaluate_range_project_list(
                    pszSheetNameRange,
                    pszStartCell,
                    pszEndCell,
                    pszCurrentSheet,
                )
            if objNode[0] == "func":
                pszFuncName: str = objNode[1]
                objArgsNodes: List[Any] = objNode[2]
                objArgsValues: List[Any] = [
                    evaluate_node_project_list(
                        objArg,
                        iCurrentRow,
                        iCurrentColumn,
                        pszCurrentSheet,
                    )
                    for objArg in objArgsNodes
                ]
                if pszFuncName == "IF":
                    bCondition: bool = bool(objArgsValues[0]) if len(objArgsValues) > 0 else False
                    if bCondition:
                        return objArgsValues[1] if len(objArgsValues) > 1 else ""
                    return objArgsValues[2] if len(objArgsValues) > 2 else ""
                if pszFuncName == "SUM":
                    fTotal: float = 0.0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                try:
                                    fTotal += float(objItem)
                                except Exception:
                                    continue
                        else:
                            try:
                                fTotal += float(objArg)
                            except Exception:
                                continue
                    return fTotal
                if pszFuncName == "COUNT":
                    iCount: int = 0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                if objItem not in ["", None]:
                                    iCount += 1
                        else:
                            if objArg not in ["", None]:
                                iCount += 1
                    return iCount
                if pszFuncName == "TEXT":
                    return str(objArgsValues[0]) if len(objArgsValues) > 0 else ""
                if pszFuncName == "CONCAT":
                    return "".join(str(objArg) for objArg in objArgsValues)
            return ""

        def evaluate_formula_project_list(
            pszFormula: str,
            iCurrentRow: int,
            iCurrentColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            objTokens = tokenize_formula_project_list(pszFormula)
            objAst, _ = parse_expression_project_list(objTokens, 0)
            return evaluate_node_project_list(
                objAst,
                iCurrentRow,
                iCurrentColumn,
                pszCurrentSheet,
            )

        def evaluate_cell_by_sheet_project_list(
            pszSheetName: str,
            iRow: int,
            iColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            if pszSheetName == "ローデータ":
                if (iRow < 0) or (iColumn < 0):
                    return ""
                if iRow >= objRawDataSheetForProjectList.shape[0] or iColumn >= objRawDataSheetForProjectList.shape[1]:
                    return ""
                objValue = objRawDataSheetForProjectList.iat[iRow, iColumn]
                return "" if pd.isna(objValue) else objValue
            return evaluate_cell_project_list(pszSheetName, iRow, iColumn)

        def evaluate_cell_project_list(
            pszSheetName: str,
            iRow: int,
            iColumn: int,
        ) -> Any:
            objKey: Tuple[str, int, int] = (pszSheetName, iRow, iColumn)
            if objKey in objEvaluatedCellsProjectList:
                return objEvaluatedCellsProjectList[objKey]
            if objKey in objEvaluatingCellsProjectList:
                raise ProjectListCircularReferenceError(
                    "Error: circular reference detected at {0}!R{1}C{2}".format(
                        "プロジェクトリスト",
                        iRow + 1,
                        iColumn + 1,
                    ),
                )
            if (iRow < 0) or (iColumn < 0):
                return ""
            if iRow >= objProjectListFormulaSheet.shape[0] or iColumn >= objProjectListFormulaSheet.shape[1]:
                return ""
            objValue = objProjectListFormulaSheet.iat[iRow, iColumn]
            pszValueStr: str = "" if pd.isna(objValue) else str(objValue)
            if len(pszValueStr) > 0 and pszValueStr.startswith("="):
                pszFormula: str = pszValueStr[1:]
                objEvaluatingCellsProjectList.add(objKey)
                try:
                    objResult: Any = evaluate_formula_project_list(
                        pszFormula,
                        iRow,
                        iColumn,
                        pszSheetName,
                    )
                except ProjectListCircularReferenceError as objCircularError:
                    with open(
                        pszOutputTsvPath.replace(".tsv", "_error.tsv"),
                        "w",
                        encoding="utf-8",
                    ) as objFile:
                        objFile.write(str(objCircularError))
                    raise
                except Exception as objException:
                    with open(
                        pszOutputTsvPath.replace(".tsv", "_error.tsv"),
                        "w",
                        encoding="utf-8",
                    ) as objFile:
                        objFile.write(
                            "Error: formula evaluation failed at {0}!R{1}C{2}: {3}".format(
                                "プロジェクトリスト",
                                iRow + 1,
                                iColumn + 1,
                                objException,
                            )
                        )
                    raise
                finally:
                    objEvaluatingCellsProjectList.discard(objKey)
                objEvaluatedCellsProjectList[objKey] = objResult
                return objResult
            objEvaluatedCellsProjectList[objKey] = pszValueStr
            return pszValueStr

        for iRowIndex in range(objProjectListFormulaSheet.shape[0]):
            for iColumnIndex in range(objProjectListFormulaSheet.shape[1]):
                objEvaluatedValueProjectList: Any = evaluate_cell_project_list(
                    "プロジェクトリスト",
                    iRowIndex,
                    iColumnIndex,
                )
                objOutputProjectList.iat[iRowIndex, iColumnIndex] = objEvaluatedValueProjectList

        objOutputProjectList.to_csv(
            pszOutputTsvPath,
            sep="\t",
            header=False,
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )

        objCheckProjectList: pd.DataFrame = pd.read_csv(
            pszOutputTsvPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            header=None,
            engine="python",
        )
        bHasFormulaCell: bool = False
        for iRowIndex in range(1, objCheckProjectList.shape[0]):
            for iColumnIndex in range(objCheckProjectList.shape[1]):
                objValue = objCheckProjectList.iat[iRowIndex, iColumnIndex]
                if isinstance(objValue, str) and objValue.startswith("="):
                    bHasFormulaCell = True
                    break
            if bHasFormulaCell:
                break
        if bHasFormulaCell:
            if os.path.isfile(pszOutputTsvPath):
                os.remove(pszOutputTsvPath)
            with open(
                pszOutputTsvPath.replace(".tsv", "_error.tsv"),
                "w",
                encoding="utf-8",
            ) as objFile:
                objFile.write(
                    "Error: Project_List.tsv still contains formula cells starting with '='."
                )
    except Exception as objException:
        if os.path.isfile(pszOutputTsvPath):
            os.remove(pszOutputTsvPath)
        with open(
            pszOutputTsvPath.replace(".tsv", "_error.tsv"),
            "w",
            encoding="utf-8",
        ) as objFile:
            objFile.write(
                "Error: unexpected exception. Detail = {0}".format(
                    objException,
                )
            )


# ///////////////////////////////////////////////////////////////
#
# main
#
# ///////////////////////////////////////////////////////////////
def main() -> int:
    objParser: argparse.ArgumentParser = argparse.ArgumentParser()
    objParser.add_argument(
        "pszInputManhourCsvPath",
        help="Input Jobcan manhour CSV file path",
    )
    objArgs: argparse.Namespace = objParser.parse_args()

    pszInputManhourCsvPath: str = objArgs.pszInputManhourCsvPath
    objInputPath: Path = Path(pszInputManhourCsvPath)

    objCandidatePaths: List[Path] = [objInputPath]

    objScriptDirectoryPath: Path = Path(__file__).resolve().parent
    objCandidatePaths.append(objScriptDirectoryPath / pszInputManhourCsvPath)

    objInputDirectoryPath: Path = Path.cwd() / "input"
    objCandidatePaths.append(objInputDirectoryPath / pszInputManhourCsvPath)

    if objInputPath.suffix.lower() == ".tsv":
        pszCsvFileName: str = objInputPath.with_suffix(".csv").name
        objCandidatePaths.append(objInputPath.with_suffix(".csv"))
        objCandidatePaths.append(objScriptDirectoryPath / pszCsvFileName)
        objCandidatePaths.append(objInputDirectoryPath / pszCsvFileName)

    objExistingPaths: List[Path] = [objPath for objPath in objCandidatePaths if objPath.exists()]
    if len(objExistingPaths) > 0:
        objInputPath = objExistingPaths[0]

    if not objInputPath.exists():
        pszErrorTextFilePath: str = str(Path.cwd() / "make_manhour_to_sheet8_error.txt")
        write_error_text_utf8(
            pszErrorTextFilePath,
            f"Error: input file not found: {pszInputManhourCsvPath}\n"
            f"CurrentDirectory: {str(Path.cwd())}\n",
        )
        raise FileNotFoundError(f"Input file not found: {pszInputManhourCsvPath}")

    objBaseDirectoryPath: Path = objInputPath.resolve().parent

    # (1) CSV → TSV (H:MM:SS 化)
    objModuleCsvToTsv: Dict[str, Any] = create_module_from_source(
        "csv_to_tsv_h_mm_ss",
        pszSource_csv_to_tsv_h_mm_ss_py,
    )
    pszStep1DefaultTsvPath: str = objModuleCsvToTsv["convert_csv_to_tsv_file"](
        str(objInputPath),
    )
    iFileYear: int
    iFileMonth: int
    iFileYear, iFileMonth = get_target_year_month_from_filename(str(objInputPath))
    pszStep1TsvPath: str = str(
        objBaseDirectoryPath / f"工数_{iFileYear}年{iFileMonth:02d}月.tsv"
    )
    if pszStep1DefaultTsvPath != pszStep1TsvPath:
        os.replace(pszStep1DefaultTsvPath, pszStep1TsvPath)

    # (2) 未入力行除去
    objModuleRemoveUninput: Dict[str, Any] = create_module_from_source(
        "manhour_remove_uninput_rows",
        pszSource_manhour_remove_uninput_rows_py,
    )
    objModuleRemoveUninput["make_removed_uninput_tsv_from_manhour_tsv"](
        pszStep1TsvPath,
    )
    pszStep2TsvPath: str = objModuleRemoveUninput["build_output_file_full_path"](
        pszStep1TsvPath,
    )

    # (3) スタッフコード順ソート
    objModuleSortByStaffCode: Dict[str, Any] = create_module_from_source(
        "sort_manhour_by_staff_code",
        pszSource_sort_manhour_by_staff_code_py,
    )
    objModuleSortByStaffCode["make_sorted_staff_code_tsv_from_manhour_tsv"](
        pszStep2TsvPath,
    )
    pszStep3TsvPath: str = objModuleSortByStaffCode["build_output_file_full_path"](
        pszStep2TsvPath,
    )

    # (4) 日付正規化 → 工数_yyyy年mm月_step04_yyyy_mm_dd.tsv
    objModuleConvertDate: Dict[str, Any] = create_module_from_source(
        "convert_yyyy_mm_dd",
        pszSource_convert_yyyy_mm_dd_py,
    )
    pszSheet4TsvPath: str = str(
        objBaseDirectoryPath / f"工数_{iFileYear}年{iFileMonth:02d}月_step04_yyyy_mm_dd.tsv"
    )
    objModuleConvertDate["make_sheet4_tsv_from_input_tsv"](
        pszStep3TsvPath,
        pszSheet4TsvPath,
    )

    # Sheet4.tsv からスタッフコード一覧を作成
    objModuleUniqueStaffCodeList: Dict[str, Any] = create_module_from_source(
        "make_unique_staff_code_list",
        pszSource_make_unique_staff_code_list_py,
    )
    objModuleUniqueStaffCodeList["make_unique_staff_code_tsv_from_sheet1_tsv"](
        pszSheet4TsvPath,
    )
    pszSheet4UniqueStaffCodeTsvPath: str = objModuleUniqueStaffCodeList[
        "build_output_file_full_path"
    ](pszSheet4TsvPath)

    # (5) Sheet4_staff_code_range.tsv
    objModuleMakeRange: Dict[str, Any] = create_module_from_source(
        "make_staff_code_range",
        pszSource_make_staff_code_range_py,
    )
    objModuleMakeRange["make_staff_code_range_tsv_from_sheet1_tsv"](
        pszSheet4TsvPath,
    )
    pszSheet4StaffCodeRangeTsvPath: str = objModuleMakeRange[
        "build_output_file_full_path"
    ](pszSheet4TsvPath)

    # (6) 工数_yyyy年mm月_step05_スタッフ別担当プロジェクト.tsv
    objModuleMakeSheet6: Dict[str, Any] = create_module_from_source(
        "make_sheet6_from_sheet4",
        pszSource_make_sheet6_from_sheet4_py,
    )
    pszSheet6DefaultTsvPath: str = objModuleMakeSheet6["build_output_file_full_path"](
        pszSheet4TsvPath,
    )
    pszSheet6TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step05_スタッフ別担当プロジェクト.tsv"
    )
    objModuleMakeSheet6["make_sheet6_from_sheet4"](
        pszSheet4TsvPath,
        pszSheet4StaffCodeRangeTsvPath,
    )
    if pszSheet6DefaultTsvPath != pszSheet6TsvPath:
        os.replace(pszSheet6DefaultTsvPath, pszSheet6TsvPath)

    # (7) 工数_yyyy年mm月_step06_プロジェクト_タスク_工数.tsv
    #     工数_yyyy年mm月_step06_旧版_スタッフ別_プロジェクト_タスク_工数.tsv
    #     工数_yyyy年mm月_step06_旧版_氏名_スタッフコード.tsv
    objModuleMakeSheet789: Dict[str, Any] = create_module_from_source(
        "make_sheet789_from_sheet4",
        pszSource_make_sheet789_from_sheet4_py,
    )
    pszSheet7DefaultTsvPath: str = objModuleMakeSheet789[
        "build_output_file_full_path_for_sheet7"
    ](pszSheet4TsvPath)
    pszSheet8DefaultTsvPath: str = objModuleMakeSheet789[
        "build_output_file_full_path_for_sheet8"
    ](pszSheet4TsvPath)
    pszSheet9DefaultTsvPath: str = objModuleMakeSheet789[
        "build_output_file_full_path_for_sheet9"
    ](pszSheet4TsvPath)
    pszSheet7TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step06_プロジェクト_タスク_工数.tsv"
    )
    pszSheet8TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step06_旧版_スタッフ別_プロジェクト_タスク_工数.tsv"
    )
    pszSheet9TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step06_旧版_氏名_スタッフコード.tsv"
    )
    objModuleMakeSheet789["make_sheet789_from_sheet4"](
        pszSheet4TsvPath,
        pszSheet4StaffCodeRangeTsvPath,
        pszSheet6TsvPath,
    )
    if pszSheet7DefaultTsvPath != pszSheet7TsvPath:
        os.replace(pszSheet7DefaultTsvPath, pszSheet7TsvPath)
    if pszSheet8DefaultTsvPath != pszSheet8TsvPath:
        os.replace(pszSheet8DefaultTsvPath, pszSheet8TsvPath)
    if pszSheet9DefaultTsvPath != pszSheet9TsvPath:
        os.replace(pszSheet9DefaultTsvPath, pszSheet9TsvPath)

    # (8) 工数_yyyy年mm月_step07_計算前_プロジェクト_工数.tsv
    #     工数_yyyy年mm月_step08_合計_プロジェクト_工数.tsv
    #     工数_yyyy年mm月_step09_昇順_合計_プロジェクト_工数.tsv
    pszSheet10TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step07_計算前_プロジェクト_工数.tsv"
    )
    pszSheet11TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step08_合計_プロジェクト_工数.tsv"
    )
    pszSheet12TsvPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step09_昇順_合計_プロジェクト_工数.tsv"
    )

    def is_blank_sheet10(value: str | None) -> bool:
        if value is None:
            return True
        if value == "":
            return True
        if value.strip() == "":
            return True
        if value.lower() == "nan":
            return True
        return False

    def normalize_project_name_sheet10(pszSource: str) -> str:
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

        if pszSource.startswith("【"):
            iBracketEndIndex: int = pszSource.find("】")
            if iBracketEndIndex != -1:
                pszAfterBracket: str = pszSource[iBracketEndIndex + 1:]
                objMatch = re.search(r"(P\d{5}|[A-OQ-Z]\d{3})", pszAfterBracket)
                if objMatch is not None:
                    pszCode: str = objMatch.group(1)
                    pszBeforeCode: str = pszAfterBracket[:objMatch.start()]
                    pszAfterCode: str = pszAfterBracket[objMatch.end():]
                    if pszAfterCode.startswith(" ") or pszAfterCode.startswith("　"):
                        pszAfterCode = pszAfterCode[1:]
                    pszRest: str = pszSource[: iBracketEndIndex + 1] + pszBeforeCode + pszAfterCode
                    return pszCode + "_" + pszRest

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

    def preprocess_line_content_sheet10(pszLineContent: str) -> str:
        if pszLineContent.startswith('"'):
            iSecondQuoteIndex: int = pszLineContent.find('"', 1)
            if iSecondQuoteIndex != -1:
                if iSecondQuoteIndex + 1 < len(pszLineContent) and pszLineContent[iSecondQuoteIndex + 1] == "\t":
                    pszQuotedContent: str = pszLineContent[1:iSecondQuoteIndex].replace("\t", "_")
                    pszLineContent = pszQuotedContent + pszLineContent[iSecondQuoteIndex + 1:]
        pszLineContent = re.sub(
            r'^"([^"]*)\t([^"]*)"([^\r\n]*)',
            r"\1_\2\3",
            pszLineContent,
        )
        pszLineContent = re.sub(r"(P\d\d\d\d\d)(?![ _\t　【])", r"\1_", pszLineContent)
        pszLineContent = re.sub(r"([A-OQ-Z]\d\d\d)(?![ _\t　【])", r"\1_", pszLineContent)
        pszLineContent = re.sub(r"^(J\d\d\d) +", r"\1_", pszLineContent)
        pszLineContent = re.sub(r"([A-OQ-Z]\d\d\d)[ 　]+", r"\1_", pszLineContent)
        pszLineContent = re.sub(r"(P\d\d\d\d\d)[ 　]+", r"\1_", pszLineContent)
        pszLineContent = re.sub(r"\t[0-9]+\t", "\t", pszLineContent)
        return pszLineContent

    def parse_manhour_to_seconds_sheet11(pszManhour: str) -> int:
        objMatch = re.match(r"^(\d+):([0-5]\d):([0-5]\d)$", pszManhour)
        if not objMatch:
            raise ValueError(f"Invalid manhour format: {pszManhour}")
        iHours: int = int(objMatch.group(1))
        iMinutes: int = int(objMatch.group(2))
        iSeconds: int = int(objMatch.group(3))
        return iHours * 3600 + iMinutes * 60 + iSeconds

    def format_seconds_to_manhour_sheet11(iTotalSeconds: int) -> str:
        if iTotalSeconds < 0:
            raise ValueError("Total seconds must not be negative.")
        iHours: int = iTotalSeconds // 3600
        iMinutes: int = (iTotalSeconds % 3600) // 60
        iSeconds: int = iTotalSeconds % 60
        return f"{iHours}:{iMinutes:02d}:{iSeconds:02d}"

    def extract_project_prefix_sheet12(pszProjectName: str) -> str:
        iUnderscoreIndex: int = pszProjectName.find("_")
        if iUnderscoreIndex == -1:
            return pszProjectName
        return pszProjectName[:iUnderscoreIndex]

    with open(pszSheet7TsvPath, "r", encoding="utf-8") as objSheet7File:
        objSheet7Lines: List[str] = objSheet7File.readlines()

    objSheet10Rows: List[Tuple[str, str]] = []
    with open(pszSheet10TsvPath, "w", encoding="utf-8") as objSheet10File:
        for pszLine in objSheet7Lines:
            pszLineContent: str = pszLine.rstrip("\n")
            if pszLineContent == "":
                objSheet10File.write("\t\n")
                objSheet10Rows.append(("", ""))
                continue
            pszLineContent = preprocess_line_content_sheet10(pszLineContent)
            objColumns: List[str] = pszLineContent.split("\t")
            pszProjectName: str = ""
            pszManhour: str = ""
            if len(objColumns) > 0:
                pszProjectName = objColumns[0]
            if len(objColumns) > 2:
                pszManhour = objColumns[2]
            elif len(objColumns) > 1:
                pszManhour = objColumns[1]
            if is_blank_sheet10(pszProjectName):
                pszNormalizedName: str = ""
            else:
                pszNormalizedName = normalize_project_name_sheet10(pszProjectName)
            objSheet10File.write(pszNormalizedName + "\t" + pszManhour + "\n")
            objSheet10Rows.append((pszNormalizedName, pszManhour))

    objAggregatedSeconds: Dict[str, int] = {}
    objAggregatedOrder: List[str] = []
    for pszProjectName, pszManhour in objSheet10Rows:
        if pszProjectName == "" and pszManhour == "":
            continue
        iSeconds = parse_manhour_to_seconds_sheet11(pszManhour)
        if pszProjectName not in objAggregatedSeconds:
            objAggregatedSeconds[pszProjectName] = 0
            objAggregatedOrder.append(pszProjectName)
        objAggregatedSeconds[pszProjectName] += iSeconds

    objSheet11Rows: List[Tuple[str, str]] = []
    with open(pszSheet11TsvPath, "w", encoding="utf-8") as objSheet11File:
        for pszProjectName in objAggregatedOrder:
            pszTotalManhour: str = format_seconds_to_manhour_sheet11(
                objAggregatedSeconds[pszProjectName],
            )
            objSheet11File.write(pszProjectName + "\t" + pszTotalManhour + "\n")
            objSheet11Rows.append((pszProjectName, pszTotalManhour))

    objIndexedSheet11Rows: List[Tuple[int, Tuple[str, str]]] = list(enumerate(objSheet11Rows))
    objIndexedSheet11Rows.sort(
        key=lambda objItem: (
            extract_project_prefix_sheet12(objItem[1][0]),
            objItem[0],
        ),
    )

    with open(pszSheet12TsvPath, "w", encoding="utf-8") as objSheet12File:
        for _, objRow in objIndexedSheet11Rows:
            objSheet12File.write(objRow[0] + "\t" + objRow[1] + "\n")

    pszStep10OutputPath: str = str(
        objBaseDirectoryPath
        / f"工数_{iFileYear}年{iFileMonth:02d}月_step10_各プロジェクトの工数.tsv"
    )
    with open(pszStep10OutputPath, "w", encoding="utf-8") as objStep10File:
        for _, (pszProjectName, pszTotalManhour) in objIndexedSheet11Rows:
            if str(pszProjectName).startswith(("A", "H")):
                continue
            objStep10File.write(pszProjectName + "\t" + pszTotalManhour + "\n")

    # Staff_List.tsv を作成する処理ブロック
    try:
        def write_error_tsv_staff_list(
            pszOutputFileFullPath: str,
            pszErrorMessage: str,
        ) -> None:
            pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
            if len(pszDirectory) > 0:
                os.makedirs(pszDirectory, exist_ok=True)

            with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
                objFile.write(pszErrorMessage)
                if not pszErrorMessage.endswith("\n"):
                    objFile.write("\n")

        def load_tsv_for_staff_list(
            pszFileFullPath: str,
        ) -> pd.DataFrame | None:
            if not os.path.isfile(pszFileFullPath):
                write_error_tsv_staff_list(
                    pszFileFullPath.replace(".tsv", "_error.tsv"),
                    "Error: input TSV file not found. Path = {0}".format(pszFileFullPath),
                )
                return None

            try:
                return pd.read_csv(
                    pszFileFullPath,
                    sep="\t",
                    dtype=str,
                    encoding="utf-8",
                    engine="python",
                )
            except Exception:
                write_error_tsv_staff_list(
                    pszFileFullPath.replace(".tsv", "_error.tsv"),
                    "Error: input TSV file not found. Path = {0}".format(pszFileFullPath),
                )
                return None

        def simplify_iferror_in_formula_cell(
            pszCellValue: str,
        ) -> str:
            if not pszCellValue.startswith("="):
                return pszCellValue

            objPattern: re.Pattern[str] = re.compile(
                r"^=IFERROR\(\s*IFERROR\(\s*(.+?)\s*,\s*(\"\"|''?)\s*\)\s*,\s*(.+?)\s*\)\s*$",
                re.IGNORECASE,
            )
            objMatch: re.Match[str] | None = objPattern.match(pszCellValue)
            if objMatch is None:
                return pszCellValue

            pszInnerExpression: str = objMatch.group(1)
            pszInnerErrorValue: str = objMatch.group(2)
            return f"=IFERROR({pszInnerExpression},{pszInnerErrorValue})"

        def normalize_staff_list_formula_dataframe(
            objDataFrame: pd.DataFrame,
        ) -> pd.DataFrame:
            objNormalized: pd.DataFrame = objDataFrame.copy()
            for iRowIndex in range(len(objNormalized.index)):
                for iColIndex in range(len(objNormalized.columns)):
                    objRawValue: object = objNormalized.iat[iRowIndex, iColIndex]
                    pszCellValue: str = "" if pd.isna(objRawValue) else str(objRawValue)
                    objNormalized.iat[iRowIndex, iColIndex] = simplify_iferror_in_formula_cell(
                        pszCellValue,
                    )
            return objNormalized

        def restore_blank_column_names(objDataFrame: pd.DataFrame) -> pd.DataFrame:
            objNewColumns: List[str] = []
            for objColumn in objDataFrame.columns:
                pszColumn: str = str(objColumn)
                if re.match(r"^Unnamed:\s*\d+$", pszColumn):
                    objNewColumns.append("")
                else:
                    objNewColumns.append(pszColumn)
            objDataFrame.columns = objNewColumns
            return objDataFrame

        pszStaffListFormulaTsvPath: str = str(objBaseDirectoryPath / "Staff_List_Formula.tsv")
        pszStaffListTsvPath: str = str(objBaseDirectoryPath / "Staff_List.tsv")
        pszRawDataTsvPath: str = str(objBaseDirectoryPath / "Raw_Data.tsv")

        objStaffListFormulaSheet: pd.DataFrame | None = load_tsv_for_staff_list(
            pszStaffListFormulaTsvPath,
        )
        if objStaffListFormulaSheet is None:
            return 1

        objStaffListFormulaSheet = restore_blank_column_names(
            normalize_staff_list_formula_dataframe(
                objStaffListFormulaSheet,
            ),
        )

        objRawDataSheetForStaffList: pd.DataFrame | None = load_tsv_for_staff_list(
            pszRawDataTsvPath,
        )
        if objRawDataSheetForStaffList is None:
            return 1

        objOutputStaffList: pd.DataFrame = pd.DataFrame(
            "",
            index=objStaffListFormulaSheet.index,
            columns=objStaffListFormulaSheet.columns,
        )

        objEvaluatedCells: Dict[Tuple[str, int, int], Any] = {}
        objEvaluatingCells: set[Tuple[str, int, int]] = set()

        class StaffListCircularReferenceError(RuntimeError):
            pass

        def tokenize_formula_staff_list(
            pszFormula: str,
        ) -> List[Tuple[str, str]]:
            objTokens: List[Tuple[str, str]] = []
            i: int = 0
            while i < len(pszFormula):
                ch: str = pszFormula[i]
                if ch.isspace():
                    i += 1
                    continue
                if ch in ":,()":
                    objTokens.append(("symbol", ch))
                    i += 1
                    continue
                if ch in "+-*/&":
                    objTokens.append(("op", ch))
                    i += 1
                    continue
                if ch in "<>=" and i + 1 < len(pszFormula):
                    ch2: str = pszFormula[i + 1]
                    if (ch == "<" and ch2 in ["=", ">"]) or (ch == ">" and ch2 == "="):
                        objTokens.append(("cmp", ch + ch2))
                        i += 2
                        continue
                if ch in "<>":
                    objTokens.append(("cmp", ch))
                    i += 1
                    continue
                if ch == "=":
                    objTokens.append(("cmp", ch))
                    i += 1
                    continue
                if ch == "\"":
                    i += 1
                    pszLiteral: str = ""
                    while i < len(pszFormula):
                        if pszFormula[i] == "\"":
                            if i + 1 < len(pszFormula) and pszFormula[i + 1] == "\"":
                                pszLiteral += "\""
                                i += 2
                                continue
                            i += 1
                            break
                        pszLiteral += pszFormula[i]
                        i += 1
                    objTokens.append(("string", pszLiteral))
                    continue
                if ch.isdigit() or ch == ".":
                    pszNumber: str = ""
                    while i < len(pszFormula) and (pszFormula[i].isdigit() or pszFormula[i] == "."):
                        pszNumber += pszFormula[i]
                        i += 1
                    objTokens.append(("number", pszNumber))
                    continue
                pszIdentifier: str = ""
                while i < len(pszFormula):
                    ch2 = pszFormula[i]
                    if ch2.isspace() or ch2 in "+-*/&,()":
                        break
                    if ch2 in ["<", ">", "="]:
                        break
                    pszIdentifier += ch2
                    i += 1
                objTokens.append(("ident", pszIdentifier))
            return objTokens

        def is_cell_reference_staff_list(
            pszText: str,
        ) -> bool:
            pszTemp: str = pszText.replace("$", "")
            if "!" in pszTemp:
                pszTemp = pszTemp.split("!", 1)[1]
            i: int = 0
            while i < len(pszTemp) and pszTemp[i].isalpha():
                i += 1
            if i == 0:
                return False
            pszRowPart: str = pszTemp[i:]
            if len(pszRowPart) == 0:
                return False
            return pszRowPart.isdigit()

        def column_label_to_index_staff_list(
            pszColumnLabel: str,
        ) -> int:
            iIndex: int = 0
            for ch in pszColumnLabel:
                iIndex = iIndex * 26 + (ord(ch.upper()) - ord("A") + 1)
            return iIndex - 1

        def parse_cell_reference_staff_list(
            pszReference: str,
        ) -> Tuple[str | None, int, int]:
            if "!" in pszReference:
                pszSheetName, pszCell = pszReference.split("!", 1)
            else:
                pszSheetName, pszCell = None, pszReference

            pszColumnPart: str = ""
            pszRowPart: str = ""
            for ch in pszCell:
                if ch.isalpha() or ch == "$":
                    pszColumnPart += ch
                else:
                    pszRowPart += ch
            pszColumnPart = pszColumnPart.replace("$", "")
            pszRowPart = pszRowPart.replace("$", "")
            iRow: int = int(pszRowPart) - 1 if pszRowPart.isdigit() else -1
            iColumn: int = column_label_to_index_staff_list(pszColumnPart) if len(pszColumnPart) > 0 else -1
            return pszSheetName, iRow, iColumn

        def parse_expression_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objValue, iPos = parse_comparison_staff_list(objTokens, iStart)
            return objValue, iPos

        def parse_comparison_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_term_staff_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "cmp":
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_term_staff_list(objTokens, iPos + 1)
                objLeft = ("cmp", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_term_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_factor_staff_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["+", "-", "&"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_factor_staff_list(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_factor_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_unary_staff_list(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["*", "/"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_unary_staff_list(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_unary_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            if iStart < len(objTokens) and objTokens[iStart][0] == "op" and objTokens[iStart][1] in ["+", "-"]:
                pszOp: str = objTokens[iStart][1]
                objValue, iPos = parse_unary_staff_list(objTokens, iStart + 1)
                return ("unary", pszOp, objValue), iPos
            return parse_primary_staff_list(objTokens, iStart)

        def parse_primary_staff_list(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objTokenType, pszToken = objTokens[iStart]
            if objTokenType == "number":
                return ("number", pszToken), iStart + 1
            if objTokenType == "string":
                return ("string", pszToken), iStart + 1
            if objTokenType == "ident":
                if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", "("):
                    pszFuncName: str = pszToken.upper()
                    objArgs: List[Any] = []
                    iPos: int = iStart + 2
                    if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                        iPos += 1
                    else:
                        while iPos < len(objTokens):
                            objArg, iPos = parse_expression_staff_list(objTokens, iPos)
                            objArgs.append(objArg)
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ","):
                                iPos += 1
                                continue
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                                iPos += 1
                                break
                    return ("func", pszFuncName, objArgs), iPos
                if is_cell_reference_staff_list(pszToken):
                    if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", ":"):
                        pszSecond: str = objTokens[iStart + 2][1] if (iStart + 2) < len(objTokens) else ""
                        return ("range", pszToken, pszSecond), iStart + 3
                    return ("cell", pszToken), iStart + 1
                return ("string", pszToken), iStart + 1
            if objTokenType == "symbol" and pszToken == "(":
                objValue, iPos = parse_expression_staff_list(objTokens, iStart + 1)
                if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                    iPos += 1
                return objValue, iPos
            return ("string", pszToken), iStart + 1

        def evaluate_range_staff_list(
            pszSheet: str | None,
            pszStartCell: str,
            pszEndCell: str,
            pszCurrentSheet: str,
        ) -> List[Any]:
            pszStartSheet, iStartRow, iStartColumn = parse_cell_reference_staff_list(
                (pszSheet + "!" if pszSheet else "") + pszStartCell,
            )
            pszEndSheet, iEndRow, iEndColumn = parse_cell_reference_staff_list(
                (pszSheet + "!" if pszSheet else "") + pszEndCell,
            )
            pszTargetSheet: str | None = pszStartSheet or pszEndSheet or pszCurrentSheet
            objValues: List[Any] = []
            for iRow in range(min(iStartRow, iEndRow), max(iStartRow, iEndRow) + 1):
                for iColumn in range(min(iStartColumn, iEndColumn), max(iStartColumn, iEndColumn) + 1):
                    objValues.append(
                        evaluate_cell_by_sheet_staff_list(
                            pszTargetSheet,
                            iRow,
                            iColumn,
                            pszCurrentSheet,
                        ),
                    )
            return objValues

        def evaluate_node_staff_list(
            objNode: Any,
            iCurrentRow: int,
            iCurrentColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            if objNode[0] == "number":
                try:
                    if "." in objNode[1]:
                        return float(objNode[1])
                    return int(objNode[1])
                except ValueError:
                    return 0
            if objNode[0] == "string":
                return str(objNode[1])
            if objNode[0] == "unary":
                pszOp: str = objNode[1]
                objValue: Any = evaluate_node_staff_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                try:
                    fNumber: float = float(objValue)
                    return -fNumber if pszOp == "-" else fNumber
                except Exception:
                    return objValue
            if objNode[0] == "op":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node_staff_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                objRight: Any = evaluate_node_staff_list(
                    objNode[3],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                if pszOp == "&":
                    return str(objLeft) + str(objRight)
                try:
                    fLeft: float = float(objLeft)
                    fRight: float = float(objRight)
                    if pszOp == "+":
                        return fLeft + fRight
                    if pszOp == "-":
                        return fLeft - fRight
                    if pszOp == "*":
                        return fLeft * fRight
                    if pszOp == "/":
                        return fLeft / fRight if fRight != 0 else 0
                except Exception:
                    if pszOp == "+":
                        return str(objLeft) + str(objRight)
                return ""
            if objNode[0] == "cmp":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node_staff_list(
                    objNode[2],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                objRight: Any = evaluate_node_staff_list(
                    objNode[3],
                    iCurrentRow,
                    iCurrentColumn,
                    pszCurrentSheet,
                )
                if pszOp == "=":
                    return objLeft == objRight
                if pszOp == "<":
                    return objLeft < objRight
                if pszOp == ">":
                    return objLeft > objRight
                if pszOp == "<=":
                    return objLeft <= objRight
                if pszOp == ">=":
                    return objLeft >= objRight
                if pszOp == "<>":
                    return objLeft != objRight
            if objNode[0] == "cell":
                pszSheet, iRow, iColumn = parse_cell_reference_staff_list(objNode[1])
                pszTargetSheet: str = pszSheet or pszCurrentSheet
                return evaluate_cell_by_sheet_staff_list(
                    pszTargetSheet,
                    iRow,
                    iColumn,
                    pszCurrentSheet,
                )
            if objNode[0] == "range":
                pszSheetStart, _, _ = parse_cell_reference_staff_list(objNode[1])
                pszSheetEnd, _, _ = parse_cell_reference_staff_list(objNode[2])
                pszSheetNameRange: str | None = pszSheetStart or pszSheetEnd or pszCurrentSheet
                pszStartCell: str = objNode[1].split("!", 1)[-1]
                pszEndCell: str = objNode[2].split("!", 1)[-1]
                return evaluate_range_staff_list(
                    pszSheetNameRange,
                    pszStartCell,
                    pszEndCell,
                    pszCurrentSheet,
                )
            if objNode[0] == "func":
                pszFuncName: str = objNode[1]
                objArgsNodes: List[Any] = objNode[2]
                objArgsValues: List[Any] = [
                    evaluate_node_staff_list(
                        objArg,
                        iCurrentRow,
                        iCurrentColumn,
                        pszCurrentSheet,
                    )
                    for objArg in objArgsNodes
                ]
                if pszFuncName == "IF":
                    bCondition: bool = bool(objArgsValues[0]) if len(objArgsValues) > 0 else False
                    if bCondition:
                        return objArgsValues[1] if len(objArgsValues) > 1 else ""
                    return objArgsValues[2] if len(objArgsValues) > 2 else ""
                if pszFuncName == "SUM":
                    fTotal: float = 0.0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                try:
                                    fTotal += float(objItem)
                                except Exception:
                                    continue
                        else:
                            try:
                                fTotal += float(objArg)
                            except Exception:
                                continue
                    return fTotal
                if pszFuncName == "COUNT":
                    iCount: int = 0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                if objItem not in ["", None]:
                                    iCount += 1
                        else:
                            if objArg not in ["", None]:
                                iCount += 1
                    return iCount
                if pszFuncName == "TEXT":
                    return str(objArgsValues[0]) if len(objArgsValues) > 0 else ""
                if pszFuncName == "CONCAT":
                    return "".join(str(objArg) for objArg in objArgsValues)
            return ""

        def evaluate_formula_staff_list(
            pszFormula: str,
            iCurrentRow: int,
            iCurrentColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            objTokens = tokenize_formula_staff_list(pszFormula)
            objAst, _ = parse_expression_staff_list(objTokens, 0)
            return evaluate_node_staff_list(
                objAst,
                iCurrentRow,
                iCurrentColumn,
                pszCurrentSheet,
            )

        def evaluate_cell_by_sheet_staff_list(
            pszSheetName: str,
            iRow: int,
            iColumn: int,
            pszCurrentSheet: str,
        ) -> Any:
            if pszSheetName == "ローデータ":
                if (iRow < 0) or (iColumn < 0):
                    return ""
                if iRow >= objRawDataSheetForStaffList.shape[0] or iColumn >= objRawDataSheetForStaffList.shape[1]:
                    return ""
                objValue = objRawDataSheetForStaffList.iat[iRow, iColumn]
                return "" if pd.isna(objValue) else objValue
            return evaluate_cell_staff_list(pszSheetName, iRow, iColumn)

        def evaluate_cell_staff_list(
            pszSheetName: str,
            iRow: int,
            iColumn: int,
        ) -> Any:
            objKey: Tuple[str, int, int] = (pszSheetName, iRow, iColumn)
            if objKey in objEvaluatedCells:
                return objEvaluatedCells[objKey]
            if objKey in objEvaluatingCells:
                raise StaffListCircularReferenceError(
                    "Error: circular reference detected at {0}!R{1}C{2}".format(
                        "社員リスト",
                        iRow + 1,
                        iColumn + 1,
                    ),
                )
            if (iRow < 0) or (iColumn < 0):
                return ""
            if iRow >= objStaffListFormulaSheet.shape[0] or iColumn >= objStaffListFormulaSheet.shape[1]:
                return ""
            objValue = objStaffListFormulaSheet.iat[iRow, iColumn]
            pszValueStr: str = "" if pd.isna(objValue) else str(objValue)
            if len(pszValueStr) > 0 and pszValueStr.startswith("="):
                pszFormula: str = pszValueStr[1:]
                objEvaluatingCells.add(objKey)
                try:
                    objResult: Any = evaluate_formula_staff_list(
                        pszFormula,
                        iRow,
                        iColumn,
                        pszSheetName,
                    )
                except StaffListCircularReferenceError as objCircularError:
                    write_error_tsv_staff_list(
                        str(objBaseDirectoryPath / "Staff_List_error.tsv"),
                        str(objCircularError),
                    )
                    raise
                except Exception as objException:
                    write_error_tsv_staff_list(
                        str(objBaseDirectoryPath / "Staff_List_error.tsv"),
                        "Error: formula evaluation failed at {0}!R{1}C{2}: {3}".format(
                            "社員リスト",
                            iRow + 1,
                            iColumn + 1,
                            objException,
                        ),
                    )
                    raise
                finally:
                    objEvaluatingCells.discard(objKey)
                objEvaluatedCells[objKey] = objResult
                return objResult
            objEvaluatedCells[objKey] = pszValueStr
            return pszValueStr

        for iRowIndex in range(objStaffListFormulaSheet.shape[0]):
            for iColumnIndex in range(objStaffListFormulaSheet.shape[1]):
                try:
                    objEvaluatedValue: Any = evaluate_cell_staff_list(
                        "社員リスト",
                        iRowIndex,
                        iColumnIndex,
                    )
                except Exception:
                    return 1
                objOutputStaffList.iat[iRowIndex, iColumnIndex] = objEvaluatedValue

        objOutputStaffList.to_csv(
            pszStaffListTsvPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception:
        pass

    pszRawDataTsvPath: str = str(objBaseDirectoryPath / "Raw_Data.tsv")
    pszProjectListFormulaTsvPath: str = str(objBaseDirectoryPath / "Project_List_Formula.tsv")
    pszProjectListTsvPath: str = str(objBaseDirectoryPath / "Project_List.tsv")
    make_project_list_tsv_from_raw_data(
        pszRawDataTsvPath,
        pszProjectListFormulaTsvPath,
        pszProjectListTsvPath,
    )

    # With_Salary.tsv を作成する処理ブロック
    try:
        def write_error_tsv_with_salary(
            pszOutputFileFullPath: str,
            pszErrorMessage: str,
        ) -> None:
            pszDirectory: str = os.path.dirname(pszOutputFileFullPath)
            if len(pszDirectory) > 0:
                os.makedirs(pszDirectory, exist_ok=True)

            with open(pszOutputFileFullPath, "w", encoding="utf-8") as objFile:
                objFile.write(pszErrorMessage + "\n")

        def load_tsv_as_dataframe(
            pszFileFullPath: str,
            bHasHeader: bool,
        ) -> pd.DataFrame:
            if not os.path.isfile(pszFileFullPath):
                write_error_tsv_with_salary(
                    pszFileFullPath.replace(".tsv", "_error.tsv"),
                    "Error: input TSV file not found. Path = {0}".format(pszFileFullPath),
                )
                return pd.DataFrame()

            try:
                if bHasHeader:
                    return pd.read_csv(
                        pszFileFullPath,
                        sep="\t",
                        dtype=str,
                        encoding="utf-8",
                        header=0,
                        engine="python",
                    )
                return pd.read_csv(
                    pszFileFullPath,
                    sep="\t",
                    dtype=str,
                    encoding="utf-8",
                    header=None,
                    engine="python",
                )
            except FileNotFoundError:
                write_error_tsv_with_salary(
                    pszFileFullPath.replace(".tsv", "_error.tsv"),
                    "Error: input TSV file not found. Path = {0}".format(pszFileFullPath),
                )
                return pd.DataFrame()

        def column_letter_to_index(
            pszColumn: str,
        ) -> int:
            iResult: int = 0
            for ch in pszColumn:
                if ch == "$":
                    continue
                iResult = iResult * 26 + (ord(ch.upper()) - ord("A") + 1)
            return iResult - 1

        def parse_cell_reference(
            pszText: str,
        ) -> Tuple[str | None, int, int]:
            pszSheet: str | None = None
            pszCell: str = pszText
            if "!" in pszText:
                pszSheet, pszCell = pszText.split("!", 1)
                if pszSheet.startswith("'") and pszSheet.endswith("'"):
                    pszSheet = pszSheet[1:-1]
            pszCellStripped: str = pszCell.replace("$", "")
            pszColumnPart: str = ""
            pszRowPart: str = ""
            for ch in pszCellStripped:
                if ch.isalpha():
                    pszColumnPart += ch
                else:
                    pszRowPart += ch
            iRowIndex: int = int(pszRowPart) - 1 if len(pszRowPart) > 0 else 0
            iColumnIndex: int = column_letter_to_index(pszColumnPart)
            return pszSheet, iRowIndex, iColumnIndex

        def is_cell_reference(
            pszText: str,
        ) -> bool:
            pszWork: str = pszText
            if "!" in pszWork:
                pszWork = pszWork.split("!", 1)[1]
            pszWork = pszWork.replace("$", "")
            iIndex: int = 0
            while iIndex < len(pszWork) and pszWork[iIndex].isalpha():
                iIndex += 1
            if iIndex == 0:
                return False
            if iIndex >= len(pszWork):
                return False
            while iIndex < len(pszWork):
                if not pszWork[iIndex].isdigit():
                    return False
                iIndex += 1
            return True

        def tokenize_formula(
            pszFormula: str,
        ) -> List[Tuple[str, str]]:
            objTokens: List[Tuple[str, str]] = []
            i: int = 0
            while i < len(pszFormula):
                ch: str = pszFormula[i]
                if ch.isspace():
                    i += 1
                    continue
                if ch in ":,()":
                    objTokens.append(("symbol", ch))
                    i += 1
                    continue
                if ch in "+-*/&":
                    objTokens.append(("op", ch))
                    i += 1
                    continue
                if ch in "<>=" and i + 1 < len(pszFormula):
                    ch2: str = pszFormula[i + 1]
                    if (ch == "<" and ch2 in ["=", ">"]) or (ch == ">" and ch2 == "="):
                        objTokens.append(("cmp", ch + ch2))
                        i += 2
                        continue
                if ch in "<>":
                    objTokens.append(("cmp", ch))
                    i += 1
                    continue
                if ch == "=":
                    objTokens.append(("cmp", ch))
                    i += 1
                    continue
                if ch == "\"":
                    i += 1
                    pszLiteral: str = ""
                    while i < len(pszFormula):
                        if pszFormula[i] == "\"":
                            if i + 1 < len(pszFormula) and pszFormula[i + 1] == "\"":
                                pszLiteral += "\""
                                i += 2
                                continue
                            i += 1
                            break
                        pszLiteral += pszFormula[i]
                        i += 1
                    objTokens.append(("string", pszLiteral))
                    continue
                if ch.isdigit() or ch == ".":
                    pszNumber: str = ""
                    while i < len(pszFormula) and (pszFormula[i].isdigit() or pszFormula[i] == "."):
                        pszNumber += pszFormula[i]
                        i += 1
                    objTokens.append(("number", pszNumber))
                    continue
                pszIdentifier: str = ""
                while i < len(pszFormula):
                    ch2 = pszFormula[i]
                    if ch2.isspace() or ch2 in "+-*/&,()":
                        break
                    if ch2 in ["<", ">", "="]:
                        break
                    pszIdentifier += ch2
                    i += 1
                objTokens.append(("ident", pszIdentifier))
            return objTokens

        def parse_expression(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objValue, iPos = parse_comparison(objTokens, iStart)
            return objValue, iPos

        def parse_comparison(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_term(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "cmp":
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_term(objTokens, iPos + 1)
                objLeft = ("cmp", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_term(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_factor(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["+", "-", "&"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_factor(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_factor(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objLeft, iPos = parse_unary(objTokens, iStart)
            while iPos < len(objTokens) and objTokens[iPos][0] == "op" and objTokens[iPos][1] in ["*", "/"]:
                pszOp: str = objTokens[iPos][1]
                objRight, iPos = parse_unary(objTokens, iPos + 1)
                objLeft = ("op", pszOp, objLeft, objRight)
            return objLeft, iPos

        def parse_unary(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            if iStart < len(objTokens) and objTokens[iStart][0] == "op" and objTokens[iStart][1] in ["+", "-"]:
                pszOp: str = objTokens[iStart][1]
                objValue, iPos = parse_unary(objTokens, iStart + 1)
                return ("unary", pszOp, objValue), iPos
            return parse_primary(objTokens, iStart)

        def parse_primary(
            objTokens: List[Tuple[str, str]],
            iStart: int,
        ) -> Tuple[Any, int]:
            objTokenType, pszToken = objTokens[iStart]
            if objTokenType == "number":
                return ("number", pszToken), iStart + 1
            if objTokenType == "string":
                return ("string", pszToken), iStart + 1
            if objTokenType == "ident":
                if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", "("):
                    pszFuncName: str = pszToken.upper()
                    objArgs: List[Any] = []
                    iPos: int = iStart + 2
                    if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                        iPos += 1
                    else:
                        while iPos < len(objTokens):
                            objArg, iPos = parse_expression(objTokens, iPos)
                            objArgs.append(objArg)
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ","):
                                iPos += 1
                                continue
                            if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                                iPos += 1
                                break
                    return ("func", pszFuncName, objArgs), iPos
                if is_cell_reference(pszToken):
                    if iStart + 1 < len(objTokens) and objTokens[iStart + 1] == ("symbol", ":"):
                        pszSecond: str = objTokens[iStart + 2][1] if (iStart + 2) < len(objTokens) else ""
                        return ("range", pszToken, pszSecond), iStart + 3
                    return ("cell", pszToken), iStart + 1
                return ("string", pszToken), iStart + 1
            if objTokenType == "symbol" and pszToken == "(":
                objValue, iPos = parse_expression(objTokens, iStart + 1)
                if iPos < len(objTokens) and objTokens[iPos] == ("symbol", ")"):
                    iPos += 1
                return objValue, iPos
            return ("string", pszToken), iStart + 1

        def evaluate_node(
            objNode: Any,
            iCurrentRow: int,
            iCurrentColumn: int,
        ) -> Any:
            if objNode[0] == "number":
                try:
                    if "." in objNode[1]:
                        return float(objNode[1])
                    return int(objNode[1])
                except ValueError:
                    return 0
            if objNode[0] == "string":
                return str(objNode[1])
            if objNode[0] == "unary":
                pszOp: str = objNode[1]
                objValue: Any = evaluate_node(objNode[2], iCurrentRow, iCurrentColumn)
                try:
                    fNumber: float = float(objValue)
                    return -fNumber if pszOp == "-" else fNumber
                except Exception:
                    return objValue
            if objNode[0] == "op":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node(objNode[2], iCurrentRow, iCurrentColumn)
                objRight: Any = evaluate_node(objNode[3], iCurrentRow, iCurrentColumn)
                if pszOp == "&":
                    return str(objLeft) + str(objRight)
                try:
                    fLeft: float = float(objLeft)
                    fRight: float = float(objRight)
                    if pszOp == "+":
                        return fLeft + fRight
                    if pszOp == "-":
                        return fLeft - fRight
                    if pszOp == "*":
                        return fLeft * fRight
                    if pszOp == "/":
                        return fLeft / fRight if fRight != 0 else 0
                except Exception:
                    if pszOp == "+":
                        return str(objLeft) + str(objRight)
                return ""
            if objNode[0] == "cmp":
                pszOp: str = objNode[1]
                objLeft: Any = evaluate_node(objNode[2], iCurrentRow, iCurrentColumn)
                objRight: Any = evaluate_node(objNode[3], iCurrentRow, iCurrentColumn)
                if pszOp == "=":
                    return objLeft == objRight
                if pszOp == "<":
                    return objLeft < objRight
                if pszOp == ">":
                    return objLeft > objRight
                if pszOp == "<=":
                    return objLeft <= objRight
                if pszOp == ">=":
                    return objLeft >= objRight
                if pszOp == "<>":
                    return objLeft != objRight
            if objNode[0] == "cell":
                pszSheet, iRow, iColumn = parse_cell_reference(objNode[1])
                ensure_valid_reference(pszSheet, iRow, iColumn, iCurrentRow, iCurrentColumn)
                return evaluate_cell_by_sheet(
                    pszSheet,
                    iRow,
                    iColumn,
                    iCurrentRow,
                    iCurrentColumn,
                )
            if objNode[0] == "range":
                pszSheetStart, _, _ = parse_cell_reference(objNode[1])
                pszSheetEnd, _, _ = parse_cell_reference(objNode[2])
                pszSheetNameRange: str | None = pszSheetStart or pszSheetEnd
                pszStartCell: str = objNode[1].split("!", 1)[-1]
                pszEndCell: str = objNode[2].split("!", 1)[-1]
                pszStartSheet, iStartRow, iStartColumn = parse_cell_reference(
                    (pszSheetNameRange + "!" if pszSheetNameRange else "") + pszStartCell,
                )
                pszEndSheet, iEndRow, iEndColumn = parse_cell_reference(
                    (pszSheetNameRange + "!" if pszSheetNameRange else "") + pszEndCell,
                )
                ensure_valid_range_reference(
                    pszStartSheet or pszEndSheet,
                    iStartRow,
                    iStartColumn,
                    iEndRow,
                    iEndColumn,
                    iCurrentRow,
                    iCurrentColumn,
                )
                return evaluate_range(
                    pszSheetNameRange,
                    pszStartCell,
                    pszEndCell,
                    iCurrentRow,
                    iCurrentColumn,
                )
            if objNode[0] == "func":
                pszFuncName: str = objNode[1]
                objArgsNodes: List[Any] = objNode[2]
                objArgsValues: List[Any] = [
                    evaluate_node(objArg, iCurrentRow, iCurrentColumn) for objArg in objArgsNodes
                ]
                if pszFuncName == "IF":
                    bCondition: bool = bool(objArgsValues[0]) if len(objArgsValues) > 0 else False
                    if bCondition:
                        return objArgsValues[1] if len(objArgsValues) > 1 else ""
                    return objArgsValues[2] if len(objArgsValues) > 2 else ""
                if pszFuncName == "SUM":
                    fTotal: float = 0.0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                try:
                                    fTotal += float(objItem)
                                except Exception:
                                    continue
                        else:
                            try:
                                fTotal += float(objArg)
                            except Exception:
                                continue
                    return fTotal
                if pszFuncName == "COUNT":
                    iCount: int = 0
                    for objArg in objArgsValues:
                        if isinstance(objArg, list):
                            for objItem in objArg:
                                if objItem not in ["", None]:
                                    iCount += 1
                        else:
                            if objArg not in ["", None]:
                                iCount += 1
                    return iCount
                if pszFuncName == "TEXT":
                    return str(objArgsValues[0]) if len(objArgsValues) > 0 else ""
                if pszFuncName == "CONCAT":
                    return "".join(str(objArg) for objArg in objArgsValues)
            return ""

        def evaluate_range(
            pszSheet: str | None,
            pszStartCell: str,
            pszEndCell: str,
            iCurrentRow: int,
            iCurrentColumn: int,
        ) -> List[Any]:
            pszStartSheet, iStartRow, iStartColumn = parse_cell_reference(
                (pszSheet + "!" if pszSheet else "") + pszStartCell,
            )
            pszEndSheet, iEndRow, iEndColumn = parse_cell_reference(
                (pszSheet + "!" if pszSheet else "") + pszEndCell,
            )
            pszTargetSheet: str | None = pszStartSheet or pszEndSheet
            objValues: List[Any] = []
            for iRow in range(min(iStartRow, iEndRow), max(iStartRow, iEndRow) + 1):
                for iColumn in range(min(iStartColumn, iEndColumn), max(iStartColumn, iEndColumn) + 1):
                    objValues.append(
                        evaluate_cell_by_sheet(
                            pszTargetSheet,
                            iRow,
                            iColumn,
                            iCurrentRow,
                            iCurrentColumn,
                        ),
                    )
            return objValues

        objRawDataSheet: pd.DataFrame = load_tsv_as_dataframe(pszRawDataTsvPath, True)
        if objRawDataSheet.empty:
            return 1

        pszWithSalaryFormulaTsvPath: str = str(objBaseDirectoryPath / "With_Salary_Formula.tsv")
        objWithSalaryFormulaSheet: pd.DataFrame = load_tsv_as_dataframe(pszWithSalaryFormulaTsvPath, True)
        if objWithSalaryFormulaSheet.empty:
            return 1

        objOutputWithSalary: pd.DataFrame = pd.DataFrame(
            "",
            index=objWithSalaryFormulaSheet.index,
            columns=objWithSalaryFormulaSheet.columns,
        )

        objCache: Dict[Tuple[int, int], Any] = {}

        class CircularReferenceError(RuntimeError):
            pass

        class InvalidReferenceError(RuntimeError):
            pass

        objEvaluating: set[Tuple[int, int]] = set()

        def ensure_valid_reference(
            pszSheetName: str | None,
            iTargetRow: int,
            iTargetColumn: int,
            iCurrentRow: int,
            iCurrentColumn: int,
        ) -> None:
            if pszSheetName is None or pszSheetName == "給与あり":
                if iTargetRow != iCurrentRow or iTargetColumn >= iCurrentColumn:
                    raise InvalidReferenceError(
                        "Unsupported cell reference at 給与あり!R{0}C{1}".format(
                            iCurrentRow + 1,
                            iCurrentColumn + 1,
                        ),
                    )

        def ensure_valid_range_reference(
            pszSheetName: str | None,
            iStartRow: int,
            iStartColumn: int,
            iEndRow: int,
            iEndColumn: int,
            iCurrentRow: int,
            iCurrentColumn: int,
        ) -> None:
            if pszSheetName is None or pszSheetName == "給与あり":
                if iStartRow != iCurrentRow or iEndRow != iCurrentRow:
                    raise InvalidReferenceError(
                        "Unsupported range reference at 給与あり!R{0}C{1}".format(
                            iCurrentRow + 1,
                            iCurrentColumn + 1,
                        ),
                    )
                if max(iStartColumn, iEndColumn) >= iCurrentColumn:
                    raise InvalidReferenceError(
                        "Unsupported range reference at 給与あり!R{0}C{1}".format(
                            iCurrentRow + 1,
                            iCurrentColumn + 1,
                        ),
                    )

        def evaluate_cell_by_sheet(
            pszSheetName: str | None,
            iRow: int,
            iColumn: int,
            iCurrentRow: int,
            iCurrentColumn: int,
        ) -> Any:
            if pszSheetName is None or pszSheetName == "給与あり":
                return evaluate_cell(iRow, iColumn)
            if pszSheetName == "ローデータ":
                if (iRow < 0) or (iColumn < 0):
                    return ""
                if iRow >= objRawDataSheet.shape[0] or iColumn >= objRawDataSheet.shape[1]:
                    return ""
                objValue = objRawDataSheet.iat[iRow, iColumn]
                return "" if pd.isna(objValue) else objValue
            raise InvalidReferenceError(
                "Unsupported sheet reference at 給与あり!R{0}C{1}".format(
                    iCurrentRow + 1,
                    iCurrentColumn + 1,
                ),
            )

        def evaluate_cell(
            iRow: int,
            iColumn: int,
        ) -> Any:
            if (iRow, iColumn) in objCache:
                return objCache[(iRow, iColumn)]
            if (iRow, iColumn) in objEvaluating:
                raise CircularReferenceError(
                    "Error: circular reference detected at 給与あり!R{0}C{1}".format(
                        iRow + 1,
                        iColumn + 1,
                    ),
                )
            if (iRow < 0) or (iColumn < 0):
                return ""
            if iRow >= objWithSalaryFormulaSheet.shape[0] or iColumn >= objWithSalaryFormulaSheet.shape[1]:
                return ""
            objValue = objWithSalaryFormulaSheet.iat[iRow, iColumn]
            pszValueStr: str = "" if pd.isna(objValue) else str(objValue)
            if len(pszValueStr) > 0 and pszValueStr.startswith("="):
                pszFormula: str = pszValueStr[1:]
                objTokens = tokenize_formula(pszFormula)
                objAst, _ = parse_expression(objTokens, 0)
                objResult: Any = objAst
                objEvaluating.add((iRow, iColumn))
                try:
                    objResult = evaluate_node(objAst, iRow, iColumn)
                except CircularReferenceError as objCircularError:
                    pszErrorPath: str = str(objBaseDirectoryPath / "With_Salary_error.tsv")
                    write_error_tsv_with_salary(pszErrorPath, str(objCircularError))
                    raise
                except Exception as objException:
                    pszErrorPath: str = str(objBaseDirectoryPath / "With_Salary_error.tsv")
                    write_error_tsv_with_salary(
                        pszErrorPath,
                        "Error: formula evaluation failed at {0}!R{1}C{2}: {3}".format(
                            "給与あり",
                            iRow + 1,
                            iColumn + 1,
                            objException,
                        ),
                    )
                    raise
                finally:
                    objEvaluating.discard((iRow, iColumn))
                objCache[(iRow, iColumn)] = objResult
                return objResult
            objCache[(iRow, iColumn)] = pszValueStr
            return pszValueStr

        for iRowIndex in range(objWithSalaryFormulaSheet.shape[0]):
            for iColumnIndex in range(objWithSalaryFormulaSheet.shape[1]):
                try:
                    objEvaluatedValue: Any = evaluate_cell(iRowIndex, iColumnIndex)
                except Exception:
                    return 1
                objOutputWithSalary.iat[iRowIndex, iColumnIndex] = objEvaluatedValue

        pszWithSalaryOutputPath: str = str(objBaseDirectoryPath / "With_Salary.tsv")
        objOutputWithSalary.to_csv(
            pszWithSalaryOutputPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n",
        )
    except Exception:
        pass

    print("OK: created files")
    print(pszStep1TsvPath)
    print(pszStep2TsvPath)
    print(pszStep3TsvPath)
    print(pszSheet4TsvPath)
    print(pszSheet4StaffCodeRangeTsvPath)
    print(pszSheet6TsvPath)
    print(pszSheet7TsvPath)
    print(pszSheet8TsvPath)
    print(pszSheet9TsvPath)
    print(pszProjectListTsvPath)

    return 0


if __name__ == "__main__":
    iExitCode: int = main()
    raise SystemExit(iExitCode)
'''


def _load_module_from_source(name: str, source: str) -> dict:
    module_dict: dict = {"__name__": name}
    exec(compile(source, f"<{name}>", "exec"), module_dict)
    return module_dict


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python Make_PjSummary_PL_Manhour.py <csv_file> [<csv_file> ...]")
        return 1

    original_argv = sys.argv[:]
    pl_module = _load_module_from_source("pl_csv_to_tsv_cmd", PL_SOURCE)
    manhour_module = _load_module_from_source("make_manhour_to_sheet8", MANHOUR_SOURCE)
    pl_main = pl_module.get("main")
    manhour_main = manhour_module.get("main")
    if pl_main is None or manhour_main is None:
        raise RuntimeError("Failed to load main() from embedded sources.")

    exit_code = 0
    for input_path in original_argv[1:]:
        sys.argv = [original_argv[0], input_path]
        exit_code |= int(pl_main())
        exit_code |= int(manhour_main())
    sys.argv = original_argv
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
