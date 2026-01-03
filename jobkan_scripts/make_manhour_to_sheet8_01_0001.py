###############################################################
# Main editable file
#
# All implementation must be done in this file.
# Other Python files under src/ are reference-only
# and must not be modified.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
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
# ===============================================================

from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Any, Dict

# ===============================================================
# EMBEDDED SOURCE: csv_to_tsv_h_mm_ss.py
# ===============================================================
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


###############################################################
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
###############################################################
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


###############################################################
#
# 工数文字列 "h:mm" を "h:mm:ss" に揃える関数。
# 例: "7:30" -> "7:30:00"
#
###############################################################
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


###############################################################
#
# 入力CSVを読み込み、TSVに変換する関数。
#
###############################################################
def convert_csv_to_tsv_file(
    pszInputCsvPath: str,
) -> str:
    if not os.path.exists(pszInputCsvPath):
        raise FileNotFoundError(f"Input CSV not found: {pszInputCsvPath}")

    pszOutputTsvPath: str = build_output_file_full_path(pszInputCsvPath, ".tsv")

    objRows: List[List[str]] = []

    with open(pszInputCsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile)
        for objRow in objReader:
            objRows.append(list(objRow))

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


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python csv_to_tsv_h_mm_ss.py <input_manhour_csv>")
        return 1

    pszInputCsvPath: str = sys.argv[1]

    pszOutputTsvPath: str = convert_csv_to_tsv_file(pszInputCsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ===============================================================
# EMBEDDED SOURCE: manhour_remove_uninput_rows.py
# ===============================================================
pszSource_manhour_remove_uninput_rows_py: str = r'''# -*- coding: utf-8 -*-
"""
manhour_remove_uninput_rows.py

ジョブカン工数 TSV を読み込み、
(未入力) の行を除去して TSV を出力する。

入力:  manhour_*.tsv
出力:  manhour_*_removed_uninput.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import sys
import os
import csv
from typing import List


###############################################################
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
###############################################################
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


###############################################################
#
# TSV の各行から、「未入力」行を除外して出力する関数。
#
###############################################################
def make_removed_uninput_tsv_from_manhour_tsv(
    pszInputTsvPath: str,
) -> str:
    if not os.path.exists(pszInputTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputTsvPath}")

    pszOutputTsvPath: str = build_output_file_full_path(pszInputTsvPath, "_removed_uninput.tsv")

    with open(pszInputTsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objAllRows: List[List[str]] = [list(objRow) for objRow in objReader]

    if len(objAllRows) <= 1:
        with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            for objRow in objAllRows:
                objWriter.writerow(objRow)
        return pszOutputTsvPath

    objHeader: List[str] = objAllRows[0]
    objBodyRows: List[List[str]] = objAllRows[1:]

    # G〜J の 4 列 (index 6〜9) がすべて "未入力" の行を除外する。
    iStartIndex: int = 6
    iEndIndex: int = 10

    objFilteredRows: List[List[str]] = []

    for objRow in objBodyRows:
        if len(objRow) < iEndIndex:
            # 列数不足の行はそのまま残す。
            objFilteredRows.append(objRow)
            continue

        bAllUninput: bool = True
        for iColumnIndex in range(iStartIndex, iEndIndex):
            if objRow[iColumnIndex] != "未入力":
                bAllUninput = False
                break

        if bAllUninput:
            continue

        objFilteredRows.append(objRow)

    with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        objWriter.writerow(objHeader)
        for objRow in objFilteredRows:
            objWriter.writerow(objRow)

    return pszOutputTsvPath


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python manhour_remove_uninput_rows.py <input_manhour_tsv>")
        return 1

    pszInputTsvPath: str = sys.argv[1]

    pszOutputTsvPath: str = make_removed_uninput_tsv_from_manhour_tsv(pszInputTsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ===============================================================
# EMBEDDED SOURCE: sort_manhour_by_staff_code.py
# ===============================================================
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

import sys
import os
import csv
from typing import List, Tuple


###############################################################
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
###############################################################
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


###############################################################
#
# スタッフコード順で並び替えた TSV を生成する関数。
#
###############################################################
def make_sorted_staff_code_tsv_from_manhour_tsv(
    pszInputTsvPath: str,
) -> str:
    if not os.path.exists(pszInputTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputTsvPath}")

    pszOutputTsvPath: str = build_output_file_full_path(pszInputTsvPath, "_sorted_staff_code.tsv")

    with open(pszInputTsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objAllRows: List[List[str]] = [list(objRow) for objRow in objReader]

    if len(objAllRows) <= 1:
        with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            for objRow in objAllRows:
                objWriter.writerow(objRow)
        return pszOutputTsvPath

    objHeader: List[str] = objAllRows[0]
    objBodyRows: List[List[str]] = objAllRows[1:]

    # スタッフコード(列B) = index 1
    iStaffCodeIndex: int = 1

    objSortable: List[Tuple[int, List[str]]] = []
    objNonSortable: List[List[str]] = []

    for objRow in objBodyRows:
        if len(objRow) <= iStaffCodeIndex:
            objNonSortable.append(objRow)
            continue

        pszStaffCodeText: str = (objRow[iStaffCodeIndex] or "").strip()
        try:
            iStaffCode: int = int(pszStaffCodeText)
            objSortable.append((iStaffCode, objRow))
        except ValueError:
            objNonSortable.append(objRow)

    objSortable.sort(key=lambda objItem: objItem[0])

    objSortedRows: List[List[str]] = [objItem[1] for objItem in objSortable] + objNonSortable

    with open(pszOutputTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        objWriter.writerow(objHeader)
        for objRow in objSortedRows:
            objWriter.writerow(objRow)

    return pszOutputTsvPath


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python sort_manhour_by_staff_code.py <input_manhour_tsv>")
        return 1

    pszInputTsvPath: str = sys.argv[1]

    pszOutputTsvPath: str = make_sorted_staff_code_tsv_from_manhour_tsv(pszInputTsvPath)
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ===============================================================
# EMBEDDED SOURCE: convert_yyyy_mm_dd.py
# ===============================================================
pszSource_convert_yyyy_mm_dd_py: str = r'''# -*- coding: utf-8 -*-
"""
convert_yyyy_mm_dd.py

ジョブカン工数 TSV を読み込み、
日付列(列C)を yyyy-mm-dd 形式に揃えて Sheet4.tsv を出力する。

入力:  manhour_*_removed_uninput_sorted_staff_code.tsv
出力:  Sheet4.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import sys
import os
import csv
import re
from typing import List


###############################################################
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
###############################################################
def build_output_file_full_path(
    pszInputFileFullPath: str, 
    pszOutputFileName: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszOutputFileFullPath: str = os.path.join(pszDirectory, pszOutputFileName)
    return pszOutputFileFullPath


###############################################################
#
# "yyyy/m/d" または "yyyy-mm-dd" 等を "yyyy-mm-dd" に揃える関数。
#
###############################################################
def normalize_yyyy_mm_dd(
    pszDateText: str,
) -> str:
    pszText: str = (pszDateText or "").strip()
    if pszText == "":
        return ""

    objMatch = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", pszText)
    if not objMatch:
        return pszText

    iYear: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iDay: int = int(objMatch.group(3))

    return f"{iYear:04d}-{iMonth:02d}-{iDay:02d}"


###############################################################
#
# 入力TSVを読み込み、日付列(列C)を正規化して Sheet4.tsv を出力する関数。
#
###############################################################
def make_sheet4_tsv_from_input_tsv(
    pszInputTsvPath: str,
    pszOutputSheet4TsvPath: str,
) -> str:
    if not os.path.exists(pszInputTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputTsvPath}")

    with open(pszInputTsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objAllRows: List[List[str]] = [list(objRow) for objRow in objReader]

    if len(objAllRows) == 0:
        with open(pszOutputSheet4TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            pass
        return pszOutputSheet4TsvPath

    # 日付列(列C) = index 2
    iDateColumnIndex: int = 2

    for iRowIndex in range(1, len(objAllRows)):
        objRow: List[str] = objAllRows[iRowIndex]
        if len(objRow) <= iDateColumnIndex:
            continue
        objRow[iDateColumnIndex] = normalize_yyyy_mm_dd(objRow[iDateColumnIndex])
        objAllRows[iRowIndex] = objRow

    with open(pszOutputSheet4TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objAllRows:
            objWriter.writerow(objRow)

    return pszOutputSheet4TsvPath


###############################################################
#
# main
#
###############################################################
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

# ===============================================================
# EMBEDDED SOURCE: make_staff_code_range.py
# ===============================================================
pszSource_make_staff_code_range_py: str = r'''# -*- coding: utf-8 -*-
"""
make_staff_code_range.py

Sheet4.tsv を読み込み、
スタッフコードの最初と最後の行番号の範囲(開始行, 終了行)を求めて出力する。

入力:  Sheet4.tsv
出力:  Sheet4_staff_code_range.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8

範囲は 1-based の Excel 行番号で出力する。
"""

import sys
import os
import csv
from typing import List, Dict, Tuple


###############################################################
#
# 指定した入力ファイルパスから、出力ファイルパスを作る関数。
#
###############################################################
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


###############################################################
#
# スタッフコード範囲 TSV を作る関数。
#
###############################################################
def make_staff_code_range_tsv_from_sheet1_tsv(
    pszInputSheet1TsvPath: str,
    pszOutputStaffCodeRangeTsvPath: str,
) -> str:
    if not os.path.exists(pszInputSheet1TsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputSheet1TsvPath}")

    with open(pszInputSheet1TsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objAllRows: List[List[str]] = [list(objRow) for objRow in objReader]

    if len(objAllRows) <= 1:
        # ヘッダのみの場合
        with open(pszOutputStaffCodeRangeTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            objWriter.writerow(["スタッフコード", "開始行", "終了行"])
        return pszOutputStaffCodeRangeTsvPath

    # スタッフコード列(列B) = index 1
    iStaffCodeIndex: int = 1

    # スタッフコードごとの開始/終了行を集計する
    # 行番号は Excel の行番号を想定するため 1-based とする。
    #   Sheet4.tsv の 1行目はヘッダ行なので、
    #   データ行は 2行目から始まる。
    objStartRowMap: Dict[str, int] = {}
    objEndRowMap: Dict[str, int] = {}

    for iRowIndex in range(1, len(objAllRows)):
        objRow: List[str] = objAllRows[iRowIndex]
        if len(objRow) <= iStaffCodeIndex:
            continue

        pszStaffCode: str = (objRow[iStaffCodeIndex] or "").strip()
        if pszStaffCode == "":
            continue

        iExcelRowNumber: int = iRowIndex + 1

        if pszStaffCode not in objStartRowMap:
            objStartRowMap[pszStaffCode] = iExcelRowNumber
        objEndRowMap[pszStaffCode] = iExcelRowNumber

    objStaffCodeList: List[str] = sorted(objStartRowMap.keys(), key=lambda pszText: int(pszText))

    with open(pszOutputStaffCodeRangeTsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        objWriter.writerow(["スタッフコード", "開始行", "終了行"])
        for pszStaffCode in objStaffCodeList:
            iStartRow: int = objStartRowMap[pszStaffCode]
            iEndRow: int = objEndRowMap[pszStaffCode]
            objWriter.writerow([pszStaffCode, str(iStartRow), str(iEndRow)])

    return pszOutputStaffCodeRangeTsvPath


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python make_staff_code_range.py <Sheet4.tsv>")
        return 1

    pszInputSheet1TsvPath: str = sys.argv[1]
    pszOutputStaffCodeRangeTsvPath: str = build_output_file_full_path(pszInputSheet1TsvPath, "_staff_code_range.tsv")

    pszOutputTsvPath: str = make_staff_code_range_tsv_from_sheet1_tsv(
        pszInputSheet1TsvPath,
        pszOutputStaffCodeRangeTsvPath,
    )
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ===============================================================
# EMBEDDED SOURCE: make_sheet6_from_sheet4.py
# ===============================================================
pszSource_make_sheet6_from_sheet4_py: str = r'''# -*- coding: utf-8 -*-
"""
make_sheet6_from_sheet4.py

Sheet4.tsv と Sheet4_staff_code_range.tsv を読み込み、
Sheet6.tsv を生成する。

入力:
  - Sheet4.tsv
  - Sheet4_staff_code_range.tsv

出力:
  - Sheet6.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import sys
import os
import csv
from typing import List, Dict, Tuple


###############################################################
#
# Sheet4.tsv と Sheet4_staff_code_range.tsv を読み込み、Sheet6.tsv を生成する関数。
#
###############################################################
def make_sheet6_from_sheet4(
    pszInputSheet4TsvPath: str,
    pszInputStaffCodeRangeTsvPath: str,
    pszOutputSheet6TsvPath: str,
) -> str:
    if not os.path.exists(pszInputSheet4TsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputSheet4TsvPath}")

    if not os.path.exists(pszInputStaffCodeRangeTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputStaffCodeRangeTsvPath}")

    with open(pszInputSheet4TsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objSheet4Rows: List[List[str]] = [list(objRow) for objRow in objReader]

    with open(pszInputStaffCodeRangeTsvPath, mode="r", encoding="utf-8", newline="") as objRangeFile:
        objRangeReader: csv.reader = csv.reader(objRangeFile, delimiter="\t")
        objRangeRows: List[List[str]] = [list(objRow) for objRow in objRangeReader]

    # Sheet4 の行数
    iSheet4RowCount: int = len(objSheet4Rows)

    if iSheet4RowCount <= 1:
        with open(pszOutputSheet6TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
            objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
            objWriter.writerow(["スタッフコード", "氏名", "開始行", "終了行"])
        return pszOutputSheet6TsvPath

    # Sheet4 のスタッフコード列(列B) = index 1
    iStaffCodeIndex: int = 1
    # Sheet4 の氏名列(列D) = index 3
    iNameIndex: int = 3

    # スタッフコード -> 氏名 を作る
    objStaffCodeToName: Dict[str, str] = {}
    for iRowIndex in range(1, len(objSheet4Rows)):
        objRow: List[str] = objSheet4Rows[iRowIndex]
        if len(objRow) <= max(iStaffCodeIndex, iNameIndex):
            continue
        pszStaffCode: str = (objRow[iStaffCodeIndex] or "").strip()
        pszName: str = (objRow[iNameIndex] or "").strip()
        if pszStaffCode == "":
            continue
        if pszStaffCode not in objStaffCodeToName:
            objStaffCodeToName[pszStaffCode] = pszName

    # 範囲TSVのヘッダ行を除外
    objRangeBody: List[List[str]] = objRangeRows[1:] if len(objRangeRows) >= 2 else []

    with open(pszOutputSheet6TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        objWriter.writerow(["スタッフコード", "氏名", "開始行", "終了行"])

        for objRangeRow in objRangeBody:
            if len(objRangeRow) < 3:
                continue
            pszStaffCode: str = (objRangeRow[0] or "").strip()
            pszStartRow: str = (objRangeRow[1] or "").strip()
            pszEndRow: str = (objRangeRow[2] or "").strip()
            pszName: str = objStaffCodeToName.get(pszStaffCode, "")

            objWriter.writerow([pszStaffCode, pszName, pszStartRow, pszEndRow])

    return pszOutputSheet6TsvPath


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: python make_sheet6_from_sheet4.py <Sheet4.tsv> <Sheet4_staff_code_range.tsv>")
        return 1

    pszInputSheet4TsvPath: str = sys.argv[1]
    pszInputStaffCodeRangeTsvPath: str = sys.argv[2]

    pszDirectory: str = os.path.dirname(pszInputSheet4TsvPath)
    pszOutputSheet6TsvPath: str = os.path.join(pszDirectory, "Sheet6.tsv")

    pszOutputTsvPath: str = make_sheet6_from_sheet4(
        pszInputSheet4TsvPath,
        pszInputStaffCodeRangeTsvPath,
        pszOutputSheet6TsvPath,
    )
    print(f"Output: {pszOutputTsvPath}")
    return 0
'''

# ===============================================================
# EMBEDDED SOURCE: make_sheet789_from_sheet4.py
# ===============================================================
pszSource_make_sheet789_from_sheet4_py: str = r'''# -*- coding: utf-8 -*-
"""
make_sheet789_from_sheet4.py

Sheet4.tsv と Sheet4_staff_code_range.tsv と Sheet6.tsv を読み込み、
Sheet7.tsv / Sheet8.tsv / Sheet9.tsv を生成する。

入力:
  - Sheet4.tsv
  - Sheet4_staff_code_range.tsv
  - Sheet6.tsv

出力:
  - Sheet7.tsv
  - Sheet8.tsv
  - Sheet9.tsv

注意:
  - 入力 TSV は UTF-8
  - 出力 TSV も UTF-8
"""

import sys
import os
import csv
from typing import List, Dict


###############################################################
#
# Sheet4.tsv などを読み込み、Sheet7/8/9.tsv を生成する関数。
#
###############################################################
def make_sheet789_from_sheet4(
    pszInputSheet4TsvPath: str,
    pszInputStaffCodeRangeTsvPath: str,
    pszInputSheet6TsvPath: str,
    pszOutputSheet7TsvPath: str,
    pszOutputSheet8TsvPath: str,
    pszOutputSheet9TsvPath: str,
) -> None:
    if not os.path.exists(pszInputSheet4TsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputSheet4TsvPath}")

    if not os.path.exists(pszInputStaffCodeRangeTsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputStaffCodeRangeTsvPath}")

    if not os.path.exists(pszInputSheet6TsvPath):
        raise FileNotFoundError(f"Input TSV not found: {pszInputSheet6TsvPath}")

    with open(pszInputSheet4TsvPath, mode="r", encoding="utf-8", newline="") as objInputFile:
        objReader: csv.reader = csv.reader(objInputFile, delimiter="\t")
        objSheet4Rows: List[List[str]] = [list(objRow) for objRow in objReader]

    with open(pszInputStaffCodeRangeTsvPath, mode="r", encoding="utf-8", newline="") as objRangeFile:
        objRangeReader: csv.reader = csv.reader(objRangeFile, delimiter="\t")
        objRangeRows: List[List[str]] = [list(objRow) for objRow in objRangeReader]

    with open(pszInputSheet6TsvPath, mode="r", encoding="utf-8", newline="") as objSheet6File:
        objSheet6Reader: csv.reader = csv.reader(objSheet6File, delimiter="\t")
        objSheet6Rows: List[List[str]] = [list(objRow) for objRow in objSheet6Reader]

    # Sheet4 のヘッダ除外
    objSheet4Body: List[List[str]] = objSheet4Rows[1:] if len(objSheet4Rows) >= 2 else []

    # 範囲TSVのヘッダ除外
    objRangeBody: List[List[str]] = objRangeRows[1:] if len(objRangeRows) >= 2 else []

    # Sheet6 のヘッダ除外
    objSheet6Body: List[List[str]] = objSheet6Rows[1:] if len(objSheet6Rows) >= 2 else []

    # Sheet6: スタッフコード -> 氏名
    objStaffCodeToName: Dict[str, str] = {}
    for objRow in objSheet6Body:
        if len(objRow) < 2:
            continue
        pszStaffCode: str = (objRow[0] or "").strip()
        pszName: str = (objRow[1] or "").strip()
        if pszStaffCode == "":
            continue
        objStaffCodeToName[pszStaffCode] = pszName

    # Sheet7: スタッフコードごとの範囲をそのまま出す (ユーザー版仕様)
    with open(pszOutputSheet7TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        objWriter.writerow(["スタッフコード", "氏名", "開始行", "終了行"])

        for objRangeRow in objRangeBody:
            if len(objRangeRow) < 3:
                continue

            pszStaffCode: str = (objRangeRow[0] or "").strip()
            pszStartRow: str = (objRangeRow[1] or "").strip()
            pszEndRow: str = (objRangeRow[2] or "").strip()
            pszName: str = objStaffCodeToName.get(pszStaffCode, "")

            objWriter.writerow([pszStaffCode, pszName, pszStartRow, pszEndRow])

    # Sheet8: Sheet4 の一部列を抽出して出す (ユーザー版仕様)
    #   Sheet4 の列構成はユーザーの正解版に合わせる。
    #   本スクリプトは「Sheet4 の全行をそのまま Sheet8 として出す」方式。
    with open(pszOutputSheet8TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objSheet4Rows:
            objWriter.writerow(objRow)

    # Sheet9: Sheet6 をそのまま出す (ユーザー版仕様)
    with open(pszOutputSheet9TsvPath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objWriter: csv.writer = csv.writer(objOutputFile, delimiter="\t")
        for objRow in objSheet6Rows:
            objWriter.writerow(objRow)


###############################################################
#
# main
#
###############################################################
def main() -> int:
    if len(sys.argv) < 4:
        print("Usage: python make_sheet789_from_sheet4.py <Sheet4.tsv> <Sheet4_staff_code_range.tsv> <Sheet6.tsv>")
        return 1

    pszInputSheet4TsvPath: str = sys.argv[1]
    pszInputStaffCodeRangeTsvPath: str = sys.argv[2]
    pszInputSheet6TsvPath: str = sys.argv[3]

    pszDirectory: str = os.path.dirname(pszInputSheet4TsvPath)
    pszOutputSheet7TsvPath: str = os.path.join(pszDirectory, "Sheet7.tsv")
    pszOutputSheet8TsvPath: str = os.path.join(pszDirectory, "Sheet8.tsv")
    pszOutputSheet9TsvPath: str = os.path.join(pszDirectory, "Sheet9.tsv")

    make_sheet789_from_sheet4(
        pszInputSheet4TsvPath,
        pszInputStaffCodeRangeTsvPath,
        pszInputSheet6TsvPath,
        pszOutputSheet7TsvPath,
        pszOutputSheet8TsvPath,
        pszOutputSheet9TsvPath,
    )

    print(f"Output: {pszOutputSheet7TsvPath}")
    print(f"Output: {pszOutputSheet8TsvPath}")
    print(f"Output: {pszOutputSheet9TsvPath}")
    return 0
'''

# ===============================================================
#
# 各スクリプトを独立名前空間で実行するためのヘルパー
#
# ===============================================================
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


# ===============================================================
#
# エラー内容を UTF-8 テキストとして書き出す関数
#
# ===============================================================
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


# ===============================================================
#
# main
#
# ===============================================================
def main() -> int:
    objParser: argparse.ArgumentParser = argparse.ArgumentParser()
    objParser.add_argument(
        "pszInputManhourCsvPath",
        help="Input Jobcan manhour CSV file path",
    )
    objArgs: argparse.Namespace = objParser.parse_args()

    pszInputManhourCsvPath: str = objArgs.pszInputManhourCsvPath
    objInputPath: Path = Path(pszInputManhourCsvPath)

    if not objInputPath.exists():
        objScriptDirectoryPath: Path = Path(__file__).resolve().parent
        objCandidatePath: Path = objScriptDirectoryPath / pszInputManhourCsvPath
        if objCandidatePath.exists():
            objInputPath = objCandidatePath

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
    objModuleCsvToTsv["convert_csv_to_tsv_file"](
        str(objInputPath),
    )
    pszStep1TsvPath: str = objModuleCsvToTsv["build_output_file_full_path"](
        str(objInputPath),
        ".tsv",
    )

    # (2) 未入力行除去
    objModuleRemoveUninput: Dict[str, Any] = create_module_from_source(
        "manhour_remove_uninput_rows",
        pszSource_manhour_remove_uninput_rows_py,
    )
    pszStep2TsvPath: str = objModuleRemoveUninput["make_removed_uninput_tsv_from_manhour_tsv"](
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
        "_sorted_staff_code.tsv",
    )

    # (4) 日付正規化 → Sheet4.tsv
    objModuleConvertDate: Dict[str, Any] = create_module_from_source(
        "convert_yyyy_mm_dd",
        pszSource_convert_yyyy_mm_dd_py,
    )
    pszSheet4TsvPath: str = str(objBaseDirectoryPath / "Sheet4.tsv")
    objModuleConvertDate["make_sheet4_tsv_from_input_tsv"](
        pszStep3TsvPath,
        pszSheet4TsvPath,
    )

    # (5) Sheet4_staff_code_range.tsv
    objModuleMakeRange: Dict[str, Any] = create_module_from_source(
        "make_staff_code_range",
        pszSource_make_staff_code_range_py,
    )
    pszSheet4StaffCodeRangeTsvPath: str = str(objBaseDirectoryPath / "Sheet4_staff_code_range.tsv")
    objModuleMakeRange["make_staff_code_range_tsv_from_sheet1_tsv"](
        pszSheet4TsvPath,
        pszSheet4StaffCodeRangeTsvPath,
    )

    # (6) Sheet6.tsv
    objModuleMakeSheet6: Dict[str, Any] = create_module_from_source(
        "make_sheet6_from_sheet4",
        pszSource_make_sheet6_from_sheet4_py,
    )
    pszSheet6TsvPath: str = str(objBaseDirectoryPath / "Sheet6.tsv")
    objModuleMakeSheet6["make_sheet6_from_sheet4"](
        pszSheet4TsvPath,
        pszSheet4StaffCodeRangeTsvPath,
        pszSheet6TsvPath,
    )

    # (7) Sheet7.tsv / Sheet8.tsv / Sheet9.tsv
    objModuleMakeSheet789: Dict[str, Any] = create_module_from_source(
        "make_sheet789_from_sheet4",
        pszSource_make_sheet789_from_sheet4_py,
    )
    pszSheet7TsvPath: str = str(objBaseDirectoryPath / "Sheet7.tsv")
    pszSheet8TsvPath: str = str(objBaseDirectoryPath / "Sheet8.tsv")
    pszSheet9TsvPath: str = str(objBaseDirectoryPath / "Sheet9.tsv")
    objModuleMakeSheet789["make_sheet789_from_sheet4"](
        pszSheet4TsvPath,
        pszSheet4StaffCodeRangeTsvPath,
        pszSheet6TsvPath,
        pszSheet7TsvPath,
        pszSheet8TsvPath,
        pszSheet9TsvPath,
    )

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

    return 0


if __name__ == "__main__":
    iExitCode: int = main()
    raise SystemExit(iExitCode)
