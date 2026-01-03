###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  make_sheet6_from_sheet4.py
#  （Sheet4.tsv と Sheet4_staff_code_range.tsv から Sheet6.tsv を作成する）
#
#  目的:
#    ・Sheet4.tsv（ジョブカン工数明細）と
#      Sheet4_staff_code_range.tsv（スタッフ毎の行範囲）を読み込み、
#      「スタッフごとに使用しているプロジェクト名一覧」を横方向に並べた
#      Sheet6.tsv を作成する。
#
#  前提:
#    ・Sheet4.tsv は、Excel の Sheet4 を .tsv 化したもの。
#      ヘッダー行あり、区切りはタブ、UTF-8。
#      少なくとも次の列が存在すること:
#        - "スタッフコード"
#        - "プロジェクト名"
#
#    ・Sheet4_staff_code_range.tsv は、Excel 上の
#      「スタッフコード」「開始行」「終了行」の 3 列を
#      そのまま .tsv 化したもの。
#
#      A列: スタッフコード
#      B列: 開始行（Excel の行番号。=MATCH(A2, Sheet4!B:B, 0) の結果など）
#      C列: 終了行（次のスタッフの開始行-1 等。ROW() 由来の Excel 行番号）
#
#      ※開始行・終了行は「Excel の行番号（1 始まり・ヘッダーを含む）」とする。
#        Python 側では、ヘッダーを除いた DataFrame の行番号に変換して用いる。
#
#  出力:
#    ・Sheet6.tsv（同じフォルダに作成）
#      1行目: 1,2,3,...（スタッフ列の通し番号）
#      2行目: スタッフコード（Sheet4_staff_code_range.tsv の 1 列目の値）
#      3行目以降: 各列が 1 人のスタッフを表し、そのスタッフの
#                  「プロジェクト名」一覧（昇順ソート＋重複削除済み）を縦方向に並べる。
#
#  実行例:
#    python make_sheet6_from_sheet4.py Sheet4.tsv Sheet4_staff_code_range.tsv
#
# ===============================================================

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
