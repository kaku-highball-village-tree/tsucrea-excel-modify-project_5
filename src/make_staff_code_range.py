###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  make_staff_code_range.py
#  （Sheet1.tsv からスタッフコードごとの行範囲一覧を作成する）
# ===============================================================

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
