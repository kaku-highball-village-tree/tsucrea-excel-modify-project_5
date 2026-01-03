###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  make_unique_staff_code_list.py
#  （Sheet1_yyyy_mm_dd.tsv から B列のスタッフコード一覧を作成する）
# ===============================================================

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
