###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  sort_manhour_by_staff_code.py
#  （manhour_*.tsv をスタッフコード（第2列）でソートする）
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
# 出力ファイルパスを構築（拡張子 .tsv → _sorted_staff_code.tsv）
# ---------------------------------------------------------------
def build_output_file_full_path(pszInputFileFullPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszRoot, pszExt = os.path.splitext(pszBaseName)

    pszOutputBase: str = pszRoot + "_sorted_staff_code.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBase
    return os.path.join(pszDirectory, pszOutputBase)


# ---------------------------------------------------------------
# メイン処理（manhour_*.tsv → *_sorted_staff_code.tsv）
# ---------------------------------------------------------------
def make_sorted_staff_code_tsv_from_manhour_tsv(
    pszInputFileFullPath: str,
) -> None:

    # 入力ファイル存在チェック
    if not os.path.isfile(pszInputFileFullPath):
        pszErrorFile: str = pszInputFileFullPath.replace(".tsv", "_error.tsv")
        write_error_tsv(
            pszErrorFile,
            "Error: input TSV file not found. Path = {0}".format(pszInputFileFullPath)
        )
        return

    # 出力ファイルパスを構築
    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep="\t",
            dtype=str,
            encoding="utf-8",
            engine="python"
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while reading manhour TSV for staff code sort. Detail = {0}".format(objException)
        )
        return

    # 列数チェック（2列目が必要）
    iColumnCount: int = objDataFrame.shape[1]
    if iColumnCount < 2:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: staff code column (2nd column) does not exist. ColumnCount = {0}".format(iColumnCount)
        )
        return

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
            errors="coerce"
        )

        # 一時列をキーに安定ソート（数値順）
        objSorted = objSorted.sort_values(
            by="__sort_staff_code__",
            ascending=True,
            kind="mergesort"
        ).drop(columns=["__sort_staff_code__"])

    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while sorting by staff code. Detail = {0}".format(objException)
        )
        return

    # TSV 書き込み
    try:
        objSorted.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
            lineterminator="\n"
        )
    except Exception as objException:
        write_error_tsv(
            pszOutputFileFullPath,
            "Error: unexpected exception while writing sorted staff-code TSV. Detail = {0}".format(objException)
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:

    iArgCount: int = len(sys.argv)

    # 引数不足 → _error_argument.tsv を出力
    if iArgCount < 2:
        pszProgram: str = os.path.basename(sys.argv[0])
        pszErrorFile: str = "sort_manhour_by_staff_code_error_argument.tsv"

        write_error_tsv(
            pszErrorFile,
            "Error: input TSV file path is not specified (insufficient arguments).\n"
            "Usage: python {0} <input_tsv_file_path>\n"
            "Example: python {0} C:\\Data\\manhour_202511181454691c0a3179197.tsv".format(pszProgram)
        )
        return

    pszInputFileFullPath: str = sys.argv[1]

    make_sorted_staff_code_tsv_from_manhour_tsv(pszInputFileFullPath)


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
