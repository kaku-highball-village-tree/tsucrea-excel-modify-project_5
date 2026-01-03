###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  manhour_remove_uninput_rows.py
#  （G,H,I,J 列に「未入力」が含まれる行を削除する）
# ===============================================================

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
