###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
#
# convert_yyyy_mm_dd.py
#
# 目的:
#   入力された TSV ファイル内に含まれる日付文字列について、
#     yyyy/m/d
#     yyyy/mm/d
#     yyyy/m/dd
#   のように「月または日が 1 桁の場合が混在している形式」を、
#
#       yyyy/mm/dd
#
#   形式へ統一する。
#
#   例:
#       2025/9/1   → 2025/09/01
#       2025/09/1  → 2025/09/01
#       2025/9/23  → 2025/09/23
#
#   対象は 1/1 ～ 12/31 の日付を想定し、
#   不正な日付 (2025/13/40 など) は変換せず元の文字列を残す。
#

import os
import sys
import re
from typing import List

import pandas as pd
from pandas import DataFrame


# ///////////////////////////////////////////////////////////////
# write_error_tsv
# 指定パスにエラーメッセージだけを書いた TSV を作成する。
# ///////////////////////////////////////////////////////////////
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


# ///////////////////////////////////////////////////////////////
# build_output_file_full_path
# 入力 TSV ファイルのパスから出力 TSV ファイルのパスを構築する。
# ///////////////////////////////////////////////////////////////
def build_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryName: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)

    pszBaseWithoutExt: str
    pszExt: str
    pszBaseWithoutExt, pszExt = os.path.splitext(pszBaseName)

    if pszExt == "":
        pszExt = ".tsv"

    pszOutputFileName: str = pszBaseWithoutExt + "_yyyy_mm_dd" + pszExt
    pszOutputFileFullPath: str = os.path.join(pszDirectoryName, pszOutputFileName)

    return pszOutputFileFullPath


# ///////////////////////////////////////////////////////////////
# normalize_yyyy_mm_dd_in_value
# 単一セルの文字列を yyyy/mm/dd 形式に正規化する。
# ///////////////////////////////////////////////////////////////
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


# ///////////////////////////////////////////////////////////////
# normalize_yyyy_mm_dd_in_dataframe
# DataFrame 全体に対して日付の正規化を行う。
#
# applymap → FutureWarning への対応として、
# 可能なら DataFrame.map を使用し、
# 古い pandas では applymap にフォールバックする。
# ///////////////////////////////////////////////////////////////
def normalize_yyyy_mm_dd_in_dataframe(
    objDataFrameInput: DataFrame,
) -> DataFrame:
    objPattern: re.Pattern = re.compile(r"^\s*(\d{4})/(\d{1,2})/(\d{1,2})\s*$")

    def _normalize_wrapper(objValue: object) -> object:
        return normalize_yyyy_mm_dd_in_value(objValue, objPattern)

    # 新しい pandas では DataFrame.map が推奨
    try:
        objDataFrameOutput: DataFrame = objDataFrameInput.map(_normalize_wrapper)
    except AttributeError:
        # map が存在しない（古い pandas）の場合
        objDataFrameOutput = objDataFrameInput.applymap(_normalize_wrapper)

    return objDataFrameOutput


# ///////////////////////////////////////////////////////////////
# make_normalized_tsv_file
# TSV を読み込み、日付形式を正規化して出力。
# ///////////////////////////////////////////////////////////////
def make_normalized_tsv_file(
    pszInputFileFullPath: str,
) -> None:
    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # TSV 読み込み
    try:
        objDataFrameInput: DataFrame = pd.read_csv(
            pszInputFileFullPath,
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
        write_error_tsv(pszOutputFileFullPath, pszErrorMessage)
        return

    # 正規化処理
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
        write_error_tsv(pszOutputFileFullPath, pszErrorMessage)
        return

    # TSV 書き出し
    try:
        objDataFrameOutput.to_csv(
            pszOutputFileFullPath,
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
        write_error_tsv(pszOutputFileFullPath, pszErrorMessage)
        return


# ///////////////////////////////////////////////////////////////
# main
# コマンドライン引数チェックとファイル存在チェック
# ///////////////////////////////////////////////////////////////
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足
    if iArgCount < 2:
        pszLine1: str = (
            "Error: input TSV file path is not specified (insufficient arguments)."
        )
        pszLine2: str = (
            "Usage: python convert_yyyy_mm_dd.py <input_tsv_file_path>"
        )
        pszLine3: str = (
            "Example: python convert_yyyy_mm_dd.py C:\\Data\\manhour_2025.tsv"
        )

        print(pszLine1)
        print(pszLine2)
        print(pszLine3)

        pszErrorTsvPath: str = os.path.join(
            os.getcwd(),
            "convert_yyyy_mm_dd_error_argument.tsv",
        )
        pszErrorMessage: str = "\n".join([pszLine1, pszLine2, pszLine3])

        write_error_tsv(pszErrorTsvPath, pszErrorMessage)
        return

    pszInputFileFullPath: str = sys.argv[1]

    # 入力ファイル存在チェック
    if not os.path.exists(pszInputFileFullPath):
        pszDirectoryName: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)

        pszBaseWithoutExt: str
        pszExt: str
        pszBaseWithoutExt, pszExt = os.path.splitext(pszBaseName)

        pszErrorFileName: str = pszBaseWithoutExt + "_error.tsv"
        pszErrorFileFullPath: str = os.path.join(pszDirectoryName, pszErrorFileFullPath)

        pszErrorMessage: str = (
            "Error: input TSV file not found. Path = " + pszInputFileFullPath
        )

        print(pszErrorMessage)
        write_error_tsv(pszErrorFileFullPath, pszErrorMessage)
        return

    # 正常処理
    make_normalized_tsv_file(pszInputFileFullPath)


if __name__ == "__main__":
    main()
