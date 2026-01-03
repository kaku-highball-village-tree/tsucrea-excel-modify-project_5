###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
#
# csv_to_tsv_h_mm_ss.py
#
# 目的:
#   カンマ区切りの CSV ファイルを読み込み、
#   同じ内容をタブ区切りの TSV ファイルとして出力する。
#   データの中身は一切変換せず、「区切り文字だけ」を
#   カンマ(,) からタブ(\t) に変更することを目的とする。
#
#   特に、F列「総労働時間」、K列「工数」に含まれる
#   '7:30', '1:30' などの値は、時刻型として扱わず、
#   文字列として読み込んで、そのまま文字列として TSV に出力する。
#
#   ※ 本バージョンでは、F列・K列の値が「h:mm」形式の場合、
#      自動的に「h:mm:00」を付与する (例: 7:30 → 7:30:00)。
#
# 仕様(入力と出力のファイル名の規則):
#   コマンドライン引数で入力 CSV ファイルのパスを 1 件だけ受け取る:
#     例:
#       python csv_to_tsv_h_mm_ss.py C:\Data\manhour_2025.csv
#
#   入力ファイルパスから、同じディレクトリ内に拡張子 .tsv の出力ファイルを作成する:
#     - 入力: C:\Data\manhour_2025.csv
#     - 出力: C:\Data\manhour_2025.tsv
#
#   拡張子の変換には os.path.splitext を用い、
#   ベース名とディレクトリ構成はそのまま利用する。
#

import os
import sys
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
# 入力 CSV ファイルのパスから、出力 TSV ファイルのパスを構築する。
# ///////////////////////////////////////////////////////////////
def build_output_file_full_path(
    pszInputFileFullPath: str,
) -> str:
    pszDirectoryName: str = os.path.dirname(pszInputFileFullPath)
    pszBaseName: str = os.path.basename(pszInputFileFullPath)
    pszBaseWithoutExt: str
    pszExt: str

    pszBaseWithoutExt, pszExt = os.path.splitext(pszBaseName)
    pszOutputFileName: str = pszBaseWithoutExt + ".tsv"
    pszOutputFileFullPath: str = os.path.join(pszDirectoryName, pszOutputFileName)

    return pszOutputFileFullPath


# ///////////////////////////////////////////////////////////////
# convert_csv_to_tsv_file
# 単一の CSV ファイルを読み込み、TSV ファイルに変換する。
# ///////////////////////////////////////////////////////////////
def convert_csv_to_tsv_file(
    pszInputFileFullPath: str,
) -> None:
    pszOutputFileFullPath: str = build_output_file_full_path(pszInputFileFullPath)

    # CSV 読み込み
    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFileFullPath,
            sep=",",
            header=0,
            dtype=str,
            encoding="utf-8",
            engine="python",
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while reading CSV. Detail = "
            + str(objException)
        )
        print(pszErrorMessage)
        write_error_tsv(pszOutputFileFullPath, pszErrorMessage)
        return

    # -------------------------------------------------------------
    # ▼▼▼（ここから追加）F列・K列の h:mm → h:mm:00 加工 ▼▼▼
    # -------------------------------------------------------------
    for pszCol in ["総労働時間", "工数"]:
        if pszCol in objDataFrame.columns:
            objDataFrame[pszCol] = objDataFrame[pszCol].apply(
                lambda v: (
                    v + ":00"
                    if isinstance(v, str)
                    and v.count(":") == 1       # 例: "7:30"
                    and len(v.split(":")[1]) <= 2  # "30" のような 2桁以下
                    else v
                )
            )
    # -------------------------------------------------------------
    # ▲▲▲（ここまで追加）▲▲▲
    # -------------------------------------------------------------

    # TSV 書き出し
    try:
        objDataFrame.to_csv(
            pszOutputFileFullPath,
            sep="\t",
            index=False,
            encoding="utf-8",
        )
    except Exception as objException:
        pszErrorMessage: str = (
            "Error: unexpected exception while writing TSV. Detail = "
            + str(objException)
        )
        print(pszErrorMessage)
        write_error_tsv(pszOutputFileFullPath, pszErrorMessage)
        return


# ///////////////////////////////////////////////////////////////
# main
# ///////////////////////////////////////////////////////////////
def main() -> None:
    iArgCount: int = len(sys.argv)

    if iArgCount < 2:
        pszLine1: str = "Error: input CSV file path is not specified (insufficient arguments)."
        pszLine2: str = "Usage: python csv_to_tsv_h_mm_ss.py <input_csv_file_path>"
        pszLine3: str = "Example: python csv_to_tsv_h_mm_ss.py C:\\Data\\manhour_2025.csv"

        print(pszLine1)
        print(pszLine2)
        print(pszLine3)

        pszErrorTsvPath: str = os.path.join(
            os.getcwd(),
            "csv_to_tsv_error_argument.tsv",
        )
        pszErrorMessage: str = "\n".join([pszLine1, pszLine2, pszLine3])

        write_error_tsv(pszErrorTsvPath, pszErrorMessage)
        return

    pszInputFileFullPath: str = sys.argv[1]

    if not os.path.exists(pszInputFileFullPath):
        pszDirectoryName: str = os.path.dirname(pszInputFileFullPath)
        pszBaseName: str = os.path.basename(pszInputFileFullPath)
        pszBaseWithoutExt: str
        pszExt: str

        pszBaseWithoutExt, pszExt = os.path.splitext(pszBaseName)
        pszErrorFileName: str = pszBaseWithoutExt + "_error.tsv"
        pszErrorFileFullPath: str = os.path.join(pszDirectoryName, pszErrorFileName)

        pszErrorMessage: str = (
            "Error: input CSV file not found. Path = " + pszInputFileFullPath
        )
        print(pszErrorMessage)
        write_error_tsv(pszErrorFileFullPath, pszErrorMessage)
        return

    convert_csv_to_tsv_file(pszInputFileFullPath)


if __name__ == "__main__":
    main()
