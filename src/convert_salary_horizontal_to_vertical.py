# -*- coding: utf-8 -*-
# ===============================================================
#  convert_salary_horizontal_to_vertical.py
#  支給・控除等一覧表_(給与) の横持ち CSV を縦持ちの一覧表に変換するスクリプト
# ===============================================================

import sys
import os
from typing import List, Tuple

import pandas as pd
from pandas import DataFrame
from pandas.api.types import is_float_dtype


# ///////////////////////////////////////////////////////////////
# //
# // エラーメッセージをテキストファイルに書き込む関数。
# // 既存の convert_* 系スクリプトと同じ形式で出力する。
# //
# ///////////////////////////////////////////////////////////////
def write_error_text(
    pszErrorMessage: str, 
    pszErrorFileFullPath: str, 
) -> None:
    pszDirectoryName: str = os.path.dirname(pszErrorFileFullPath)
    if pszDirectoryName and not os.path.exists(pszDirectoryName):
        os.makedirs(pszDirectoryName, exist_ok=True)

    with open(pszErrorFileFullPath, mode="w", encoding="utf-8") as objFile:
        objFile.write(pszErrorMessage)


# ///////////////////////////////////////////////////////////////
# //
# // 入力 CSV ファイルパスから、出力用のベースパスを作成する関数。
# // 例: input.csv -> input_vertical.csv / input_vertical.tsv / input_vertical.xlsx
# //
# ///////////////////////////////////////////////////////////////
def build_output_base_path(
    pszInputFileFullPath: str, 
) -> str:
    pszDirectoryName: str = os.path.dirname(pszInputFileFullPath)
    pszFileName: str = os.path.basename(pszInputFileFullPath)
    pszFileNameWithoutExt: str
    pszExt: str
    pszFileNameWithoutExt, pszExt = os.path.splitext(pszFileName)
    pszOutputFileNameBase: str = pszFileNameWithoutExt + "_vertical"
    pszOutputBasePath: str = os.path.join(pszDirectoryName, pszOutputFileNameBase)
    return pszOutputBasePath


# ///////////////////////////////////////////////////////////////
# //
# // DataFrame 内の float 型列を四捨五入して整数に変換する関数。
# // 「従業員名」列のような文字列列はそのまま残す。
# //
# ///////////////////////////////////////////////////////////////
def convert_decimal_columns_to_integer(
    objDataFrameSource: DataFrame, 
) -> DataFrame:
    objDataFrameResult: DataFrame = objDataFrameSource.copy()
    for pszColumnName in objDataFrameResult.columns:
        if pszColumnName == "従業員名":
            continue
        if is_float_dtype(objDataFrameResult[pszColumnName]):
            objDataFrameResult[pszColumnName] = (
                objDataFrameResult[pszColumnName].round().astype("Int64")
            )
    return objDataFrameResult


# ///////////////////////////////////////////////////////////////
# //
# // 給与合計と給与合計(旧)の数式列を追加する関数。
# // 各行の Excel 行番号に合わせて、SUM 最適化版と旧版の式文字列を生成する。
# //
# ///////////////////////////////////////////////////////////////
def add_salary_total_formula_columns(
    objDataFrameVertical: DataFrame, 
) -> DataFrame:
    objDataFrameResult: DataFrame = objDataFrameVertical.copy()
    iRowCount: int = len(objDataFrameResult)

    objListNewTotalFormula: List[str] = []
    objListOldTotalFormula: List[str] = []

    for iRowIndex in range(iRowCount):
        iExcelRowNumber: int = iRowIndex + 2
        pszRowNumberText: str = str(iExcelRowNumber)

        pszNewTotalFormula: str = (
            f"=SUM(C{pszRowNumberText}:N{pszRowNumberText},Q{pszRowNumberText}:R{pszRowNumberText})"
            f" - SUM(O{pszRowNumberText},P{pszRowNumberText},S{pszRowNumberText})"
            f" - E{pszRowNumberText} - Q{pszRowNumberText}"
        )

        pszOldTotalFormula: str = (
            f"=C{pszRowNumberText}+J{pszRowNumberText}+F{pszRowNumberText}+H{pszRowNumberText}"
            f"+K{pszRowNumberText}+N{pszRowNumberText}+I{pszRowNumberText}"
            f"-S{pszRowNumberText}-O{pszRowNumberText}-P{pszRowNumberText}"
            f"+L{pszRowNumberText}+R{pszRowNumberText}+M{pszRowNumberText}+G{pszRowNumberText}"
        )

        objListNewTotalFormula.append(pszNewTotalFormula)
        objListOldTotalFormula.append(pszOldTotalFormula)

    objDataFrameResult["給与合計"] = objListNewTotalFormula
    objDataFrameResult["給与合計(旧)"] = objListOldTotalFormula

    return objDataFrameResult


# ///////////////////////////////////////////////////////////////
# //
# // 支給・控除等一覧表_(給与) の横持ち CSV を読み込み、
# // 従業員単位の縦持ち DataFrame に変換する関数。
# //
# ///////////////////////////////////////////////////////////////
def convert_salary_horizontal_to_vertical(
    pszInputFileFullPath: str, 
) -> DataFrame:
    # CSV を読み込む
    objDataFrameInput: DataFrame = pd.read_csv(
        pszInputFileFullPath, 
        encoding="utf-8-sig", 
    )

    # 先頭列の列名 (従業員名) を取得
    pszFirstColumnName: str = objDataFrameInput.columns[0]

    # 先頭列をインデックスにして、「項目名」行を列にする (横持ち -> 縦持ち準備)
    objDataFrameIndexed: DataFrame = objDataFrameInput.set_index(pszFirstColumnName)

    # 転置して、行:従業員 / 列:項目名 の形式にする
    objDataFrameTransposed: DataFrame = objDataFrameIndexed.T

    # インデックス (従業員名) を列として追加する
    objDataFrameTransposed = objDataFrameTransposed.copy()
    objDataFrameTransposed.insert(0, "従業員名", objDataFrameTransposed.index)

    # 列順を「従業員名」「スタッフコード」その他の項目 の順に並べ替える
    objListAllColumns: List[str] = list(objDataFrameTransposed.columns)
    objListColumnOrder: List[str] = []
    if "従業員名" in objListAllColumns:
        objListColumnOrder.append("従業員名")
    if "スタッフコード" in objListAllColumns:
        objListColumnOrder.append("スタッフコード")
    for pszColumnName in objListAllColumns:
        if pszColumnName not in objListColumnOrder:
            objListColumnOrder.append(pszColumnName)

    objDataFrameVertical: DataFrame = objDataFrameTransposed[objListColumnOrder]

    # インデックスは単純な連番にリセットしておく
    objDataFrameVertical = objDataFrameVertical.reset_index(drop=True)

    # 小数を含む数値列を整数に変換する
    objDataFrameVertical = convert_decimal_columns_to_integer(
        objDataFrameVertical, 
    )

    # 給与合計と給与合計(旧)の数式列を末尾に追加する
    objDataFrameVertical = add_salary_total_formula_columns(
        objDataFrameVertical, 
    )

    return objDataFrameVertical


# ///////////////////////////////////////////////////////////////
# //
# // 変換結果の DataFrame を CSV / TSV / Excel(xlsx) に出力する関数。
# //
# ///////////////////////////////////////////////////////////////
def save_vertical_salary_files(
    objDataFrameVertical: DataFrame, 
    pszOutputBasePath: str, 
) -> Tuple[str, str, str]:
    pszCsvPath: str = pszOutputBasePath + ".csv"
    pszTsvPath: str = pszOutputBasePath + ".tsv"
    pszExcelPath: str = pszOutputBasePath + ".xlsx"

    # CSV (カンマ区切り)
    objDataFrameVertical.to_csv(
        pszCsvPath, 
        index=False, 
        encoding="utf-8-sig", 
    )

    # TSV (タブ区切り)
    objDataFrameVertical.to_csv(
        pszTsvPath, 
        sep="\t", 
        index=False, 
        encoding="utf-8-sig", 
    )

    # Excel
    objDataFrameVertical.to_excel(
        pszExcelPath, 
        index=False, 
    )

    return pszCsvPath, pszTsvPath, pszExcelPath


# ///////////////////////////////////////////////////////////////
# //
# // メイン処理。
# // コマンドライン引数で指定された CSV を変換し、
# // CSV / TSV / Excel を同じフォルダに出力する。
# //
# ///////////////////////////////////////////////////////////////
def main() -> int:
    iArgc: int = len(sys.argv)
    if iArgc != 2:
        pszScriptName: str = os.path.basename(sys.argv[0])
        pszMessage: str = (
            "Usage: python " + pszScriptName + " <input_csv_path>\n"
            "  支給・控除等一覧表_(給与) の横持ち CSV を指定してください。\n"
        )
        sys.stderr.write(pszMessage)
        return 1

    pszInputFileFullPath: str = sys.argv[1]

    if not os.path.exists(pszInputFileFullPath):
        pszMessage: str = "入力ファイルが見つかりません: " + pszInputFileFullPath + "\n"
        sys.stderr.write(pszMessage)
        return 1

    pszOutputBasePath: str = build_output_base_path(pszInputFileFullPath)
    pszErrorFilePath: str = pszOutputBasePath + "_error.txt"

    try:
        objDataFrameVertical: DataFrame = convert_salary_horizontal_to_vertical(
            pszInputFileFullPath, 
        )
        pszCsvPath: str
        pszTsvPath: str
        pszExcelPath: str
        pszCsvPath, pszTsvPath, pszExcelPath = save_vertical_salary_files(
            objDataFrameVertical, 
            pszOutputBasePath, 
        )

        pszMessage: str = (
            "変換が完了しました。\n"
            "  CSV : " + pszCsvPath + "\n"
            "  TSV : " + pszTsvPath + "\n"
            "  Excel : " + pszExcelPath + "\n"
        )
        sys.stdout.write(pszMessage)
        return 0

    except Exception as objException:
        pszMessage: str = "エラーが発生しました: " + str(objException) + "\n"
        write_error_text(
            pszMessage, 
            pszErrorFilePath, 
        )
        sys.stderr.write(pszMessage)
        return 1


if __name__ == "__main__":
    iExitCode: int = main()
    sys.exit(iExitCode)
