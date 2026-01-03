import os
import sys
from typing import Any

import pandas as pd
from pandas import DataFrame


def b_is_blank_value(objValue: Any) -> bool:
    """Check whether the provided value should be treated as blank."""
    bIsNaN: bool = False
    try:
        bIsNaN = bool(pd.isna(objValue))
    except Exception:
        bIsNaN = False
    if bIsNaN:
        return True
    if objValue is None:
        return True
    if isinstance(objValue, str):
        if objValue.strip() == "":
            return True
    return False


def main() -> None:
    pszInputFilePath: str = os.path.join("input", "Raw_Data.tsv")
    pszOutputFilePath: str = os.path.join("input", "Raw_Data_remove_blank_rows.tsv")

    if not os.path.isfile(pszInputFilePath):
        print("入力ファイルが見つかりません: Raw_Data.tsv")
        sys.exit(1)

    try:
        objDataFrame: DataFrame = pd.read_csv(
            pszInputFilePath,
            sep="\t",
            dtype=str,
            keep_default_na=True,
        )
    except Exception as objException:
        print(f"読み込みに失敗しました: {objException}")
        sys.exit(1)

    if "スタッフコード" not in objDataFrame.columns:
        print("列が見つかりません: スタッフコード")
        sys.exit(1)
    if "処理関数1(スタッフ名)" not in objDataFrame.columns:
        print("列が見つかりません: 処理関数1(スタッフ名)")
        sys.exit(1)

    iCutIndex: int = -1
    for iRowIndex in range(len(objDataFrame)):
        objRow: Any = objDataFrame.iloc[iRowIndex]
        objStaffCode: Any = objRow["スタッフコード"]
        objStaffName: Any = objRow["処理関数1(スタッフ名)"]

        bStaffCodeBlank: bool = b_is_blank_value(objStaffCode)
        bStaffNameBlank: bool = b_is_blank_value(objStaffName)

        if bStaffCodeBlank and bStaffNameBlank:
            iCutIndex = iRowIndex
            break

    if iCutIndex != -1:
        objDataFrame = objDataFrame.iloc[:iCutIndex]

    try:
        objDataFrame.to_csv(pszOutputFilePath, sep="\t", index=False)
    except Exception as objException:
        print(f"書き込みに失敗しました: {objException}")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as objException:
        print(f"エラーが発生しました: {objException}")
        sys.exit(1)
