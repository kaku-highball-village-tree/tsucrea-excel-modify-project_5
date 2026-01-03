###############################################################
# Reference-only file
# This file is provided as a logic reference.
# DO NOT MODIFY this file.
# Copy logic manually into make_manhour_to_sheet8_01_0001.py
# if needed.
###############################################################

# -*- coding: utf-8 -*-
# ===============================================================
#  make_sheet789_from_sheet4.py
#  （Sheet4.tsv / Sheet4_staff_code_range.tsv / Sheet6.tsv から
#    Sheet7.tsv と Sheet8.tsv と Sheet9.tsv を作成する・Sheet6主体アルゴリズム版）
#
#  目的:
#    ・Sheet6.tsv を「マスタ」として用い、
#      - Sheet6 の 2 行目に並ぶスタッフコードの順番をそのまま採用し、
#      - 各列の 3 行目以降に並ぶプロジェクト名一覧をそのまま利用して、
#      Sheet7.tsv / Sheet8.tsv を作成する。
#
#    ・工数の実データは Sheet4.tsv から読み取り、
#      スタッフコードとプロジェクト名毎に工数を合計する。
#
#    ・Sheet9.tsv には「氏名・スタッフコード」の一覧を重複なくもれなく出力する。
#      1 行目はヘッダー行（氏名, スタッフコード）とする。
#
#  各シートのイメージ:
#
#    Sheet6.tsv（make_sheet6_from_sheet4.py の出力）
#      1 行目: 1, 2, 3, ...                   ← スタッフ通し番号
#      2 行目: スタッフコード(1列目), スタッフコード(2列目), ...
#      3 行目以降: 各列のスタッフが関わったプロジェクト名一覧（縦方向）
#
#    Sheet7.tsv（本スクリプトの出力）
#      1 列目: プロジェクト名（Sheet6 の A,B,C,... 列を縦に並べたもの）
#      2 列目: スタッフコード
#               └ 各スタッフの行すべてに同じスタッフコードを入れる
#      3 列目: 合計工数時間（hh:mm:ss 形式、24 時間超も総時間数）
#
#    Sheet8.tsv（本スクリプトの出力）
#      1 列目: 氏名
#               └ 各スタッフの「最初の行」にだけ値を入れ、それ以外は空文字
#      2 列目: プロジェクト名（Sheet7.tsv の 1 列目と同じ）
#      3 列目: スタッフコード（Sheet7.tsv の 2 列目と同じ。全行に値あり）
#      4 列目: 合計工数時間（Sheet7.tsv の 3 列目と同じ）
#
#    ※Sheet8.tsv の 2～4 列目は、Sheet7.tsv の 1～3 列目と同じ内容になる。
#
#    Sheet9.tsv（本スクリプトの出力）
#      1 行目: 氏名, スタッフコード（ヘッダー）
#      2 行目以降:
#        ・氏名, スタッフコード
#        ・同じスタッフコードは 1 回だけ出力する（重複なし）。
#        ・Sheet6 のスタッフコード一覧と Sheet4 の「スタッフコード→氏名」の
#          両方を元に、可能な限り「もれなく」一覧化する。
#
#  前提:
#    ・Sheet4.tsv は、Excel の Sheet4 を .tsv 化したもの。
#      ヘッダー行あり、区切りはタブ。
#      エンコーディングは UTF-8(BOM付き) または cp932(Shift-JIS) のいずれか。
#      少なくとも次の列が存在すること:
#        - "スタッフコード"
#        - "プロジェクト名"
#        - "工数"
#        - 氏名列（"姓 名" または "氏名" のいずれか）
#
#    ・Sheet4_staff_code_range.tsv は、Excel 上の
#      「スタッフコード」「開始行」「終了行」の 3 列を
#      そのまま .tsv 化したもの。
#
#      A列: スタッフコード
#      B列: 開始行（Excel 行番号。=MATCH(A2, Sheet4!B:B, 0) の結果など）
#      C列: 終了行（次のスタッフの開始行-1 等。ROW() 由来の Excel 行番号）
#
#      ※開始行・終了行は「Excel の行番号（1 始まり・ヘッダーを含む）」とする。
#        Python 側では、ヘッダーを除いた DataFrame の行番号に変換して用いる。
#
#    ・Sheet6.tsv は、make_sheet6_from_sheet4.py によって作成された形式
#      （1 行目: 通し番号, 2 行目: スタッフコード, 3 行目以降: プロジェクト名一覧）であること。
#
#  実行例:
#    python make_sheet789_from_sheet4.py Sheet4.tsv Sheet4_staff_code_range.tsv Sheet6.tsv
#
#  出力:
#    ・Sheet7.tsv（Sheet4.tsv と同じフォルダ）
#    ・Sheet8.tsv（Sheet4.tsv と同じフォルダ）
#    ・Sheet9.tsv（Sheet4.tsv と同じフォルダ）
#
# ===============================================================

import os
import sys
from typing import Dict, List, Tuple, Set

import pandas as pd
from pandas import DataFrame


# ///////////////////////////////////////////////////////////////
# //
# //  TSV を UTF-8(BOM付き) 優先・cp932 併用で読み込む関数。
# //  bHasHeader=True のときは 1 行目をヘッダーとして扱う。
# //  bHasHeader=False のときはヘッダー無し（header=None）で読み込む。
# //
# ///////////////////////////////////////////////////////////////
def read_tsv_with_encoding_candidates(
    pszInputFileFullPath: str, 
    bHasHeader: bool, 
) -> DataFrame:
    objEncodingCandidateList: List[str] = [
        "utf-8-sig", 
        "cp932", 
    ]
    bLoaded: bool = False
    objLastDecodeError: Exception | None = None
    objDataFrameResult: DataFrame | None = None

    for pszEncoding in objEncodingCandidateList:
        try:
            if bHasHeader:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath, 
                    sep="\t", 
                    dtype=str, 
                    encoding=pszEncoding, 
                    header=0, 
                    engine="python", 
                )
            else:
                objDataFrameResult = pd.read_csv(
                    pszInputFileFullPath, 
                    sep="\t", 
                    dtype=str, 
                    encoding=pszEncoding, 
                    header=None, 
                    engine="python", 
                )
            bLoaded = True
            break
        except UnicodeDecodeError as objDecodeError:
            objLastDecodeError = objDecodeError
            continue

    if not bLoaded or objDataFrameResult is None:
        if objLastDecodeError is not None:
            raise objLastDecodeError
        raise UnicodeDecodeError(
            "utf-8-sig", 
            b"", 
            0, 
            1, 
            "cannot decode TSV with utf-8-sig nor cp932", 
        )

    return objDataFrameResult


# ---------------------------------------------------------------
# 指定したパスにエラーTSVを書き込む（既存スタイル踏襲）
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
# Sheet7.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet7(
    pszSheet4FileFullPath: str, 
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet7.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# Sheet8.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet8(
    pszSheet4FileFullPath: str, 
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet8.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ---------------------------------------------------------------
# Sheet9.tsv の出力ファイルパスを構築
# ---------------------------------------------------------------
def build_output_file_full_path_for_sheet9(
    pszSheet4FileFullPath: str, 
) -> str:
    pszDirectory: str = os.path.dirname(pszSheet4FileFullPath)
    pszOutputBaseName: str = "Sheet9.tsv"
    if len(pszDirectory) == 0:
        return pszOutputBaseName
    return os.path.join(pszDirectory, pszOutputBaseName)


# ///////////////////////////////////////////////////////////////
# //
# //  文字列の時間（hh:mm:ss または hh:mm）を「秒数(int)」に変換する関数。
# //  不正な形式や空文字の場合は 0 秒として扱う。
# //
# ///////////////////////////////////////////////////////////////
def convert_time_string_to_seconds(
    pszTimeText: str, 
) -> int:
    if pszTimeText is None:
        return 0

    pszWork: str = str(pszTimeText).strip()
    if len(pszWork) == 0:
        return 0

    objParts: List[str] = pszWork.split(":")
    try:
        if len(objParts) == 3:
            iHour: int = int(objParts[0])
            iMinute: int = int(objParts[1])
            iSecond: int = int(objParts[2])
        elif len(objParts) == 2:
            iHour: int = int(objParts[0])
            iMinute: int = int(objParts[1])
            iSecond: int = 0
        else:
            return 0
    except ValueError:
        return 0

    iTotalSeconds: int = iHour * 3600 + iMinute * 60 + iSecond
    return iTotalSeconds


# ///////////////////////////////////////////////////////////////
# //
# //  秒数(int)を「hh:mm:ss」形式の文字列に変換する関数。
# //  24 時間を超えても、そのまま総時間数として出力する。
# //
# ///////////////////////////////////////////////////////////////
def convert_seconds_to_time_string(
    iTotalSeconds: int, 
) -> str:
    if iTotalSeconds <= 0:
        return "0:00:00"

    iHour: int = iTotalSeconds // 3600
    iRemain: int = iTotalSeconds % 3600
    iMinute: int = iRemain // 60
    iSecond: int = iRemain % 60

    pszTimeText: str = "{0}:{1:02d}:{2:02d}".format(iHour, iMinute, iSecond)
    return pszTimeText


# ---------------------------------------------------------------
# Sheet4.tsv, Sheet4_staff_code_range.tsv, Sheet6.tsv から
# Sheet7.tsv と Sheet8.tsv と Sheet9.tsv を作成する（Sheet6 主体アルゴリズム）
# ---------------------------------------------------------------
def make_sheet789_from_sheet4(
    pszSheet4FileFullPath: str, 
    pszRangeFileFullPath: str, 
    pszSheet6FileFullPath: str, 
) -> None:

    # 出力ファイルパス（エラー時もこのフォルダを基準にする）
    pszSheet7FileFullPath: str = build_output_file_full_path_for_sheet7(pszSheet4FileFullPath)
    pszSheet7ErrorFileFullPath: str = pszSheet7FileFullPath.replace(".tsv", "_error.tsv")
    pszSheet8FileFullPath: str = build_output_file_full_path_for_sheet8(pszSheet4FileFullPath)
    pszSheet8ErrorFileFullPath: str = pszSheet8FileFullPath.replace(".tsv", "_error.tsv")
    pszSheet9FileFullPath: str = build_output_file_full_path_for_sheet9(pszSheet4FileFullPath)
    pszSheet9ErrorFileFullPath: str = pszSheet9FileFullPath.replace(".tsv", "_error.tsv")

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

    # 入力ファイル存在チェック（Sheet6）
    if not os.path.isfile(pszSheet6FileFullPath):
        write_error_tsv(
            pszSheet6FileFullPath.replace(".tsv", "_error.tsv"), 
            "Error: Sheet6 TSV file not found. Path = {0}".format(pszSheet6FileFullPath), 
        )
        return

    # -----------------------------------------------------------
    # Sheet4.tsv 読み込み（UTF-8-sig → cp932 の順で試す）
    # -----------------------------------------------------------
    try:
        objDataFrameSheet4: DataFrame = read_tsv_with_encoding_candidates(
            pszSheet4FileFullPath, 
            True, 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while reading Sheet4 TSV. Detail = {0}".format(objException), 
        )
        return

    objSheet4Columns: List[str] = list(objDataFrameSheet4.columns)
    if (
        ("スタッフコード" not in objSheet4Columns) 
        or ("プロジェクト名" not in objSheet4Columns) 
        or ("工数" not in objSheet4Columns) 
    ):
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: required columns not found in Sheet4 TSV. "
            "Required columns: スタッフコード, プロジェクト名, 工数. "
            "Columns = {0}".format(", ".join(objSheet4Columns)), 
        )
        return

    # 氏名列の候補を探す（"姓 名" または "氏名"）
    pszNameColumn: str = ""
    if "姓 名" in objSheet4Columns:
        pszNameColumn = "姓 名"
    elif "氏名" in objSheet4Columns:
        pszNameColumn = "氏名"

    # 工数列を秒数に変換した補助列を追加
    try:
        objDataFrameSheet4 = objDataFrameSheet4.copy()
        objDataFrameSheet4["__time_seconds__"] = objDataFrameSheet4["工数"].apply(
            lambda pszValue: convert_time_string_to_seconds(pszValue), 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while converting time text to seconds. Detail = {0}".format(objException), 
        )
        return

    iSheet4RowCount: int = objDataFrameSheet4.shape[0]

    # -----------------------------------------------------------
    # Sheet4_staff_code_range.tsv 読み込み
    #   → スタッフコード → (開始インデックス, 終了インデックス) の辞書を作る
    # -----------------------------------------------------------
    try:
        objDataFrameRange: DataFrame = read_tsv_with_encoding_candidates(
            pszRangeFileFullPath, 
            True, 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while reading Sheet4_staff_code_range TSV. Detail = {0}".format(
                objException, 
            ), 
        )
        return

    iRangeColumnCount: int = objDataFrameRange.shape[1]
    if iRangeColumnCount < 3:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: Sheet4_staff_code_range TSV must have at least 3 columns "
            "(staff_code, start_row, end_row). ColumnCount = {0}".format(iRangeColumnCount), 
        )
        return

    try:
        objDataFrameRange = objDataFrameRange.copy()
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
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while converting start/end rows to numeric. Detail = {0}".format(
                objException, 
            ), 
        )
        return

    if objDataFrameRange["__start_row_excel__"].isna().any() or objDataFrameRange["__end_row_excel__"].isna().any():
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: start_row or end_row contains non-numeric value in Sheet4_staff_code_range TSV.", 
        )
        return

    # スタッフコード → (開始Index, 終了Index) の辞書を構築
    objDictStaffCodeToRange: Dict[str, Tuple[int, int]] = {}
    for iIndex, objRow in objDataFrameRange.iterrows():
        pszStaffCodeRange: str = str(objRow.iloc[0]).strip()
        if len(pszStaffCodeRange) == 0:
            continue

        iStartRowExcel: int = int(objRow["__start_row_excel__"])
        iEndRowExcel: int = int(objRow["__end_row_excel__"])

        iStartIndex: int = iStartRowExcel - 2  # 1 行目: ヘッダー, 2 行目: index 0
        iEndIndex: int = iEndRowExcel - 2

        if (iStartIndex < 0) or (iEndIndex < 0) or (iStartIndex > iEndIndex) or (iEndIndex >= iSheet4RowCount):
            write_error_tsv(
                pszSheet7ErrorFileFullPath, 
                "Error: invalid row range for staff code {0}. "
                "StartExcel={1}, EndExcel={2}, StartIndex={3}, EndIndex={4}, Sheet4RowCount={5}".format(
                    pszStaffCodeRange, 
                    iStartRowExcel, 
                    iEndRowExcel, 
                    iStartIndex, 
                    iEndIndex, 
                    iSheet4RowCount, 
                ), 
            )
            return

        # 同じスタッフコードが複数行に出てきた場合は最初のものを優先
        if pszStaffCodeRange not in objDictStaffCodeToRange:
            objDictStaffCodeToRange[pszStaffCodeRange] = (iStartIndex, iEndIndex)

    # -----------------------------------------------------------
    # Sheet4 から「スタッフコード → 氏名」の辞書を作る（あれば）
    # -----------------------------------------------------------
    objDictStaffCodeToName: Dict[str, str] = {}
    if pszNameColumn != "":
        for iIndex, objRow in objDataFrameSheet4.iterrows():
            pszStaffCodeFromSheet4: str = str(objRow["スタッフコード"]).strip()
            if len(pszStaffCodeFromSheet4) == 0:
                continue
            pszStaffName: str = str(objRow[pszNameColumn]).strip()
            if pszStaffCodeFromSheet4 not in objDictStaffCodeToName:
                objDictStaffCodeToName[pszStaffCodeFromSheet4] = pszStaffName

    # -----------------------------------------------------------
    # Sheet6.tsv 読み込み（スタッフコード順・プロジェクト一覧を利用）
    # -----------------------------------------------------------
    try:
        # ヘッダー行なしとして読み込む（1 行目〜をそのまま扱う）
        objDataFrameSheet6: DataFrame = read_tsv_with_encoding_candidates(
            pszSheet6FileFullPath, 
            False, 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while reading Sheet6 TSV. Detail = {0}".format(objException), 
        )
        return

    if objDataFrameSheet6.shape[0] < 2:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: Sheet6 TSV must have at least 2 rows (1st: index, 2nd: staff_code).", 
        )
        return

    iSheet6RowCount: int = objDataFrameSheet6.shape[0]
    iSheet6ColumnCount: int = objDataFrameSheet6.shape[1]

    # 2 行目(インデックス 1) のスタッフコード一覧を取得
    objListStaffCodeFromSheet6: List[str] = []
    for iColumnIndex in range(iSheet6ColumnCount):
        pszStaffCodeFromSheet6: str = ""
        objValue = objDataFrameSheet6.iat[1, iColumnIndex]
        if objValue is not None:
            pszStaffCodeFromSheet6 = str(objValue).strip()
        if len(pszStaffCodeFromSheet6) == 0:
            # 空列はスキップ
            continue
        objListStaffCodeFromSheet6.append(pszStaffCodeFromSheet6)

    # -----------------------------------------------------------
    # Sheet7 / Sheet8 / Sheet9 の出力行リストを構築
    # -----------------------------------------------------------
    objListOutputRowsSheet7: List[List[str]] = []
    objListOutputRowsSheet8: List[List[str]] = []
    objListOutputRowsSheet9: List[List[str]] = []

    # Sheet9 用: スタッフコード一覧を元に「氏名・スタッフコード」を作成
    objSetAddedStaffCodeForSheet9: Set[str] = set()

    for pszStaffCodeForSheet9 in objListStaffCodeFromSheet6:
        pszStaffNameForSheet9: str = objDictStaffCodeToName.get(pszStaffCodeForSheet9, "")
        objListOutputRowsSheet9.append(
            [pszStaffNameForSheet9, pszStaffCodeForSheet9], 
        )
        objSetAddedStaffCodeForSheet9.add(pszStaffCodeForSheet9)

    # Sheet4 由来のスタッフコードで、Sheet6 に無いものがあれば末尾に追加
    for pszStaffCodeExtra, pszStaffNameExtra in objDictStaffCodeToName.items():
        if pszStaffCodeExtra in objSetAddedStaffCodeForSheet9:
            continue
        objListOutputRowsSheet9.append(
            [pszStaffNameExtra, pszStaffCodeExtra], 
        )
        objSetAddedStaffCodeForSheet9.add(pszStaffCodeExtra)

    # スタッフコードを Sheet6 の並び順でループ（Sheet7 / Sheet8 用）
    for pszStaffCode in objListStaffCodeFromSheet6:
        # 該当スタッフの行範囲（Sheet4）を辞書から取得
        if pszStaffCode in objDictStaffCodeToRange:
            iStartIndex, iEndIndex = objDictStaffCodeToRange[pszStaffCode]
            objDataFrameSubStaff: DataFrame = objDataFrameSheet4.iloc[iStartIndex : iEndIndex + 1]
            objDataFrameSubStaff = objDataFrameSubStaff[objDataFrameSubStaff["スタッフコード"] == pszStaffCode]
        else:
            # 念のため、辞書に無い場合はSheet4全体からフィルタ
            objDataFrameSubStaff = objDataFrameSheet4[
                objDataFrameSheet4["スタッフコード"] == pszStaffCode
            ]

        if objDataFrameSubStaff.empty:
            # 該当スタッフの工数データが無ければスキップ
            continue

        pszStaffNameForSheet8: str = objDictStaffCodeToName.get(pszStaffCode, "")
        iRowIndexWithinStaff: int = 0

        # Sheet6 の列からプロジェクト名一覧を縦方向に取得
        #  1 行目: 通し番号
        #  2 行目: スタッフコード
        #  3 行目以降: プロジェクト名
        #
        # ここでは、全行を走査し、指定したスタッフコードの列だけを見る。
        for iColumnIndex in range(iSheet6ColumnCount):
            objValueStaffCodeAtColumn = objDataFrameSheet6.iat[1, iColumnIndex]
            pszStaffCodeAtColumn: str = ""
            if objValueStaffCodeAtColumn is not None:
                pszStaffCodeAtColumn = str(objValueStaffCodeAtColumn).strip()
            if pszStaffCodeAtColumn != pszStaffCode:
                # この列は別のスタッフなのでスキップ
                continue

            # この列が対象スタッフの列
            for iRowIndex in range(2, iSheet6RowCount):
                objValueProject = objDataFrameSheet6.iat[iRowIndex, iColumnIndex]
                pszProjectNameFromSheet6: str = ""
                if objValueProject is not None:
                    pszProjectNameFromSheet6 = str(objValueProject).strip()
                if len(pszProjectNameFromSheet6) == 0:
                    # 空セル（これ以降も空の可能性が高いが、全行チェックでも問題なし）
                    continue

                # このスタッフ・このプロジェクトについて、工数秒数を合計
                objDataFrameSubProject: DataFrame = objDataFrameSubStaff[
                    objDataFrameSubStaff["プロジェクト名"] == pszProjectNameFromSheet6
                ]

                if objDataFrameSubProject.empty:
                    # データ不整合で Sheet6 にだけプロジェクト名が存在する場合はスキップ
                    continue

                # __time_seconds__ を合計
                try:
                    iTotalSeconds: int = int(objDataFrameSubProject["__time_seconds__"].sum())
                except Exception:
                    iTotalSeconds = 0

                pszTimeTotal: str = convert_seconds_to_time_string(iTotalSeconds)

                # スタッフコードは各行すべてに出力
                pszStaffCodeForRow: str = pszStaffCode
                # 氏名はスタッフの最初の行だけ出力
                if iRowIndexWithinStaff == 0:
                    pszStaffNameForRow: str = pszStaffNameForSheet8
                else:
                    pszStaffNameForRow = ""

                # Sheet7 用の 1 行
                objListOutputRowsSheet7.append(
                    [pszProjectNameFromSheet6, pszStaffCodeForRow, pszTimeTotal], 
                )

                # Sheet8 用の 1 行（Sheet7 の前に氏名列を追加）
                objListOutputRowsSheet8.append(
                    [pszStaffNameForRow, pszProjectNameFromSheet6, pszStaffCodeForRow, pszTimeTotal], 
                )

                iRowIndexWithinStaff += 1

            # 同じスタッフコードの列は 1 列だけの想定なので break
            break

    # -----------------------------------------------------------
    # TSV 出力（Sheet7, Sheet8, Sheet9）
    # -----------------------------------------------------------
    try:
        objDataFrameOutputSheet7: DataFrame = DataFrame(objListOutputRowsSheet7)
        objDataFrameOutputSheet7.to_csv(
            pszSheet7FileFullPath, 
            sep="\t", 
            index=False, 
            header=False, 
            encoding="utf-8", 
            lineterminator="\n", 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet7ErrorFileFullPath, 
            "Error: unexpected exception while writing Sheet7 TSV. Detail = {0}".format(objException), 
        )
        return

    try:
        objDataFrameOutputSheet8: DataFrame = DataFrame(objListOutputRowsSheet8)
        objDataFrameOutputSheet8.to_csv(
            pszSheet8FileFullPath, 
            sep="\t", 
            index=False, 
            header=False, 
            encoding="utf-8", 
            lineterminator="\n", 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet8ErrorFileFullPath, 
            "Error: unexpected exception while writing Sheet8 TSV. Detail = {0}".format(objException), 
        )
        return

    try:
        # Sheet9 はヘッダー行あり（氏名, スタッフコード）
        objDataFrameOutputSheet9: DataFrame = DataFrame(
            objListOutputRowsSheet9, 
            columns=["氏名", "スタッフコード"], 
        )
        objDataFrameOutputSheet9.to_csv(
            pszSheet9FileFullPath, 
            sep="\t", 
            index=False, 
            header=True, 
            encoding="utf-8", 
            lineterminator="\n", 
        )
    except Exception as objException:
        write_error_tsv(
            pszSheet9ErrorFileFullPath, 
            "Error: unexpected exception while writing Sheet9 TSV. Detail = {0}".format(objException), 
        )
        return


# ---------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------
def main() -> None:
    iArgCount: int = len(sys.argv)

    # 引数不足 → _error_argument.tsv を出力
    if iArgCount < 4:
        pszProgram: str = os.path.basename(sys.argv[0])
        pszErrorFile: str = "make_sheet789_from_sheet4_error_argument.tsv"

        write_error_tsv(
            pszErrorFile, 
            "Error: input TSV file paths are not specified (insufficient arguments).\n"
            "Usage: python {0} <Sheet4_tsv_file_path> <Sheet4_staff_code_range_tsv_file_path> <Sheet6_tsv_file_path>\n"
            "Example: python {0} Sheet4.tsv Sheet4_staff_code_range.tsv Sheet6.tsv".format(pszProgram), 
        )
        return

    pszSheet4FileFullPath: str = sys.argv[1]
    pszRangeFileFullPath: str = sys.argv[2]
    pszSheet6FileFullPath: str = sys.argv[3]

    make_sheet789_from_sheet4(
        pszSheet4FileFullPath, 
        pszRangeFileFullPath, 
        pszSheet6FileFullPath, 
    )


# ---------------------------------------------------------------
if __name__ == "__main__":
    main()
