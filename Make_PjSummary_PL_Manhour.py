import csv
import os
import re
import sys
from typing import List, Tuple

from src import make_manhour_to_sheet8_01_0001


def get_target_year_month_from_filename(pszInputFilePath: str) -> Tuple[int, int]:
    pszBaseName: str = os.path.basename(pszInputFilePath)
    objMatch: re.Match[str] | None = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        raise ValueError("入力ファイル名から対象年月を取得できません。")
    iYearTwoDigits: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    iYear: int = 2000 + iYearTwoDigits
    return iYear, iMonth


def get_target_year_month_from_period_row(pszRowA: str) -> Tuple[int, int]:
    pszNormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
    pszNormalized = pszNormalized.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    objMatch: re.Match[str] | None = re.search(r"(?:自)?(\d{4})年(\d{1,2})月(?:度)?", pszNormalized)
    if objMatch is None:
        objMatch = re.search(r"(\d{4})[./-](\d{1,2})", pszNormalized)
    if objMatch is None:
        raise ValueError("集計期間から対象年月を取得できません。")
    iYear: int = int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    return iYear, iMonth


def read_csv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    try:
        with open(pszInputFilePath, mode="r", encoding="utf-8-sig", errors="strict", newline="") as objFile:
            objReader: csv.reader = csv.reader(objFile)
            for objRow in objReader:
                objRows.append(objRow)
        append_debug_log("input decoded as utf-8-sig")
        return objRows
    except UnicodeDecodeError:
        append_debug_log("utf-8-sig decode failed; retrying with cp932")

    with open(pszInputFilePath, mode="r", encoding="cp932", errors="strict", newline="") as objFile:
        objReader = csv.reader(objFile)
        for objRow in objReader:
            objRows.append(objRow)
    append_debug_log("input decoded as cp932")
    return objRows


def write_tsv_rows(pszOutputFilePath: str, objRows: List[List[str]]) -> None:
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objFile:
        objWriter: csv.writer = csv.writer(objFile, delimiter="\t", lineterminator="\n")
        for objRow in objRows:
            objWriter.writerow(objRow)


def read_tsv_rows(pszInputFilePath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objFile:
        objReader: csv.reader = csv.reader(objFile, delimiter="\t")
        for objRow in objReader:
            objRows.append(objRow)
    return objRows


def build_first_column_rows(objRows: List[List[str]]) -> List[List[str]]:
    return [[objRow[0] if objRow else ""] for objRow in objRows]



def build_unique_subjects(objSubjectRows: List[List[str]]) -> List[str]:
    objSubjects: List[str] = []
    objSeen: set[str] = set()
    for objRow in objSubjectRows:
        pszValue: str = objRow[0] if objRow else ""
        if pszValue == "" or pszValue in objSeen:
            continue
        objSeen.add(pszValue)
        objSubjects.append(pszValue)
    return objSubjects



def build_union_subject_order(objSubjectLists: List[List[str]]) -> List[str]:
    objAppearanceOrder: dict[str, int] = {}
    iCounter: int = 0
    for objSubjectList in objSubjectLists:
        for pszSubject in objSubjectList:
            if pszSubject not in objAppearanceOrder:
                objAppearanceOrder[pszSubject] = iCounter
                iCounter += 1

    objAdjacency: dict[str, set[str]] = {psz: set() for psz in objAppearanceOrder}
    objIndegree: dict[str, int] = {psz: 0 for psz in objAppearanceOrder}
    for objSubjectList in objSubjectLists:
        for iIndex in range(len(objSubjectList) - 1):
            pszBefore: str = objSubjectList[iIndex]
            pszAfter: str = objSubjectList[iIndex + 1]
            if pszAfter not in objAdjacency[pszBefore]:
                objAdjacency[pszBefore].add(pszAfter)
                objIndegree[pszAfter] += 1

    objOrderedSubjects: List[str] = []
    objReady: List[str] = [
        pszSubject for pszSubject, iDegree in objIndegree.items() if iDegree == 0
    ]
    objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    while objReady:
        pszSubject = objReady.pop(0)
        objOrderedSubjects.append(pszSubject)
        for pszNext in sorted(objAdjacency[pszSubject], key=lambda psz: objAppearanceOrder[psz]):
            objIndegree[pszNext] -= 1
            if objIndegree[pszNext] == 0:
                objReady.append(pszNext)
        objReady.sort(key=lambda pszSubject: objAppearanceOrder[pszSubject])

    if len(objOrderedSubjects) != len(objAppearanceOrder):
        return list(objAppearanceOrder.keys())

    return objOrderedSubjects



def build_subject_vertical_rows(objSubjects: List[str]) -> List[List[str]]:
    return [[pszSubject] for pszSubject in objSubjects]



def normalize_project_name(pszProjectName: str) -> str:
    if pszProjectName == "":
        return pszProjectName
    normalized = pszProjectName.replace("\t", "_")
    normalized = re.sub(r"^([A-OQ-Z]\d{3})([ \u3000]+)", r"\1_", normalized)
    normalized = re.sub(r"^([A-OQ-Z]\d{3})(【)", r"\1_\2", normalized)
    normalized = re.sub(r"^(P\d{5})([ \u3000]+)", r"\1_", normalized)
    normalized = re.sub(r"^(P\d{5})(【)", r"\1_\2", normalized)
    return normalized


def normalize_project_names_in_row(objRows: List[List[str]], iRowIndex: int) -> None:
    if iRowIndex < 0 or iRowIndex >= len(objRows):
        return
    objTargetRow = objRows[iRowIndex]
    for iIndex, pszProjectName in enumerate(objTargetRow):
        objTargetRow[iIndex] = normalize_project_name(pszProjectName)


def find_row_index_with_subject_tab(objRows: List[List[str]], iStartIndex: int) -> int | None:
    for iRowIndex in range(iStartIndex, len(objRows)):
        objRow = objRows[iRowIndex]
        if any(
            "科目名\t" in pszValue or pszValue.strip() == "科目名"
            for pszValue in objRow
        ):
            return iRowIndex
    return None


def build_pj_name_vertical_rows(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []

    objHeaderRow: List[str] = objRows[0]
    objItemRows: List[List[str]] = objRows[1:]

    objVerticalRows: List[List[str]] = []
    objVerticalHeader: List[str] = ["PJ名称"]
    for objItemRow in objItemRows:
        pszItemName: str = objItemRow[0] if len(objItemRow) > 0 else ""
        objVerticalHeader.append(pszItemName)
    objVerticalRows.append(objVerticalHeader)

    for iColumnIndex in range(1, len(objHeaderRow)):
        pszProjectName: str = objHeaderRow[iColumnIndex]
        objVerticalRow: List[str] = [pszProjectName]
        for objItemRow in objItemRows:
            pszValue: str = objItemRow[iColumnIndex] if len(objItemRow) > iColumnIndex else ""
            objVerticalRow.append(pszValue)
        objVerticalRows.append(objVerticalRow)

    return objVerticalRows


def write_first_row_tabs_to_newlines(pszInputFilePath: str, pszOutputFilePath: str) -> None:
    with open(pszInputFilePath, mode="r", encoding="utf-8", newline="") as objInputFile:
        pszFirstLine: str = objInputFile.readline()
    pszConverted: str = pszFirstLine.replace("\t", "\n")
    with open(pszOutputFilePath, mode="w", encoding="utf-8", newline="") as objOutputFile:
        objOutputFile.write(pszConverted)


def insert_company_expense_columns(objRows: List[List[str]]) -> None:
    if not objRows:
        return
    objHeaderRow: List[str] = objRows[0]
    try:
        iHeadOfficeIndex: int = objHeaderRow.index("本部")
    except ValueError:
        return

    objExpenseColumns: List[str] = [
        "1Cカンパニー販管費",
        "2Cカンパニー販管費",
        "3Cカンパニー販管費",
        "4Cカンパニー販管費",
        "事業開発カンパニー販管費",
        "社長室カンパニー販管費",
        "本部カンパニー販管費",
    ]
    iInsertIndex: int = iHeadOfficeIndex + 1
    objHeaderRow[iInsertIndex:iInsertIndex] = objExpenseColumns
    for objRow in objRows[1:]:
        objRow[iInsertIndex:iInsertIndex] = ["0"] * len(objExpenseColumns)


COMPANY_EXPENSE_REPLACEMENTS: dict[str, str] = {
    "1Cカンパニー販管費": "C001_1Cカンパニー販管費",
    "2Cカンパニー販管費": "C002_2Cカンパニー販管費",
    "3Cカンパニー販管費": "C003_3Cカンパニー販管費",
    "4Cカンパニー販管費": "C004_4Cカンパニー販管費",
    "事業開発カンパニー販管費": "C005_事業開発カンパニー販管費",
    "社長室カンパニー販管費": "C006_社長室カンパニー販管費",
    "本部カンパニー販管費": "C007_本部カンパニー販管費",
}


def replace_company_expense_labels(objRows: List[List[str]], objReplacementMap: dict[str, str]) -> None:
    for objRow in objRows:
        for iIndex, pszValue in enumerate(objRow):
            if pszValue in objReplacementMap:
                objRow[iIndex] = objReplacementMap[pszValue]


def append_debug_log(pszMessage: str, pszDebugFilePath: str = "debug.txt") -> None:
    with open(pszDebugFilePath, mode="a", encoding="utf-8", newline="") as objDebugFile:
        objDebugFile.write(f"{pszMessage}\n")


def create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths: List[str]) -> None:
    if not objCostReportVerticalFilePaths:
        return

    objSubjectLists: List[List[str]] = []
    for pszFilePath in objCostReportVerticalFilePaths:
        objRows: List[List[str]] = read_tsv_rows(pszFilePath)
        objSubjectLists.append(build_unique_subjects(objRows))
    objUnionSubjects: List[str] = build_union_subject_order(objSubjectLists)
    objUnionRows: List[List[str]] = build_subject_vertical_rows(objUnionSubjects)

    for pszFilePath in objCostReportVerticalFilePaths:
        pszUnionFilePath: str = pszFilePath.replace("_科目名_vertical.tsv", "_科目名_A∪B_vertical.tsv")
        write_tsv_rows(pszUnionFilePath, objUnionRows)
        append_debug_log(f"union vertical tsv written: {pszUnionFilePath}")


def run_pl_csv_to_tsv(pszInputFilePath: str) -> Tuple[int, List[str]]:
    iExitCode: int = 0
    objCostReportVerticalFilePaths: List[str] = []
    try:
        append_debug_log("start")
        iFileYear: int
        iFileMonth: int
        iFileYear, iFileMonth = get_target_year_month_from_filename(pszInputFilePath)
        append_debug_log(f"filename parsed: {iFileYear}-{iFileMonth:02d}")

        if not os.path.isfile(pszInputFilePath):
            raise FileNotFoundError(f"入力ファイルが存在しません: {pszInputFilePath}")

        objRows: List[List[str]] = read_csv_rows(pszInputFilePath)
        if len(objRows) < 2:
            raise ValueError("集計期間の取得に必要な行が存在しません。")
        append_debug_log(f"rows read: {len(objRows)}")

        normalize_project_names_in_row(objRows, 7)
        iSubjectRowIndex = find_row_index_with_subject_tab(objRows, 8)
        if iSubjectRowIndex is not None:
            normalize_project_names_in_row(objRows, iSubjectRowIndex)
        append_debug_log("project names normalized")

        pszRowA: str = objRows[1][1] if len(objRows[1]) > 1 else ""
        append_debug_log(f"B2 value: {pszRowA}")
        pszRowANormalized: str = re.sub(r"[ \u3000]", "", pszRowA)
        if "期首振戻" in pszRowANormalized:
            append_debug_log("period parse skipped due to 期首振戻; using filename")
        else:
            iPeriodYear: int
            iPeriodMonth: int
            iPeriodYear, iPeriodMonth = get_target_year_month_from_period_row(pszRowA)
            append_debug_log(f"period parsed: {iPeriodYear}-{iPeriodMonth:02d}")

            if iFileYear != iPeriodYear or iFileMonth != iPeriodMonth:
                raise ValueError("ファイル名と集計期間の対象年月が一致しません。")
            append_debug_log("period matches filename")

        pszMonth: str = f"{iFileMonth:02d}"
        pszOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月.tsv"
        pszCostReportFilePath: str = f"製造原価報告書_{iFileYear}年{pszMonth}月.tsv"
        objOutputRows: List[List[str]] = []
        objCostReportRows: List[List[str]] = []
        iSplitIndex: int | None = None
        for iRowIndex in range(7, len(objRows) - 1):
            objRow: List[str] = objRows[iRowIndex]
            objNextRow: List[str] = objRows[iRowIndex + 1]
            if objRow and objNextRow and objRow[0] == "当期純利益" and objNextRow[0] == "科目名":
                iSplitIndex = iRowIndex
                break

        if iSplitIndex is None:
            for iRowIndex in range(7, len(objRows)):
                objRow = objRows[iRowIndex]
                objOutputRows.append(objRow[:])
        else:
            for iRowIndex in range(7, iSplitIndex + 1):
                objRow = objRows[iRowIndex]
                objOutputRows.append(objRow[:])
            for iRowIndex in range(iSplitIndex + 1, len(objRows)):
                objRow = objRows[iRowIndex]
                objCostReportRows.append(objRow[:])
        append_debug_log(f"output rows prepared: {len(objOutputRows)}")

        if (iFileYear, iFileMonth) <= (2025, 7):
            insert_company_expense_columns(objOutputRows)
            append_debug_log("company expense columns inserted")

        replace_company_expense_labels(
            objOutputRows,
            COMPANY_EXPENSE_REPLACEMENTS,
        )
        append_debug_log("company expense labels replaced")

        write_tsv_rows(pszOutputFilePath, objOutputRows)
        append_debug_log(f"tsv written: {pszOutputFilePath}")

        if objCostReportRows:
            write_tsv_rows(pszCostReportFilePath, objCostReportRows)
            append_debug_log(f"tsv written: {pszCostReportFilePath}")
            objCostReportTsvRows: List[List[str]] = read_tsv_rows(pszCostReportFilePath)
            objCostReportVerticalRows: List[List[str]] = build_first_column_rows(objCostReportTsvRows)
            pszCostReportVerticalFilePath: str = (
                f"製造原価報告書_{iFileYear}年{pszMonth}月_科目名_vertical.tsv"
            )
            write_tsv_rows(pszCostReportVerticalFilePath, objCostReportVerticalRows)
            append_debug_log(f"vertical tsv written: {pszCostReportVerticalFilePath}")

            objCostReportVerticalFilePaths.append(pszCostReportVerticalFilePath)

        pszVerticalOutputFilePath: str = f"損益計算書_{iFileYear}年{pszMonth}月_PJ名称_vertical.tsv"
        objPjNameVerticalRows: List[List[str]] = build_pj_name_vertical_rows(objOutputRows)
        write_tsv_rows(pszVerticalOutputFilePath, objPjNameVerticalRows)
        append_debug_log(f"vertical tsv written: {pszVerticalOutputFilePath}")
    except Exception as objException:
        iExitCode = 1
        append_debug_log(f"error: {objException}")
        print(objException)
        try:
            iErrorYear: int
            iErrorMonth: int
            iErrorYear, iErrorMonth = get_target_year_month_from_filename(pszInputFilePath)
            pszErrorMonth: str = f"{iErrorMonth:02d}"
            pszErrorFilePath: str = f"損益計算書_{iErrorYear}年{pszErrorMonth}月_error.txt"
        except Exception:
            pszBaseName: str = os.path.basename(pszInputFilePath)
            pszErrorFilePath = f"{pszBaseName}_error.txt"
        with open(pszErrorFilePath, mode="w", encoding="utf-8", newline="") as objErrorFile:
            objErrorFile.write(str(objException))

    return iExitCode, objCostReportVerticalFilePaths


def run_manhour_pipeline(pszInputFilePath: str) -> int:
    objOriginalArgv: List[str] = sys.argv[:]
    try:
        sys.argv = ["make_manhour_to_sheet8_01_0001.py", pszInputFilePath]
        return make_manhour_to_sheet8_01_0001.main()
    finally:
        sys.argv = objOriginalArgv


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python Make_PjSummary_PL_Manhour.py <csv_file> [<csv_file> ...]")
        return 1

    iExitCode: int = 0
    objCostReportVerticalFilePaths: List[str] = []
    for pszInputFilePath in sys.argv[1:]:
        iPlExitCode: int
        objPlCostReportVerticalFilePaths: List[str]
        iPlExitCode, objPlCostReportVerticalFilePaths = run_pl_csv_to_tsv(pszInputFilePath)
        objCostReportVerticalFilePaths.extend(objPlCostReportVerticalFilePaths)
        if iPlExitCode != 0:
            iExitCode = 1

        iManhourExitCode: int = run_manhour_pipeline(pszInputFilePath)
        if iManhourExitCode != 0:
            iExitCode = 1

    create_union_subject_vertical_tsvs(objCostReportVerticalFilePaths)
    return iExitCode


if __name__ == "__main__":
    sys.exit(main())
