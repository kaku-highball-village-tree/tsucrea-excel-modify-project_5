import os
import re
import sys
from typing import Dict, List, Optional, Tuple


def is_blank(pszValue: Optional[str]) -> bool:
    if pszValue is None:
        return True
    if pszValue == "":
        return True
    if pszValue.strip() == "":
        return True
    return False


def normalize_value(pszValue: str) -> str:
    return pszValue.strip()


def detect_type(pszValue: str) -> Optional[str]:
    pszNormalized: str = normalize_value(pszValue)
    if re.fullmatch(r"\d+:\d{2}:\d{2}", pszNormalized):
        return "time"
    if re.fullmatch(r"[+-]?\d+", pszNormalized):
        return "int"
    if re.fullmatch(r"[+-]?\d+\.\d+", pszNormalized):
        return "float"
    return None


def parse_arguments(objArgvList: List[str]) -> Optional[Tuple[str, str]]:
    objPositionalList: List[str] = []
    pszDelimiter: Optional[str] = None
    iIndex: int = 1
    while iIndex < len(objArgvList):
        pszArg: str = objArgvList[iIndex]
        if pszArg == "--delimiter":
            if iIndex + 1 >= len(objArgvList):
                print("--delimiter requires a value.")
                return None
            pszDelimiter = objArgvList[iIndex + 1]
            iIndex += 2
            continue
        objPositionalList.append(pszArg)
        iIndex += 1

    if len(objPositionalList) != 1:
        print("Usage: python FillZeroToBlank_Cmd.py INPUT [--delimiter \\\"\\t\\\"]")
        return None

    pszInputPath: str = objPositionalList[0]

    if pszDelimiter is None:
        if pszInputPath.lower().endswith(".tsv"):
            pszDelimiter = "\t"
        elif pszInputPath.lower().endswith(".csv"):
            pszDelimiter = ","
        else:
            print("Unable to determine delimiter from input extension. Specify --delimiter.")
            return None

    if pszDelimiter not in {"\t", ","}:
        print("Invalid delimiter. Use TAB (\\t) or comma.")
        return None

    return pszInputPath, pszDelimiter


def build_output_path(pszInputPath: str) -> Optional[str]:
    pszDirectory: str
    pszFileName: str
    pszDirectory, pszFileName = os.path.split(pszInputPath)
    pszRoot: str
    pszExt: str
    pszRoot, pszExt = os.path.splitext(pszFileName)
    if pszExt.lower() not in {".tsv", ".csv"}:
        return None
    pszOutputFileName: str = f"{pszRoot}_AppendNull{pszExt}"
    return os.path.join(pszDirectory, pszOutputFileName)


def load_rows(pszInputPath: str, pszDelimiter: str) -> List[List[str]]:
    objRowsList: List[List[str]] = []
    with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
        for pszLine in objFile:
            pszStripped: str = pszLine.rstrip("\n").rstrip("\r")
            objRowsList.append(pszStripped.split(pszDelimiter))
    return objRowsList


def get_max_columns(objRowsList: List[List[str]]) -> int:
    if not objRowsList:
        return 0
    return max(len(objRow) for objRow in objRowsList)


def determine_unit_score(objValuesList: List[str]) -> Tuple[Optional[str], Optional[float]]:
    iTotalNonBlank: int = 0
    objCountsDict: Dict[str, int] = {"time": 0, "int": 0, "float": 0}

    for pszValue in objValuesList:
        if is_blank(pszValue):
            continue
        iTotalNonBlank += 1
        pszType: Optional[str] = detect_type(pszValue)
        if pszType is None:
            continue
        objCountsDict[pszType] += 1

    if iTotalNonBlank == 0:
        return None, None

    pszRepresentative: Optional[str] = None
    iMaxCount: int = 0
    for pszKey in ("time", "int", "float"):
        if objCountsDict[pszKey] > iMaxCount:
            iMaxCount = objCountsDict[pszKey]
            pszRepresentative = pszKey

    if pszRepresentative is None or iMaxCount == 0:
        return None, None

    fScore: float = iMaxCount / float(iTotalNonBlank)
    return pszRepresentative, fScore


def evaluate_direction(objRowsList: List[List[str]]) -> str:
    iMaxColumns: int = get_max_columns(objRowsList)
    objColumnScoresList: List[float] = []
    for iCol in range(iMaxColumns):
        objValuesList: List[str] = []
        for objRow in objRowsList:
            if iCol < len(objRow):
                objValuesList.append(objRow[iCol])
            else:
                objValuesList.append("")
        pszRepresentative, fScore = determine_unit_score(objValuesList)
        if fScore is not None:
            objColumnScoresList.append(fScore)

    objRowScoresList: List[float] = []
    for objRow in objRowsList:
        pszRepresentative, fScore = determine_unit_score(objRow)
        if fScore is not None:
            objRowScoresList.append(fScore)

    fColumnScore: float = sum(objColumnScoresList) / len(objColumnScoresList) if objColumnScoresList else 0.0
    fRowScore: float = sum(objRowScoresList) / len(objRowScoresList) if objRowScoresList else 0.0

    if fColumnScore >= fRowScore:
        return "column"
    return "row"


def determine_representatives(objRowsList: List[List[str]], pszDirection: str) -> List[Optional[str]]:
    objRepresentativesList: List[Optional[str]] = []
    if pszDirection == "column":
        iMaxColumns: int = get_max_columns(objRowsList)
        for iCol in range(iMaxColumns):
            objValuesList: List[str] = []
            for objRow in objRowsList:
                if iCol < len(objRow):
                    objValuesList.append(objRow[iCol])
                else:
                    objValuesList.append("")
            pszRepresentative, _ = determine_unit_score(objValuesList)
            objRepresentativesList.append(pszRepresentative)
    else:
        for objRow in objRowsList:
            pszRepresentative, _ = determine_unit_score(objRow)
            objRepresentativesList.append(pszRepresentative)
    return objRepresentativesList


def should_blank_time(pszValue: str) -> bool:
    pszNormalized: str = normalize_value(pszValue)
    return pszNormalized in {"0:00:00", "00:00:00"}


def should_blank_int(pszValue: str) -> bool:
    pszNormalized: str = normalize_value(pszValue)
    if not re.fullmatch(r"[+-]?\d+", pszNormalized):
        return False
    try:
        iNumber: int = int(pszNormalized)
    except ValueError:
        return False
    return iNumber == 0


def should_blank_float(pszValue: str) -> bool:
    pszNormalized: str = normalize_value(pszValue)
    if not re.fullmatch(r"[+-]?\d+\.\d+", pszNormalized):
        return False
    try:
        fNumber: float = float(pszNormalized)
    except ValueError:
        return False
    return fNumber == 0.0


def convert_cells(objRowsList: List[List[str]], pszDirection: str, objRepresentativesList: List[Optional[str]]) -> Tuple[List[List[str]], int, int]:
    objResultRowsList: List[List[str]] = []
    iBlankedCount: int = 0
    iUnchangedCount: int = 0

    if pszDirection == "column":
        iMaxColumns: int = get_max_columns(objRowsList)
        for objRow in objRowsList:
            objNewRow: List[str] = []
            for iCol in range(iMaxColumns):
                pszValue: str = objRow[iCol] if iCol < len(objRow) else ""
                pszRepresentative: Optional[str] = objRepresentativesList[iCol] if iCol < len(objRepresentativesList) else None
                if is_blank(pszValue) or pszRepresentative is None:
                    objNewRow.append(pszValue)
                    iUnchangedCount += 1
                    continue
                if pszRepresentative == "time" and should_blank_time(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                if pszRepresentative == "int" and should_blank_int(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                if pszRepresentative == "float" and should_blank_float(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                objNewRow.append(pszValue)
                iUnchangedCount += 1
            objResultRowsList.append(objNewRow)
    else:
        for iRow, objRow in enumerate(objRowsList):
            pszRepresentative: Optional[str] = objRepresentativesList[iRow] if iRow < len(objRepresentativesList) else None
            objNewRow: List[str] = []
            for pszValue in objRow:
                if is_blank(pszValue) or pszRepresentative is None:
                    objNewRow.append(pszValue)
                    iUnchangedCount += 1
                    continue
                if pszRepresentative == "time" and should_blank_time(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                if pszRepresentative == "int" and should_blank_int(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                if pszRepresentative == "float" and should_blank_float(pszValue):
                    objNewRow.append("")
                    iBlankedCount += 1
                    continue
                objNewRow.append(pszValue)
                iUnchangedCount += 1
            objResultRowsList.append(objNewRow)

    return objResultRowsList, iBlankedCount, iUnchangedCount


def write_output(pszOutputPath: str, pszDelimiter: str, objRowsList: List[List[str]]) -> None:
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objFile:
        for objRow in objRowsList:
            objFile.write(pszDelimiter.join(objRow) + "\n")


def main(objArgvList: List[str]) -> int:
    objParsed = parse_arguments(objArgvList)
    if objParsed is None:
        return 1

    pszInputPath, pszDelimiter = objParsed

    pszOutputPath: Optional[str] = build_output_path(pszInputPath)
    if pszOutputPath is None:
        print("Input extension must be .tsv or .csv.")
        return 1

    if pszInputPath == pszOutputPath:
        print("Input and output paths must be different.")
        return 1

    if not os.path.exists(pszInputPath):
        print("Input file not found.")
        return 1

    try:
        objRowsList: List[List[str]] = load_rows(pszInputPath, pszDelimiter)
    except Exception as objExc:  # noqa: BLE001
        print(f"Failed to read input: {objExc}")
        return 1

    pszDirection: str = evaluate_direction(objRowsList)
    objRepresentativesList: List[Optional[str]] = determine_representatives(objRowsList, pszDirection)

    objConvertedRowsList, iBlankedCount, iUnchangedCount = convert_cells(objRowsList, pszDirection, objRepresentativesList)

    try:
        write_output(pszOutputPath, pszDelimiter, objConvertedRowsList)
    except Exception as objExc:  # noqa: BLE001
        print(f"Failed to write output: {objExc}")
        return 1

    print(f"direction={pszDirection}, blanked={iBlankedCount}, unchanged={iUnchangedCount}, output={pszOutputPath}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
