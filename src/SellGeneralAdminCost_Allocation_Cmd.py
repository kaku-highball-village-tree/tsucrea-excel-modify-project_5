# -*- coding: utf-8 -*-
"""
SellGeneralAdminCost_Allocation_Cmd.py

販管費配賦（工数付与）TSV を生成する。

入力:
  1) 工数_yyyy年mm月_step10_各プロジェクトの工数.tsv
  2) 損益計算書_yyyy年mm月_A∪B_プロジェクト名_C∪D_vertical.tsv

出力:
  損益計算書_yyyy年mm月_販管費配賦_A∪B_プロジェクト名_C∪D_vertical.tsv

処理:
  損益計算書TSVの各行に対し、
  プロジェクト行(A/C/J/Pで始まり、"_"までのキー)のみ
  工数TSVに同一キーがある場合は工数を、
  無い場合は 0:00:00 を末尾に追加する。
  非プロジェクト行はそのまま出力する。
"""

from __future__ import annotations

import os
import shutil
import re
import sys
from typing import Dict, List, Optional, Tuple


def print_usage() -> None:
    pszUsage: str = (
        "Usage: python SellGeneralAdminCost_Allocation_Cmd.py "
        "<manhour_tsv_path> <pl_tsv_path> [output_tsv_path]\n"
        "   or: python SellGeneralAdminCost_Allocation_Cmd.py "
        "<manhour_tsv_path> <pl_tsv_path> <manhour_tsv_path> <pl_tsv_path> ...\n"
        "   or: python SellGeneralAdminCost_Allocation_Cmd.py "
        "<manhour_tsv_path> ... <pl_tsv_path> ..."
    )
    print(pszUsage)


def build_default_output_path(pszInputPlPath: str) -> str:
    pszDirectory: str
    pszFileName: str
    pszDirectory, pszFileName = os.path.split(pszInputPlPath)

    pszStem: str
    pszExt: str
    pszStem, pszExt = os.path.splitext(pszFileName)
    if pszExt == "":
        pszExt = ".tsv"

    pszTargetMarker: str = "損益計算書_"
    pszSuffix: str = "販管費配賦_"
    pszStepMarker: str = "販管費配賦_step0006_"
    pszStepMarkerOld: str = "販管費配賦_step0001_"
    pszStepMarkerPrevious: str = "販管費配賦_step0002_"
    pszStepMarkerCurrent: str = "販管費配賦_step0003_"
    pszStepMarkerNext: str = "販管費配賦_step0004_"
    pszStepMarkerAfterNext: str = "販管費配賦_step0005_"

    if pszStepMarkerOld in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerOld, pszStepMarker, 1)
    elif pszStepMarkerPrevious in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerPrevious, pszStepMarker, 1)
    elif pszStepMarkerCurrent in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerCurrent, pszStepMarker, 1)
    elif pszStepMarkerNext in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerNext, pszStepMarker, 1)
    elif pszStepMarkerAfterNext in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerAfterNext, pszStepMarker, 1)
    elif pszTargetMarker in pszStem and pszSuffix not in pszStem:
        pszOutputStem = pszStem.replace(pszTargetMarker, pszTargetMarker + pszStepMarker, 1)
    elif pszSuffix in pszStem and pszStepMarker not in pszStem:
        pszOutputStem = pszStem.replace(pszSuffix, pszStepMarker, 1)
    else:
        pszOutputStem = pszStem + "_販管費配賦"

    pszOutputFileName: str = pszOutputStem + pszExt
    pszOutputPath: str = os.path.join(pszDirectory, pszOutputFileName)
    return pszOutputPath


def build_output_path_with_step(pszInputPlPath: str, pszStepMarker: str) -> str:
    pszDirectory: str
    pszFileName: str
    pszDirectory, pszFileName = os.path.split(pszInputPlPath)

    pszStem: str
    pszExt: str
    pszStem, pszExt = os.path.splitext(pszFileName)
    if pszExt == "":
        pszExt = ".tsv"

    pszTargetMarker: str = "損益計算書_"
    pszSuffix: str = "販管費配賦_"
    pszStepMarkerOld: str = "販管費配賦_step0001_"
    pszStepMarkerCurrent: str = "販管費配賦_step0002_"
    pszStepMarkerNext: str = "販管費配賦_step0003_"
    pszStepMarkerAfterNext: str = "販管費配賦_step0004_"
    pszStepMarkerAfterAfterNext: str = "販管費配賦_step0005_"

    if pszStepMarkerOld in pszStem:
        pszOutputStem: str = pszStem.replace(pszStepMarkerOld, pszStepMarker, 1)
    elif pszStepMarkerCurrent in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerCurrent, pszStepMarker, 1)
    elif pszStepMarkerNext in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerNext, pszStepMarker, 1)
    elif pszStepMarkerAfterNext in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerAfterNext, pszStepMarker, 1)
    elif pszStepMarkerAfterAfterNext in pszStem:
        pszOutputStem = pszStem.replace(pszStepMarkerAfterAfterNext, pszStepMarker, 1)
    elif pszTargetMarker in pszStem and pszSuffix not in pszStem:
        pszOutputStem = pszStem.replace(pszTargetMarker, pszTargetMarker + pszStepMarker, 1)
    elif pszSuffix in pszStem and pszStepMarker not in pszStem:
        pszOutputStem = pszStem.replace(pszSuffix, pszStepMarker, 1)
    else:
        pszOutputStem = pszStem + "_販管費配賦"

    pszOutputFileName: str = pszOutputStem + pszExt
    pszOutputPath: str = os.path.join(pszDirectory, pszOutputFileName)
    return pszOutputPath


def extract_project_key(pszProjectName: str) -> Optional[str]:
    pszText: str = (pszProjectName or "").strip()
    if pszText == "":
        return None

    iUnderscoreIndex: int = pszText.find("_")
    if iUnderscoreIndex <= 0:
        pszKey = pszText.split(" ", 1)[0]
    else:
        pszKey = pszText[:iUnderscoreIndex]
    if pszKey == "":
        return None

    cPrefix: str = pszKey[0]
    if cPrefix in ("A", "C", "J", "P"):
        return pszKey
    return None


def load_manhour_map(pszManhourPath: str) -> Dict[str, str]:
    objManhourMap: Dict[str, str] = {}
    with open(pszManhourPath, "r", encoding="utf-8", newline="") as objInputFile:
        for pszLine in objInputFile:
            pszLineText: str = pszLine.rstrip("\n").rstrip("\r")
            if pszLineText == "":
                continue

            pszFirstColumn: str
            if "\t" in pszLineText:
                pszFirstColumn = pszLineText.split("\t", 1)[0]
            else:
                pszFirstColumn = pszLineText

            pszKey: Optional[str] = extract_project_key(pszFirstColumn)
            if pszKey is None:
                continue

            pszTimeValue: str = ""
            if "\t" in pszLineText:
                objParts: List[str] = pszLineText.split("\t")
                if len(objParts) >= 2:
                    pszTimeValue = objParts[1]
            objManhourMap[pszKey] = pszTimeValue

    return objManhourMap


def parse_number(pszText: str) -> float:
    pszValue: str = (pszText or "").strip()
    if pszValue == "":
        return 0.0
    try:
        return float(pszValue)
    except ValueError:
        return 0.0


def parse_time_to_seconds(pszTimeText: str) -> float:
    pszValue: str = (pszTimeText or "").strip()
    if pszValue == "":
        return 0.0

    objParts: List[str] = pszValue.split(":")
    if len(objParts) != 3:
        return 0.0
    try:
        iHours: int = int(objParts[0])
        iMinutes: int = int(objParts[1])
        iSeconds: int = int(objParts[2])
    except ValueError:
        return 0.0

    return float(iHours * 3600 + iMinutes * 60 + iSeconds)


def format_number(fValue: float) -> str:
    if abs(fValue - round(fValue)) < 0.0000001:
        return str(int(round(fValue)))
    pszText: str = f"{fValue:.6f}"
    pszText = pszText.rstrip("0").rstrip(".")
    return pszText


def calculate_allocation(
    objRows: List[List[str]],
    iSellGeneralAdminCostColumnIndex: int,
    iAllocationColumnIndex: int,
    iManhourColumnIndex: int,
) -> None:
    iRowIndexTotal: int = 1
    iRowIndexAllocationStart: int = 3
    iRowIndexAllocationEnd: int = 7
    iRowIndexProjectStart: int = 10
    iRowIndexProjectEnd: int = 123

    fSellGeneralAdminCostTotal: float = 0.0
    if iRowIndexTotal < len(objRows) and iSellGeneralAdminCostColumnIndex >= 0:
        objRowTotal: List[str] = objRows[iRowIndexTotal]
        if iSellGeneralAdminCostColumnIndex < len(objRowTotal):
            fSellGeneralAdminCostTotal = parse_number(objRowTotal[iSellGeneralAdminCostColumnIndex])

    fAllocatedSum: float = 0.0
    for iRowIndex in range(iRowIndexAllocationStart, iRowIndexAllocationEnd + 1):
        if iRowIndex >= len(objRows):
            break
        objRow: List[str] = objRows[iRowIndex]
        if iAllocationColumnIndex < len(objRow):
            fAllocatedSum += parse_number(objRow[iAllocationColumnIndex])

    fSellGeneralAdminCostAllocation: float = fSellGeneralAdminCostTotal - fAllocatedSum

    fTotalManhours: float = 0.0
    for iRowIndex in range(iRowIndexProjectStart, iRowIndexProjectEnd + 1):
        if iRowIndex >= len(objRows):
            break
        objRow: List[str] = objRows[iRowIndex]
        if iManhourColumnIndex < len(objRow):
            fTotalManhours += parse_time_to_seconds(objRow[iManhourColumnIndex])

    if fTotalManhours <= 0.0:
        return

    for iRowIndex in range(iRowIndexProjectStart, iRowIndexProjectEnd + 1):
        if iRowIndex >= len(objRows):
            break
        objRow = objRows[iRowIndex]
        if iManhourColumnIndex >= len(objRow):
            continue
        fManhourSeconds: float = parse_time_to_seconds(objRow[iManhourColumnIndex])
        fAllocation: float = fSellGeneralAdminCostAllocation * fManhourSeconds / fTotalManhours
        fAllocation = float(int(round(fAllocation)))

        if iAllocationColumnIndex >= len(objRow):
            iAppendCount: int = iAllocationColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iAllocationColumnIndex] = format_number(fAllocation)
        objRows[iRowIndex] = objRow


def recalculate_operating_profit(
    objRows: List[List[str]],
    iGrossProfitColumnIndex: int,
    iOperatingProfitColumnIndex: int,
) -> None:
    if iGrossProfitColumnIndex < 0 or iOperatingProfitColumnIndex < 0:
        return
    if iOperatingProfitColumnIndex <= iGrossProfitColumnIndex:
        return

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]
        if iGrossProfitColumnIndex >= len(objRow):
            continue

        fGrossProfit: float = parse_number(objRow[iGrossProfitColumnIndex])
        fDeductionSum: float = 0.0
        for iColumnIndex in range(iGrossProfitColumnIndex + 1, iOperatingProfitColumnIndex):
            if iColumnIndex >= len(objRow):
                continue
            fDeductionSum += parse_number(objRow[iColumnIndex])

        fOperatingProfit: float = fGrossProfit - fDeductionSum
        if iOperatingProfitColumnIndex >= len(objRow):
            iAppendCount: int = iOperatingProfitColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iOperatingProfitColumnIndex] = format_number(fOperatingProfit)
        objRows[iRowIndex] = objRow


def recalculate_ordinary_profit(
    objRows: List[List[str]],
    iOperatingProfitColumnIndex: int,
    iNonOperatingIncomeColumnIndex: int,
    iNonOperatingExpenseColumnIndex: int,
    iOrdinaryProfitColumnIndex: int,
) -> None:
    if (
        iOperatingProfitColumnIndex < 0
        or iNonOperatingIncomeColumnIndex < 0
        or iNonOperatingExpenseColumnIndex < 0
        or iOrdinaryProfitColumnIndex < 0
    ):
        return
    if iOperatingProfitColumnIndex >= iNonOperatingIncomeColumnIndex:
        return
    if iNonOperatingIncomeColumnIndex >= iNonOperatingExpenseColumnIndex:
        return
    if iNonOperatingExpenseColumnIndex >= iOrdinaryProfitColumnIndex:
        return

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]
        if iOperatingProfitColumnIndex >= len(objRow):
            continue

        fNonOperatingIncome: float = 0.0
        for iColumnIndex in range(iOperatingProfitColumnIndex + 1, iNonOperatingIncomeColumnIndex):
            if iColumnIndex >= len(objRow):
                continue
            fNonOperatingIncome += parse_number(objRow[iColumnIndex])

        fNonOperatingExpense: float = 0.0
        for iColumnIndex in range(iNonOperatingIncomeColumnIndex + 1, iNonOperatingExpenseColumnIndex):
            if iColumnIndex >= len(objRow):
                continue
            fNonOperatingExpense += parse_number(objRow[iColumnIndex])

        if iNonOperatingIncomeColumnIndex >= len(objRow):
            iAppendCount = iNonOperatingIncomeColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iNonOperatingIncomeColumnIndex] = format_number(fNonOperatingIncome)

        if iNonOperatingExpenseColumnIndex >= len(objRow):
            iAppendCount = iNonOperatingExpenseColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iNonOperatingExpenseColumnIndex] = format_number(fNonOperatingExpense)

        fOperatingProfit: float = parse_number(objRow[iOperatingProfitColumnIndex])
        fOrdinaryProfit: float = fOperatingProfit + fNonOperatingIncome - fNonOperatingExpense
        if iOrdinaryProfitColumnIndex >= len(objRow):
            iAppendCount = iOrdinaryProfitColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iOrdinaryProfitColumnIndex] = format_number(fOrdinaryProfit)
        objRows[iRowIndex] = objRow


def recalculate_pre_tax_profit(
    objRows: List[List[str]],
    iOrdinaryProfitColumnIndex: int,
    iExtraordinaryIncomeColumnIndex: int,
    iExtraordinaryLossColumnIndex: int,
    iPreTaxProfitColumnIndex: int,
) -> None:
    if (
        iOrdinaryProfitColumnIndex < 0
        or iExtraordinaryIncomeColumnIndex < 0
        or iExtraordinaryLossColumnIndex < 0
        or iPreTaxProfitColumnIndex < 0
    ):
        return

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]
        if iOrdinaryProfitColumnIndex >= len(objRow):
            continue

        fOrdinaryProfit: float = parse_number(objRow[iOrdinaryProfitColumnIndex])
        fExtraordinaryIncome: float = 0.0
        if iExtraordinaryIncomeColumnIndex < len(objRow):
            fExtraordinaryIncome = parse_number(objRow[iExtraordinaryIncomeColumnIndex])
        fExtraordinaryLoss: float = 0.0
        if iExtraordinaryLossColumnIndex < len(objRow):
            fExtraordinaryLoss = parse_number(objRow[iExtraordinaryLossColumnIndex])

        fPreTaxProfit: float = fOrdinaryProfit + fExtraordinaryIncome - fExtraordinaryLoss
        if iPreTaxProfitColumnIndex >= len(objRow):
            iAppendCount: int = iPreTaxProfitColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iPreTaxProfitColumnIndex] = format_number(fPreTaxProfit)
        objRows[iRowIndex] = objRow


def recalculate_net_profit(
    objRows: List[List[str]],
    iCorporateTaxColumnIndex: int,
    iCorporateTaxTotalColumnIndex: int,
    iPreTaxProfitColumnIndex: int,
    iNetProfitColumnIndex: int,
) -> None:
    if (
        iCorporateTaxColumnIndex < 0
        or iCorporateTaxTotalColumnIndex < 0
        or iPreTaxProfitColumnIndex < 0
        or iNetProfitColumnIndex < 0
    ):
        return

    for iRowIndex in range(1, len(objRows)):
        objRow: List[str] = objRows[iRowIndex]
        if iCorporateTaxColumnIndex >= len(objRow):
            continue

        fCorporateTax: float = parse_number(objRow[iCorporateTaxColumnIndex])
        if iCorporateTaxTotalColumnIndex >= len(objRow):
            iAppendCount: int = iCorporateTaxTotalColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iCorporateTaxTotalColumnIndex] = format_number(fCorporateTax)

        fPreTaxProfit: float = 0.0
        if iPreTaxProfitColumnIndex < len(objRow):
            fPreTaxProfit = parse_number(objRow[iPreTaxProfitColumnIndex])

        fNetProfit: float = fPreTaxProfit - fCorporateTax
        if iNetProfitColumnIndex >= len(objRow):
            iAppendCount = iNetProfitColumnIndex + 1 - len(objRow)
            objRow.extend([""] * iAppendCount)
        objRow[iNetProfitColumnIndex] = format_number(fNetProfit)
        objRows[iRowIndex] = objRow


def process_pl_tsv(
    pszPlPath: str,
    pszOutputPath: str,
    pszOutputStep0001Path: str,
    pszOutputStep0002Path: str,
    pszOutputStep0003Path: str,
    pszOutputStep0004Path: str,
    pszOutputStep0005Path: str,
    pszOutputStep0006Path: str,
    pszOutputFinalPath: str,
    objManhourMap: Dict[str, str],
) -> None:
    objRows: List[List[str]] = []
    with open(pszPlPath, "r", encoding="utf-8", newline="") as objInputFile:
        for pszLine in objInputFile:
            pszLineText: str = pszLine.rstrip("\n").rstrip("\r")
            objRows.append(pszLineText.split("\t") if pszLineText != "" else [""])

    for iRowIndex, objRow in enumerate(objRows):
        pszFirstColumn: str = objRow[0] if objRow else ""
        if iRowIndex == 0:
            if len(objRow) == 0:
                objRow = [""]
            objRow.append("工数")
            objRows[iRowIndex] = objRow
            continue

        pszKey: Optional[str] = extract_project_key(pszFirstColumn)
        if pszKey is None:
            continue

        pszManhour: str = objManhourMap.get(pszKey, "")
        if pszManhour == "":
            pszManhour = "0:00:00"

        objRow.append(pszManhour)
        objRows[iRowIndex] = objRow

    with open(pszOutputStep0001Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    iSellGeneralAdminCostColumnIndex: int = -1
    iAllocationColumnIndex: int = -1
    iManhourColumnIndex: int = -1
    if objRows:
        objHeaderRow: List[str] = objRows[0]
        for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
            if pszColumnName == "販売費及び一般管理費計":
                iSellGeneralAdminCostColumnIndex = iColumnIndex
            elif pszColumnName == "配賦販管費":
                iAllocationColumnIndex = iColumnIndex
            elif pszColumnName == "工数":
                iManhourColumnIndex = iColumnIndex

    if iSellGeneralAdminCostColumnIndex >= 0 and iAllocationColumnIndex >= 0 and iManhourColumnIndex >= 0:
        calculate_allocation(
            objRows,
            iSellGeneralAdminCostColumnIndex,
            iAllocationColumnIndex,
            iManhourColumnIndex,
        )

    with open(pszOutputStep0002Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    iGrossProfitColumnIndex: int = -1
    iOperatingProfitColumnIndex: int = -1
    if objRows:
        objHeaderRow = objRows[0]
        for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
            if pszColumnName == "売上総利益":
                iGrossProfitColumnIndex = iColumnIndex
            elif pszColumnName == "営業利益":
                iOperatingProfitColumnIndex = iColumnIndex

    if iGrossProfitColumnIndex >= 0 and iOperatingProfitColumnIndex >= 0:
        recalculate_operating_profit(
            objRows,
            iGrossProfitColumnIndex,
            iOperatingProfitColumnIndex,
        )

    with open(pszOutputStep0003Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    iNonOperatingIncomeColumnIndex: int = -1
    iNonOperatingExpenseColumnIndex: int = -1
    iOrdinaryProfitColumnIndex: int = -1
    if objRows:
        objHeaderRow = objRows[0]
        for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
            if pszColumnName == "営業外収益":
                iNonOperatingIncomeColumnIndex = iColumnIndex
            elif pszColumnName == "営業外費用":
                iNonOperatingExpenseColumnIndex = iColumnIndex
            elif pszColumnName == "経常利益":
                iOrdinaryProfitColumnIndex = iColumnIndex

    if (
        iOperatingProfitColumnIndex >= 0
        and iNonOperatingIncomeColumnIndex >= 0
        and iNonOperatingExpenseColumnIndex >= 0
        and iOrdinaryProfitColumnIndex >= 0
    ):
        recalculate_ordinary_profit(
            objRows,
            iOperatingProfitColumnIndex,
            iNonOperatingIncomeColumnIndex,
            iNonOperatingExpenseColumnIndex,
            iOrdinaryProfitColumnIndex,
        )

    with open(pszOutputStep0004Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    iExtraordinaryIncomeColumnIndex: int = -1
    iExtraordinaryLossColumnIndex: int = -1
    iPreTaxProfitColumnIndex: int = -1
    if objRows:
        objHeaderRow = objRows[0]
        for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
            if pszColumnName == "特別利益":
                iExtraordinaryIncomeColumnIndex = iColumnIndex
            elif pszColumnName == "特別損失":
                iExtraordinaryLossColumnIndex = iColumnIndex
            elif pszColumnName == "税引前当期純利益":
                iPreTaxProfitColumnIndex = iColumnIndex

    if (
        iOrdinaryProfitColumnIndex >= 0
        and iExtraordinaryIncomeColumnIndex >= 0
        and iExtraordinaryLossColumnIndex >= 0
        and iPreTaxProfitColumnIndex >= 0
    ):
        recalculate_pre_tax_profit(
            objRows,
            iOrdinaryProfitColumnIndex,
            iExtraordinaryIncomeColumnIndex,
            iExtraordinaryLossColumnIndex,
            iPreTaxProfitColumnIndex,
        )

    with open(pszOutputStep0005Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    iCorporateTaxColumnIndex: int = -1
    iCorporateTaxTotalColumnIndex: int = -1
    iNetProfitColumnIndex: int = -1
    if objRows:
        objHeaderRow = objRows[0]
        for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
            if pszColumnName == "法人税、住民税及び事業税":
                iCorporateTaxColumnIndex = iColumnIndex
            elif pszColumnName == "法人税等":
                iCorporateTaxTotalColumnIndex = iColumnIndex
            elif pszColumnName == "当期純利益":
                iNetProfitColumnIndex = iColumnIndex

    if (
        iCorporateTaxColumnIndex >= 0
        and iCorporateTaxTotalColumnIndex >= 0
        and iPreTaxProfitColumnIndex >= 0
        and iNetProfitColumnIndex >= 0
    ):
        recalculate_net_profit(
            objRows,
            iCorporateTaxColumnIndex,
            iCorporateTaxTotalColumnIndex,
            iPreTaxProfitColumnIndex,
            iNetProfitColumnIndex,
        )

    with open(pszOutputStep0006Path, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")

    with open(pszOutputFinalPath, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objRows:
            objOutputFile.write("\t".join(objRow) + "\n")
    write_transposed_tsv(pszOutputFinalPath)


def transpose_rows(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []
    iMaxColumns: int = max(len(objRow) for objRow in objRows)
    objNormalized: List[List[str]] = []
    for objRow in objRows:
        objNormalized.append(objRow + [""] * (iMaxColumns - len(objRow)))

    objTransposed: List[List[str]] = []
    for iColumnIndex in range(iMaxColumns):
        objTransposed.append([objRow[iColumnIndex] for objRow in objNormalized])
    return objTransposed


def write_transposed_tsv(pszInputPath: str) -> None:
    pszDirectory: str
    pszFileName: str
    pszDirectory, pszFileName = os.path.split(pszInputPath)
    pszOutputFileName: str = pszFileName.replace("_vertical", "")
    pszOutputPath: str = os.path.join(pszDirectory, pszOutputFileName)

    objRows: List[List[str]] = []
    with open(pszInputPath, "r", encoding="utf-8", newline="") as objInputFile:
        for pszLine in objInputFile:
            pszLineText: str = pszLine.rstrip("\n").rstrip("\r")
            objRows.append(pszLineText.split("\t"))

    objTransposed = transpose_rows(objRows)
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        for objRow in objTransposed:
            objOutputFile.write("\t".join(objRow) + "\n")


def find_selected_range_path(pszBaseDirectory: str) -> Optional[str]:
    pszFileName: str = "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt"
    objCandidates: List[str] = [
        os.path.join(pszBaseDirectory, pszFileName),
        os.path.join(os.path.dirname(__file__), pszFileName),
    ]
    for pszCandidate in objCandidates:
        if os.path.isfile(pszCandidate):
            return pszCandidate
    return None


def parse_selected_range(pszRangePath: str) -> Optional[Tuple[Tuple[int, int], Tuple[int, int]]]:
    try:
        with open(pszRangePath, "r", encoding="utf-8", newline="") as objFile:
            pszLine: str = objFile.readline().strip()
    except OSError:
        return None

    objMatch = re.search(r"採用範囲:\s*(\d{4})年(\d{1,2})月〜(\d{4})年(\d{1,2})月", pszLine)
    if objMatch is None:
        return None
    iStartYear: int = int(objMatch.group(1))
    iStartMonth: int = int(objMatch.group(2))
    iEndYear: int = int(objMatch.group(3))
    iEndMonth: int = int(objMatch.group(4))
    if not (1 <= iStartMonth <= 12 and 1 <= iEndMonth <= 12):
        return None
    return (iStartYear, iStartMonth), (iEndYear, iEndMonth)


def next_year_month(iYear: int, iMonth: int) -> Tuple[int, int]:
    iMonth += 1
    if iMonth > 12:
        iMonth = 1
        iYear += 1
    return iYear, iMonth


def build_month_sequence(
    objStart: Tuple[int, int],
    objEnd: Tuple[int, int],
) -> List[Tuple[int, int]]:
    objMonths: List[Tuple[int, int]] = []
    iYear, iMonth = objStart
    while True:
        objMonths.append((iYear, iMonth))
        if (iYear, iMonth) == objEnd:
            break
        iYear, iMonth = next_year_month(iYear, iMonth)
    return objMonths


def split_by_fiscal_boundary(
    objStart: Tuple[int, int],
    objEnd: Tuple[int, int],
    iBoundaryEndMonth: int,
) -> List[Tuple[Tuple[int, int], Tuple[int, int]]]:
    objMonths = build_month_sequence(objStart, objEnd)
    if not objMonths:
        return []

    objRanges: List[Tuple[Tuple[int, int], Tuple[int, int]]] = []
    objRangeStart: Tuple[int, int] = objMonths[0]
    for iIndex, objMonth in enumerate(objMonths):
        iYear, iMonth = objMonth
        is_last: bool = iIndex == len(objMonths) - 1
        if iMonth == iBoundaryEndMonth and not is_last:
            objRanges.append((objRangeStart, objMonth))
            objRangeStart = objMonths[iIndex + 1]
    objRanges.append((objRangeStart, objMonths[-1]))
    return objRanges


def try_parse_float(pszText: str) -> Optional[float]:
    pszValue: str = (pszText or "").strip()
    if pszValue == "":
        return None
    try:
        return float(pszValue)
    except ValueError:
        return None


def read_tsv_rows(pszPath: str) -> List[List[str]]:
    objRows: List[List[str]] = []
    with open(pszPath, "r", encoding="utf-8", newline="") as objFile:
        for pszLine in objFile:
            pszLineText: str = pszLine.rstrip("\n").rstrip("\r")
            objRows.append(pszLineText.split("\t") if pszLineText != "" else [""])
    return objRows


def sum_tsv_rows(objBaseRows: List[List[str]], objAddRows: List[List[str]]) -> List[List[str]]:
    if not objBaseRows:
        return [list(objRow) for objRow in objAddRows]
    if not objAddRows:
        return objBaseRows

    objBaseKeyIndices: Dict[str, List[int]] = {}
    for iRowIndex, objRow in enumerate(objBaseRows):
        pszKey: str = objRow[0] if objRow else ""
        objBaseKeyIndices.setdefault(pszKey, []).append(iRowIndex)

    objBaseKeyCursor: Dict[str, int] = {pszKey: 0 for pszKey in objBaseKeyIndices}

    for iRowIndex, objAddRow in enumerate(objAddRows):
        if iRowIndex == 0:
            objBaseHeader: List[str] = objBaseRows[0]
            iColumnCount: int = max(len(objBaseHeader), len(objAddRow))
            if len(objBaseHeader) < iColumnCount:
                objBaseHeader.extend([""] * (iColumnCount - len(objBaseHeader)))
            if len(objAddRow) < iColumnCount:
                objAddRow = objAddRow + [""] * (iColumnCount - len(objAddRow))
            for iColumnIndex in range(iColumnCount):
                if objBaseHeader[iColumnIndex].strip() == "" and objAddRow[iColumnIndex].strip() != "":
                    objBaseHeader[iColumnIndex] = objAddRow[iColumnIndex]
            objBaseRows[0] = objBaseHeader
            continue

        pszKey = objAddRow[0] if objAddRow else ""
        objIndices: List[int] = objBaseKeyIndices.get(pszKey, [])
        iCursor: int = objBaseKeyCursor.get(pszKey, 0)
        if iCursor < len(objIndices):
            iTargetIndex = objIndices[iCursor]
            objBaseKeyCursor[pszKey] = iCursor + 1
        else:
            iTargetIndex = len(objBaseRows)
            objBaseRows.append(list(objAddRow))
            objBaseKeyIndices.setdefault(pszKey, []).append(iTargetIndex)
            objBaseKeyCursor[pszKey] = objBaseKeyCursor.get(pszKey, 0) + 1
            continue

        objBaseRow = objBaseRows[iTargetIndex]
        iColumnCount = max(len(objBaseRow), len(objAddRow))
        if len(objBaseRow) < iColumnCount:
            objBaseRow.extend([""] * (iColumnCount - len(objBaseRow)))
        if len(objAddRow) < iColumnCount:
            objAddRow = objAddRow + [""] * (iColumnCount - len(objAddRow))

        for iColumnIndex in range(1, iColumnCount):
            pszBaseValue = objBaseRow[iColumnIndex]
            pszAddValue = objAddRow[iColumnIndex]
            fBase = try_parse_float(pszBaseValue)
            fAdd = try_parse_float(pszAddValue)

            if fBase is not None and fAdd is not None:
                objBaseRow[iColumnIndex] = format_number(fBase + fAdd)
            elif fBase is None and pszBaseValue.strip() == "" and fAdd is not None:
                objBaseRow[iColumnIndex] = format_number(fAdd)
            elif fBase is None and pszBaseValue.strip() == "" and pszAddValue.strip() != "":
                objBaseRow[iColumnIndex] = pszAddValue

        objBaseRows[iTargetIndex] = objBaseRow

    return objBaseRows


def write_tsv_rows(pszPath: str, objRows: List[List[str]]) -> None:
    with open(pszPath, "w", encoding="utf-8", newline="") as objFile:
        for objRow in objRows:
            objFile.write("\t".join(objRow) + "\n")


def append_gross_margin_column(objRows: List[List[str]]) -> List[List[str]]:
    if not objRows:
        return []
    objHeader: List[str] = objRows[0]
    iSalesIndex: int = find_column_index(objHeader, "純売上高")
    iGrossProfitIndex: int = find_column_index(objHeader, "売上総利益")
    objOutputRows: List[List[str]] = []

    for iRowIndex, objRow in enumerate(objRows):
        objNewRow: List[str] = list(objRow)
        if iRowIndex == 0:
            objNewRow.append("粗利益率")
            objOutputRows.append(objNewRow)
            continue

        fSales: float = 0.0
        fGrossProfit: float = 0.0
        if 0 <= iSalesIndex < len(objRow):
            fSales = parse_number(objRow[iSalesIndex])
        if 0 <= iGrossProfitIndex < len(objRow):
            fGrossProfit = parse_number(objRow[iGrossProfitIndex])

        if abs(fSales) < 0.0000001:
            if fGrossProfit > 0:
                objNewRow.append("'＋∞")
            elif fGrossProfit < 0:
                objNewRow.append("'－∞")
            else:
                objNewRow.append("0")
        else:
            objNewRow.append(format_number(fGrossProfit / fSales))

        objOutputRows.append(objNewRow)
    return objOutputRows


def build_report_file_path(
    pszDirectory: str,
    pszPrefix: str,
    objYearMonth: Tuple[int, int],
) -> str:
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    pszFileName: str = f"{pszPrefix}_{iYear}年{pszMonth}月_A∪B_プロジェクト名_C∪D.tsv"
    return os.path.join(pszDirectory, pszFileName)


def build_report_vertical_file_path(
    pszDirectory: str,
    pszPrefix: str,
    objYearMonth: Tuple[int, int],
) -> str:
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    pszFileName: str = f"{pszPrefix}_{iYear}年{pszMonth}月_A∪B_プロジェクト名_C∪D_vertical.tsv"
    return os.path.join(pszDirectory, pszFileName)


def build_cumulative_file_path(
    pszDirectory: str,
    pszPrefix: str,
    objStart: Tuple[int, int],
    objEnd: Tuple[int, int],
) -> str:
    iStartYear, iStartMonth = objStart
    iEndYear, iEndMonth = objEnd
    pszStartMonth: str = f"{iStartMonth:02d}"
    pszEndMonth: str = f"{iEndMonth:02d}"
    pszFileName: str = (
        f"累計_{pszPrefix}_{iStartYear}年{pszStartMonth}月_{iEndYear}年{pszEndMonth}月.tsv"
    )
    return os.path.join(pszDirectory, pszFileName)


def read_report_rows(
    pszDirectory: str,
    pszPrefix: str,
    objYearMonth: Tuple[int, int],
) -> Optional[List[List[str]]]:
    pszHorizontalPath: str = build_report_file_path(pszDirectory, pszPrefix, objYearMonth)
    if os.path.isfile(pszHorizontalPath):
        return read_tsv_rows(pszHorizontalPath)

    pszVerticalPath: str = build_report_vertical_file_path(pszDirectory, pszPrefix, objYearMonth)
    if os.path.isfile(pszVerticalPath):
        objVerticalRows: List[List[str]] = read_tsv_rows(pszVerticalPath)
        return transpose_rows(objVerticalRows)

    print(f"Input file not found: {pszHorizontalPath}")
    print(f"Input file not found: {pszVerticalPath}")
    return None


def find_column_index(objHeader: List[str], pszName: str) -> int:
    for iIndex, pszValue in enumerate(objHeader):
        if pszValue == pszName:
            return iIndex
    return -1


def is_company_project(pszProjectName: str) -> bool:
    return re.match(r"^C\d{3}_", pszProjectName) is not None


def is_summary_project(pszProjectName: str) -> bool:
    return pszProjectName.startswith("合計")


def is_project_code(pszProjectName: str, pszPrefix: str, iDigits: int) -> bool:
    return re.match(rf"^{pszPrefix}\d{{{iDigits}}}_", pszProjectName) is not None


def collect_project_rows(
    objRows: List[List[str]],
    iProjectNameColumnIndex: int,
) -> List[List[str]]:
    if not objRows:
        return []
    iStartIndex: int = 1
    for iIndex, objRow in enumerate(objRows[1:], start=1):
        pszName: str = ""
        if 0 <= iProjectNameColumnIndex < len(objRow):
            pszName = objRow[iProjectNameColumnIndex]
        if pszName.startswith("本部"):
            iStartIndex = iIndex
            break
    return objRows[iStartIndex:]


def build_project_rows_for_summary(
    objRows: List[List[str]],
    iProjectNameColumnIndex: int,
) -> List[List[str]]:
    objCandidateRows: List[List[str]] = collect_project_rows(objRows, iProjectNameColumnIndex)
    objOrderedRows: List[List[str]] = []
    objRules: List[Tuple[str, int]] = [("J", 3), ("P", 5)]
    for pszPrefix, iDigits in objRules:
        for objRow in objCandidateRows:
            pszName: str = ""
            if 0 <= iProjectNameColumnIndex < len(objRow):
                pszName = objRow[iProjectNameColumnIndex]
            if pszName == "" or is_company_project(pszName) or is_summary_project(pszName):
                continue
            if is_project_code(pszName, pszPrefix, iDigits):
                objOrderedRows.append(objRow)
    return objOrderedRows


def extract_project_values(
    objRows: List[List[str]],
    iProjectNameColumnIndex: int,
    iValueColumnIndex: int,
) -> List[str]:
    objValues: List[str] = []
    for objRow in build_project_rows_for_summary(objRows, iProjectNameColumnIndex):
        if iValueColumnIndex < 0 or iValueColumnIndex >= len(objRow):
            objValues.append("")
        else:
            objValues.append(objRow[iValueColumnIndex])
    return objValues


def extract_project_names(
    objRows: List[List[str]],
    iProjectNameColumnIndex: int,
) -> List[str]:
    objNames: List[str] = []
    for objRow in build_project_rows_for_summary(objRows, iProjectNameColumnIndex):
        pszName: str = ""
        if 0 <= iProjectNameColumnIndex < len(objRow):
            pszName = objRow[iProjectNameColumnIndex]
        objNames.append(pszName)
    return objNames


def build_gross_margin_values(
    objRows: List[List[str]],
    iProjectNameColumnIndex: int,
    iGrossProfitColumnIndex: int,
    iSalesColumnIndex: int,
) -> List[str]:
    objValues: List[str] = []
    for objRow in build_project_rows_for_summary(objRows, iProjectNameColumnIndex):
        fGrossProfit: float = 0.0
        fSales: float = 0.0
        if 0 <= iGrossProfitColumnIndex < len(objRow):
            fGrossProfit = parse_number(objRow[iGrossProfitColumnIndex])
        if 0 <= iSalesColumnIndex < len(objRow):
            fSales = parse_number(objRow[iSalesColumnIndex])
        if abs(fSales) < 0.0000001:
            if fGrossProfit > 0:
                objValues.append("'＋∞")
            elif fGrossProfit < 0:
                objValues.append("'－∞")
            else:
                objValues.append("0")
        else:
            objValues.append(format_number(fGrossProfit / fSales))
    return objValues


def write_pj_summary(
    pszOutputPath: str,
    objSingleRows: List[List[str]],
    objCumulativeRows: List[List[str]],
) -> None:
    if not objSingleRows or not objCumulativeRows:
        return
    objSingleHeader: List[str] = objSingleRows[0]
    objCumulativeHeader: List[str] = objCumulativeRows[0]

    iProjectNameColumnIndex: int = 0
    iSingleSalesIndex: int = 2
    iCumulativeSalesIndex: int = 2
    iSingleCostIndex: int = find_column_index(objSingleHeader, "売上原価")
    iCumulativeCostIndex: int = find_column_index(objCumulativeHeader, "売上原価")
    iSingleGrossIndex: int = find_column_index(objSingleHeader, "売上総利益")
    iCumulativeGrossIndex: int = find_column_index(objCumulativeHeader, "売上総利益")
    iSingleAllocationIndex: int = find_column_index(objSingleHeader, "配賦販管費")
    iCumulativeAllocationIndex: int = find_column_index(objCumulativeHeader, "配賦販管費")

    objProjectNames: List[str] = extract_project_names(objSingleRows, iProjectNameColumnIndex)
    objSingleSales: List[str] = extract_project_values(
        objSingleRows,
        iProjectNameColumnIndex,
        iSingleSalesIndex,
    )
    objCumulativeSales: List[str] = extract_project_values(
        objCumulativeRows,
        iProjectNameColumnIndex,
        iCumulativeSalesIndex,
    )
    objSingleCost: List[str] = extract_project_values(
        objSingleRows,
        iProjectNameColumnIndex,
        iSingleCostIndex,
    )
    objCumulativeCost: List[str] = extract_project_values(
        objCumulativeRows,
        iProjectNameColumnIndex,
        iCumulativeCostIndex,
    )
    objSingleGross: List[str] = extract_project_values(
        objSingleRows,
        iProjectNameColumnIndex,
        iSingleGrossIndex,
    )
    objCumulativeGross: List[str] = extract_project_values(
        objCumulativeRows,
        iProjectNameColumnIndex,
        iCumulativeGrossIndex,
    )
    objSingleAllocation: List[str] = extract_project_values(
        objSingleRows,
        iProjectNameColumnIndex,
        iSingleAllocationIndex,
    )
    objCumulativeAllocation: List[str] = extract_project_values(
        objCumulativeRows,
        iProjectNameColumnIndex,
        iCumulativeAllocationIndex,
    )
    objSingleMargin: List[str] = build_gross_margin_values(
        objSingleRows,
        iProjectNameColumnIndex,
        iSingleGrossIndex,
        iSingleSalesIndex,
    )
    objCumulativeMargin: List[str] = build_gross_margin_values(
        objCumulativeRows,
        iProjectNameColumnIndex,
        iCumulativeGrossIndex,
        iCumulativeSalesIndex,
    )

    objRows: List[List[str]] = []
    objRows.append(
        [
            "1",
            "PJ名称",
            "単月_純売上高",
            "累計_純売上高",
            "単月_売上原価",
            "累計_売上原価",
            "単月_売上総利益",
            "累計_売上総利益",
            "単月_カンパニー販管費",
            "累計_カンパニー販管費",
            "単月_配賦販管費",
            "累計_配賦販管費",
            "粗利益率",
            "粗利益率",
        ]
    )

    for iIndex, pszProjectName in enumerate(objProjectNames):
        objRows.append(
            [
                str(iIndex + 2),
                pszProjectName,
                objSingleSales[iIndex] if iIndex < len(objSingleSales) else "",
                objCumulativeSales[iIndex] if iIndex < len(objCumulativeSales) else "",
                objSingleCost[iIndex] if iIndex < len(objSingleCost) else "",
                objCumulativeCost[iIndex] if iIndex < len(objCumulativeCost) else "",
                objSingleGross[iIndex] if iIndex < len(objSingleGross) else "",
                objCumulativeGross[iIndex] if iIndex < len(objCumulativeGross) else "",
                "0",
                "0",
                objSingleAllocation[iIndex] if iIndex < len(objSingleAllocation) else "",
                objCumulativeAllocation[iIndex] if iIndex < len(objCumulativeAllocation) else "",
                objSingleMargin[iIndex] if iIndex < len(objSingleMargin) else "",
                objCumulativeMargin[iIndex] if iIndex < len(objCumulativeMargin) else "",
            ]
        )

    write_tsv_rows(pszOutputPath, objRows)


def filter_rows_by_columns(
    objRows: List[List[str]],
    objTargetColumns: List[str],
) -> List[List[str]]:
    if not objRows:
        return []
    objHeader: List[str] = objRows[0]
    objColumnIndices: List[int] = [
        find_column_index(objHeader, pszColumn)
        for pszColumn in objTargetColumns
    ]
    objFilteredRows: List[List[str]] = []
    for objRow in objRows:
        objFilteredRow: List[str] = []
        for iColumnIndex in objColumnIndices:
            if 0 <= iColumnIndex < len(objRow):
                objFilteredRow.append(objRow[iColumnIndex])
            else:
                objFilteredRow.append("")
        objFilteredRows.append(objFilteredRow)
    return objFilteredRows


def filter_rows_by_names(
    objRows: List[List[str]],
    objTargetNames: List[str],
) -> List[List[str]]:
    if not objRows:
        return []
    objTargetSet = set(objTargetNames)
    objFilteredRows: List[List[str]] = []
    for objRow in objRows:
        if not objRow:
            continue
        if objRow[0] in objTargetSet:
            objFilteredRows.append(objRow)
    return objFilteredRows


def create_pj_summary(
    pszPlPath: str,
    objRange: Tuple[Tuple[int, int], Tuple[int, int]],
) -> None:
    objStart, objEnd = objRange
    pszDirectory: str = os.path.dirname(pszPlPath)
    iEndYear, iEndMonth = objEnd
    pszEndMonth: str = f"{iEndMonth:02d}"
    pszSinglePlPath: str = os.path.join(
        pszDirectory,
        f"損益計算書_販管費配賦_{iEndYear}年{pszEndMonth}月_A∪B_プロジェクト名_C∪D_vertical.tsv",
    )
    pszCumulativePlPath: str = build_cumulative_file_path(
        pszDirectory,
        "損益計算書",
        objStart,
        objEnd,
    ).replace(".tsv", "_vertical.tsv")

    if not os.path.isfile(pszSinglePlPath) or not os.path.isfile(pszCumulativePlPath):
        return

    objSingleRows: List[List[str]] = read_tsv_rows(pszSinglePlPath)
    objCumulativeRows: List[List[str]] = read_tsv_rows(pszCumulativePlPath)

    objSingleOutputRows: List[List[str]] = []
    for objRow in objSingleRows:
        pszName: str = objRow[0] if objRow else ""
        if pszName == "合計" or pszName.startswith("C"):
            continue
        objSingleOutputRows.append(objRow)

    objCumulativeOutputRows: List[List[str]] = []
    for objRow in objCumulativeRows:
        pszName: str = objRow[0] if objRow else ""
        if pszName == "合計" or pszName.startswith("C"):
            continue
        objCumulativeOutputRows.append(objRow)

    pszSingleOutputPath: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0001_単月_損益計算書.tsv",
    )
    pszCumulativeOutputPath: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0001_累計_損益計算書.tsv",
    )
    write_tsv_rows(pszSingleOutputPath, objSingleOutputRows)
    write_tsv_rows(pszCumulativeOutputPath, objCumulativeOutputRows)

    pszSingleCostReportPath: str = os.path.join(
        pszDirectory,
        f"製造原価報告書_{iEndYear}年{pszEndMonth}月_A∪B_プロジェクト名_C∪D.tsv",
    )
    pszCumulativeCostReportPath: str = build_cumulative_file_path(
        pszDirectory,
        "製造原価報告書",
        objStart,
        objEnd,
    )
    if os.path.isfile(pszSingleCostReportPath):
        objCostReportSingleRows: List[List[str]] = read_tsv_rows(pszSingleCostReportPath)
        pszCostReportSingleOutputPath: str = os.path.join(
            pszDirectory,
            "0001_PJサマリ_step0001_単月_製造原価報告書.tsv",
        )
        write_tsv_rows(pszCostReportSingleOutputPath, objCostReportSingleRows)
    if os.path.isfile(pszCumulativeCostReportPath):
        objCostReportCumulativeRows: List[List[str]] = read_tsv_rows(pszCumulativeCostReportPath)
        pszCostReportCumulativeOutputPath: str = os.path.join(
            pszDirectory,
            "0001_PJサマリ_step0001_累計_製造原価報告書.tsv",
        )
        write_tsv_rows(pszCostReportCumulativeOutputPath, objCostReportCumulativeRows)

    objSingleOutputVerticalRows = transpose_rows(objSingleOutputRows)
    pszSingleOutputVerticalPath: str = os.path.join(
        pszDirectory,
        "0003_PJサマリ_step0001_単月_損益計算書.tsv",
    )
    write_tsv_rows(pszSingleOutputVerticalPath, objSingleOutputVerticalRows)

    objCumulativeOutputVerticalRows = transpose_rows(objCumulativeOutputRows)
    pszCumulativeOutputVerticalPath: str = os.path.join(
        pszDirectory,
        "0003_PJサマリ_step0001_累計_損益計算書.tsv",
    )
    write_tsv_rows(pszCumulativeOutputVerticalPath, objCumulativeOutputVerticalRows)

    if os.path.isfile(pszSingleCostReportPath):
        pszCostReportSingleOutputPath: str = os.path.join(
            pszDirectory,
            "0003_PJサマリ_step0001_単月_製造原価報告書.tsv",
        )
        shutil.copy2(pszSingleCostReportPath, pszCostReportSingleOutputPath)
    if os.path.isfile(pszCumulativeCostReportPath):
        pszCostReportCumulativeOutputPath: str = os.path.join(
            pszDirectory,
            "0003_PJサマリ_step0001_累計_製造原価報告書.tsv",
        )
        shutil.copy2(pszCumulativeCostReportPath, pszCostReportCumulativeOutputPath)

    objTargetNames: List[str] = [
        "科目名",
        "純売上高",
        "売上総利益",
        "配賦販管費",
        "営業利益",
    ]
    objSingleStep0002Rows = filter_rows_by_names(
        objSingleOutputVerticalRows,
        objTargetNames,
    )
    objCumulativeStep0002Rows = filter_rows_by_names(
        objCumulativeOutputVerticalRows,
        objTargetNames,
    )
    pszSingleStep0002Path: str = os.path.join(
        pszDirectory,
        "0003_PJサマリ_step0002_単月_損益計算書.tsv",
    )
    pszCumulativeStep0002Path: str = os.path.join(
        pszDirectory,
        "0003_PJサマリ_step0002_累計_損益計算書.tsv",
    )
    write_tsv_rows(pszSingleStep0002Path, objSingleStep0002Rows)
    write_tsv_rows(pszCumulativeStep0002Path, objCumulativeStep0002Rows)

    if os.path.isfile(pszSingleCostReportPath):
        pszCostReportSingleStep0002Path: str = os.path.join(
            pszDirectory,
            "0003_PJサマリ_step0002_単月_製造原価報告書.tsv",
        )
        shutil.copy2(pszSingleCostReportPath, pszCostReportSingleStep0002Path)
    if os.path.isfile(pszCumulativeCostReportPath):
        pszCostReportCumulativeStep0002Path: str = os.path.join(
            pszDirectory,
            "0003_PJサマリ_step0002_累計_製造原価報告書.tsv",
        )
        shutil.copy2(pszCumulativeCostReportPath, pszCostReportCumulativeStep0002Path)

    objTargetColumns: List[str] = [
        "科目名",
        "純売上高",
        "売上原価",
        "売上総利益",
        "配賦販管費",
    ]
    objSingleStep0002Rows: List[List[str]] = filter_rows_by_columns(
        objSingleOutputRows,
        objTargetColumns,
    )
    objCumulativeStep0002Rows: List[List[str]] = filter_rows_by_columns(
        objCumulativeOutputRows,
        objTargetColumns,
    )
    pszSingleStep0002Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0002_単月_損益計算書.tsv",
    )
    pszCumulativeStep0002Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0002_累計_損益計算書.tsv",
    )
    write_tsv_rows(pszSingleStep0002Path, objSingleStep0002Rows)
    write_tsv_rows(pszCumulativeStep0002Path, objCumulativeStep0002Rows)

    objSingleStep0003Rows: List[List[str]] = append_gross_margin_column(objSingleStep0002Rows)
    objCumulativeStep0003Rows: List[List[str]] = append_gross_margin_column(
        objCumulativeStep0002Rows
    )
    pszSingleStep0003Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0003_単月_損益計算書.tsv",
    )
    pszCumulativeStep0003Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0003_累計_損益計算書.tsv",
    )
    write_tsv_rows(pszSingleStep0003Path, objSingleStep0003Rows)
    write_tsv_rows(pszCumulativeStep0003Path, objCumulativeStep0003Rows)

    if len(objSingleStep0003Rows) != len(objCumulativeStep0003Rows):
        print("Error: step0003 row count mismatch between single and cumulative.")
        return

    for iRowIndex, objRow in enumerate(objSingleStep0003Rows):
        pszSingleKey: str = objRow[0] if objRow else ""
        objCumulativeRow: List[str] = objCumulativeStep0003Rows[iRowIndex]
        pszCumulativeKey: str = objCumulativeRow[0] if objCumulativeRow else ""
        if pszSingleKey != pszCumulativeKey:
            print(
                "Error: step0003 first-column mismatch at row "
                + str(iRowIndex)
                + ". single="
                + pszSingleKey
                + " cumulative="
                + pszCumulativeKey
            )
            return

    objStep0004Rows: List[List[str]] = []
    for iRowIndex, objRow in enumerate(objSingleStep0003Rows):
        objCumulativeRow = objCumulativeStep0003Rows[iRowIndex]
        if iRowIndex == 0:
            objHeader: List[str] = [objRow[0] if objRow else ""]
            iMaxColumns: int = max(len(objRow), len(objCumulativeRow))
            for iColumnIndex in range(1, iMaxColumns):
                pszSingleHeader: str = objRow[iColumnIndex] if iColumnIndex < len(objRow) else ""
                pszCumulativeHeader: str = (
                    objCumulativeRow[iColumnIndex] if iColumnIndex < len(objCumulativeRow) else ""
                )
                objHeader.append(pszSingleHeader)
                objHeader.append(pszCumulativeHeader)
            objStep0004Rows.append(objHeader)
            continue

        objOutputRow: List[str] = [objRow[0] if objRow else ""]
        iMaxColumns = max(len(objRow), len(objCumulativeRow))
        for iColumnIndex in range(1, iMaxColumns):
            objOutputRow.append(objRow[iColumnIndex] if iColumnIndex < len(objRow) else "")
            objOutputRow.append(
                objCumulativeRow[iColumnIndex] if iColumnIndex < len(objCumulativeRow) else ""
            )
        objStep0004Rows.append(objOutputRow)

    pszStep0004Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0004_単月・累計_損益計算書.tsv",
    )
    write_tsv_rows(pszStep0004Path, objStep0004Rows)

    if not objStep0004Rows:
        return

    iStep0004ColumnCount: int = max(len(objRow) for objRow in objStep0004Rows)
    objStep0005Rows: List[List[str]] = []
    objHeaderRow: List[str] = ["単／累"]
    for iColumnIndex in range(1, iStep0004ColumnCount):
        if iColumnIndex % 2 == 1:
            objHeaderRow.append("単月")
        else:
            objHeaderRow.append("累計")
    objStep0005Rows.append(objHeaderRow)
    objStep0005Rows.extend(objStep0004Rows)

    pszStep0005Path: str = os.path.join(
        pszDirectory,
        "0001_PJサマリ_step0005_単月・累計_損益計算書.tsv",
    )
    write_tsv_rows(pszStep0005Path, objStep0005Rows)

    objGrossProfitColumns: List[str] = ["科目名", "売上総利益", "純売上高"]
    objGrossProfitSingleRows: List[List[str]] = filter_rows_by_columns(
        objSingleOutputRows,
        objGrossProfitColumns,
    )
    objGrossProfitCumulativeRows: List[List[str]] = filter_rows_by_columns(
        objCumulativeOutputRows,
        objGrossProfitColumns,
    )
    pszGrossProfitSinglePath: str = os.path.join(
        pszDirectory,
        "0002_PJサマリ_step0001_単月_粗利金額ランキング.tsv",
    )
    pszGrossProfitCumulativePath: str = os.path.join(
        pszDirectory,
        "0002_PJサマリ_step0001_累計_粗利金額ランキング.tsv",
    )
    write_tsv_rows(pszGrossProfitSinglePath, objGrossProfitSingleRows)
    write_tsv_rows(pszGrossProfitCumulativePath, objGrossProfitCumulativeRows)

    # 単月_粗利金額ランキング
    objGrossProfitSingleSortedRows: List[List[str]] = []
    objGrossProfitCumulativeSortedRows: List[List[str]] = []

    if objGrossProfitSingleRows:
        objSingleHeader: List[str] = objGrossProfitSingleRows[0]
        objSingleBody: List[List[str]] = objGrossProfitSingleRows[1:]
        objSingleBody.sort(
            key=lambda objRow: try_parse_float(objRow[1] if len(objRow) > 1 else "") or 0.0,
            reverse=True,
        )
        objGrossProfitSingleSortedRows = [objSingleHeader] + objSingleBody
        pszGrossProfitSingleSortedPath: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0002_単月_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitSingleSortedPath, objGrossProfitSingleSortedRows)

    # 累計_粗利金額ランキング
    if objGrossProfitCumulativeRows:
        objCumulativeHeader: List[str] = objGrossProfitCumulativeRows[0]
        objCumulativeBody: List[List[str]] = objGrossProfitCumulativeRows[1:]
        objCumulativeBody.sort(
            key=lambda objRow: try_parse_float(objRow[1] if len(objRow) > 1 else "") or 0.0,
            reverse=True,
        )
        objGrossProfitCumulativeSortedRows = [objCumulativeHeader] + objCumulativeBody
        pszGrossProfitCumulativeSortedPath: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0002_累計_粗利金額ランキング.tsv",
        )
        write_tsv_rows(
            pszGrossProfitCumulativeSortedPath,
            objGrossProfitCumulativeSortedRows,
        )

    if objGrossProfitSingleSortedRows and objGrossProfitCumulativeSortedRows:
        if len(objGrossProfitSingleSortedRows) != len(objGrossProfitCumulativeSortedRows):
            print("Error: gross profit ranking row count mismatch.")
            return

        objGrossProfitCombinedRows = [list(objRow) for objRow in objGrossProfitSingleSortedRows]
        for objRow in objGrossProfitCombinedRows:
            objRow.append("")

        for iRowIndex, objCumulativeRow in enumerate(objGrossProfitCumulativeSortedRows):
            if len(objGrossProfitCombinedRows[iRowIndex]) < 3:
                objGrossProfitCombinedRows[iRowIndex].extend(
                    [""] * (3 - len(objGrossProfitCombinedRows[iRowIndex]))
                )
            pszCumulativeProject: str = objCumulativeRow[0] if objCumulativeRow else ""
            pszCumulativeValue: str = objCumulativeRow[1] if len(objCumulativeRow) > 1 else ""
            objGrossProfitCombinedRows[iRowIndex].append(pszCumulativeProject)
            objGrossProfitCombinedRows[iRowIndex].append(pszCumulativeValue)

        pszGrossProfitCombinedPath: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0003_単月・累計_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitCombinedPath, objGrossProfitCombinedRows)

    if objGrossProfitSingleSortedRows:
        objGrossProfitSingleRankRows: List[List[str]] = []
        for iRowIndex, objRow in enumerate(objGrossProfitSingleSortedRows):
            if iRowIndex == 0:
                objGrossProfitSingleRankRows.append(
                    ["0", "プロジェクト名", "売上総利益", "純売上高", "利益率"]
                )
                continue
            pszGrossProfit: str = objRow[1] if len(objRow) > 1 else ""
            pszSales: str = objRow[2] if len(objRow) > 2 else ""
            fGrossProfit: float = parse_number(pszGrossProfit)
            fSales: float = parse_number(pszSales)
            if abs(fSales) < 0.0000001:
                if fGrossProfit > 0:
                    pszMargin = "'＋∞"
                elif fGrossProfit < 0:
                    pszMargin = "'－∞"
                else:
                    pszMargin = "0"
            else:
                pszMargin = format_number(fGrossProfit / fSales)
            objGrossProfitSingleRankRows.append(
                [
                    str(iRowIndex),
                    objRow[0] if objRow else "",
                    pszGrossProfit,
                    pszSales,
                    pszMargin,
                ]
            )
        pszGrossProfitSingleRankPath: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0003_単月_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitSingleRankPath, objGrossProfitSingleRankRows)

    if objGrossProfitCumulativeSortedRows:
        objGrossProfitCumulativeRankRows: List[List[str]] = []
        for iRowIndex, objRow in enumerate(objGrossProfitCumulativeSortedRows):
            if iRowIndex == 0:
                objGrossProfitCumulativeRankRows.append(
                    ["0", "プロジェクト名", "売上総利益", "純売上高", "利益率"]
                )
                continue
            pszGrossProfit = objRow[1] if len(objRow) > 1 else ""
            pszSales = objRow[2] if len(objRow) > 2 else ""
            fGrossProfit = parse_number(pszGrossProfit)
            fSales = parse_number(pszSales)
            if abs(fSales) < 0.0000001:
                if fGrossProfit > 0:
                    pszMargin = "'＋∞"
                elif fGrossProfit < 0:
                    pszMargin = "'－∞"
                else:
                    pszMargin = "0"
            else:
                pszMargin = format_number(fGrossProfit / fSales)
            objGrossProfitCumulativeRankRows.append(
                [
                    str(iRowIndex),
                    objRow[0] if objRow else "",
                    pszGrossProfit,
                    pszSales,
                    pszMargin,
                ]
            )
        pszGrossProfitCumulativeRankPath: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0003_累計_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitCumulativeRankPath, objGrossProfitCumulativeRankRows)

    if objGrossProfitSingleRankRows and objGrossProfitCumulativeRankRows:
        if len(objGrossProfitSingleRankRows) != len(objGrossProfitCumulativeRankRows):
            print("Error: gross profit ranking step0004 row count mismatch.")
            return

        objGrossProfitStep0004Rows: List[List[str]] = []
        for iRowIndex, objSingleRow in enumerate(objGrossProfitSingleRankRows):
            objCumulativeRow = objGrossProfitCumulativeRankRows[iRowIndex]
            objOutputRow: List[str] = list(objSingleRow)
            objOutputRow.append("")
            objOutputRow.extend(objCumulativeRow)
            objGrossProfitStep0004Rows.append(objOutputRow)

        pszGrossProfitStep0004Path: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0004_単月・累計_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitStep0004Path, objGrossProfitStep0004Rows)

        objGrossProfitStep0005Rows: List[List[str]] = [
            [
                "",
                "粗利金額ランキング",
                "粗利金額",
                "",
                "単月",
                "",
                "",
                "粗利金額ランキング",
                "粗利金額",
                "",
                "累計",
            ]
        ]
        objGrossProfitStep0005Rows.extend(objGrossProfitStep0004Rows)
        pszGrossProfitStep0005Path: str = os.path.join(
            pszDirectory,
            "0002_PJサマリ_step0005_単月・累計_粗利金額ランキング.tsv",
        )
        write_tsv_rows(pszGrossProfitStep0005Path, objGrossProfitStep0005Rows)

        if objGrossProfitStep0005Rows:
            objGrossProfitStep0006Rows: List[List[str]] = []
            objSalesIndices: List[int] = []
            objNumberColumnIndices: List[int] = []
            objHeaderRow: List[str] = objGrossProfitStep0005Rows[0]
            for iColumnIndex, pszColumnName in enumerate(objHeaderRow):
                if pszColumnName == "純売上高":
                    objSalesIndices.append(iColumnIndex)

            if objGrossProfitSingleRankRows:
                iSingleColumnCount = len(objGrossProfitSingleRankRows[0])
                objNumberColumnIndices = [0, iSingleColumnCount + 1]

            objSalesIndexSet = set(objSalesIndices)
            objNumberIndexSet = set(objNumberColumnIndices)
            for objRow in objGrossProfitStep0005Rows:
                objFilteredRow: List[str] = []
                for iColumnIndex, pszValue in enumerate(objRow):
                    if iColumnIndex in objSalesIndexSet:
                        continue
                    if iColumnIndex in objNumberIndexSet and pszValue == "0":
                        objFilteredRow.append("")
                    else:
                        objFilteredRow.append(pszValue)
                objGrossProfitStep0006Rows.append(objFilteredRow)

            pszGrossProfitStep0006Path: str = os.path.join(
                pszDirectory,
                "0002_PJサマリ_step0006_単月・累計_粗利金額ランキング.tsv",
            )
            write_tsv_rows(pszGrossProfitStep0006Path, objGrossProfitStep0006Rows)

            pszGrossProfitFinalPath: str = os.path.join(
                pszDirectory,
                "0002_PJサマリ_単月・累計_粗利金額ランキング.tsv",
            )
            write_tsv_rows(pszGrossProfitFinalPath, objGrossProfitStep0006Rows)


def create_cumulative_report(
    pszDirectory: str,
    pszPrefix: str,
    objRange: Tuple[Tuple[int, int], Tuple[int, int]],
    pszInputPrefix: Optional[str] = None,
) -> None:
    objStart, objEnd = objRange
    objMonths = build_month_sequence(objStart, objEnd)
    if not objMonths:
        return

    if pszInputPrefix is None:
        pszInputPrefix = pszPrefix

    objTotalRows: Optional[List[List[str]]] = None
    for objMonth in objMonths:
        objRows: Optional[List[List[str]]] = read_report_rows(
            pszDirectory,
            pszInputPrefix,
            objMonth,
        )
        if objRows is None:
            return
        if objTotalRows is None:
            objTotalRows = objRows
        else:
            objTotalRows = sum_tsv_rows(objTotalRows, objRows)

    if objTotalRows is None:
        return
    pszOutputPath: str = build_cumulative_file_path(pszDirectory, pszPrefix, objStart, objEnd)
    write_tsv_rows(pszOutputPath, objTotalRows)
    print(f"Output: {pszOutputPath}")
    pszVerticalOutputPath: str = pszOutputPath.replace(".tsv", "_vertical.tsv")
    objVerticalRows: List[List[str]] = transpose_rows(objTotalRows)
    write_tsv_rows(pszVerticalOutputPath, objVerticalRows)
    print(f"Output: {pszVerticalOutputPath}")


def build_pj_summary_range(
    objRange: Tuple[Tuple[int, int], Tuple[int, int]],
) -> Tuple[Tuple[int, int], Tuple[int, int]]:
    _, objEnd = objRange
    iEndYear, iEndMonth = objEnd
    if iEndMonth >= 4:
        iStartYear: int = iEndYear
    else:
        iStartYear = iEndYear - 1
    return (iStartYear, 4), (iEndYear, iEndMonth)


def create_cumulative_reports(pszPlPath: str) -> None:
    pszDirectory: str = os.path.dirname(pszPlPath)
    pszRangePath: Optional[str] = find_selected_range_path(pszDirectory)
    if pszRangePath is None:
        return

    objRange = parse_selected_range(pszRangePath)
    if objRange is None:
        return

    objStart, objEnd = objRange
    objFiscalARanges = split_by_fiscal_boundary(objStart, objEnd, 3)
    objFiscalBRanges = split_by_fiscal_boundary(objStart, objEnd, 8)
    objAllRanges = objFiscalARanges + objFiscalBRanges

    for objRangeItem in objAllRanges:
        create_cumulative_report(
            pszDirectory,
            "損益計算書",
            objRangeItem,
            pszInputPrefix="損益計算書_販管費配賦",
        )
        create_cumulative_report(pszDirectory, "製造原価報告書", objRangeItem)
    objPjSummaryRange = build_pj_summary_range(objRange)
    create_pj_summary(pszPlPath, objPjSummaryRange)


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print_usage()
        return 1

    if len(argv) == 4:
        objPairs: List[List[str]] = [[argv[1], argv[2], argv[3]]]
    else:
        iArgCount: int = len(argv) - 1
        if iArgCount % 2 != 0:
            print_usage()
            return 1
        objPairs = []
        objManhourCandidates: List[str] = []
        objPlCandidates: List[str] = []
        bGroupedOrder: bool = True
        bSeenPl: bool = False
        for pszCandidate in argv[1:]:
            pszBaseName: str = os.path.basename(pszCandidate)
            if pszBaseName.startswith("工数_"):
                if bSeenPl:
                    bGroupedOrder = False
                objManhourCandidates.append(pszCandidate)
                continue
            if pszBaseName.startswith("損益計算書_"):
                bSeenPl = True
                objPlCandidates.append(pszCandidate)
                continue
            bGroupedOrder = False
            break

        bSplitByGroup: bool = (
            bGroupedOrder
            and len(objManhourCandidates) == len(objPlCandidates)
            and len(objManhourCandidates) + len(objPlCandidates) == iArgCount
        )

        if bSplitByGroup:
            for iIndex in range(len(objManhourCandidates)):
                objPairs.append([objManhourCandidates[iIndex], objPlCandidates[iIndex]])
        else:
            for iIndex in range(1, len(argv), 2):
                if iIndex + 1 >= len(argv):
                    print_usage()
                    return 1
                objPairs.append([argv[iIndex], argv[iIndex + 1]])

    for objPair in objPairs:
        pszManhourPath: str = objPair[0]
        pszPlPath: str = objPair[1]
        pszOutputPath: str
        if len(objPair) == 3:
            pszOutputPath = objPair[2]
        else:
            pszOutputPath = build_default_output_path(pszPlPath)
        pszOutputFinalPath: str = build_output_path_with_step(pszPlPath, "販管費配賦_")
        pszOutputStep0001Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0001_")
        pszOutputStep0002Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0002_")
        pszOutputStep0003Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0003_")
        pszOutputStep0004Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0004_")
        pszOutputStep0005Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0005_")
        pszOutputStep0006Path: str = build_output_path_with_step(pszPlPath, "販管費配賦_step0006_")

        if not os.path.exists(pszManhourPath):
            print(f"Input file not found: {pszManhourPath}")
            return 1
        if not os.path.exists(pszPlPath):
            print(f"Input file not found: {pszPlPath}")
            return 1

        objManhourMap: Dict[str, str] = load_manhour_map(pszManhourPath)
        process_pl_tsv(
            pszPlPath,
            pszOutputPath,
            pszOutputStep0001Path,
            pszOutputStep0002Path,
            pszOutputStep0003Path,
            pszOutputStep0004Path,
            pszOutputStep0005Path,
            pszOutputStep0006Path,
            pszOutputFinalPath,
            objManhourMap,
        )

        print(f"Output: {pszOutputStep0001Path}")
        print(f"Output: {pszOutputStep0002Path}")
        print(f"Output: {pszOutputStep0003Path}")
        print(f"Output: {pszOutputStep0004Path}")
        print(f"Output: {pszOutputStep0005Path}")
        print(f"Output: {pszOutputStep0006Path}")
        print(f"Output: {pszOutputFinalPath}")

    if objPairs:
        create_cumulative_reports(objPairs[0][1])
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
