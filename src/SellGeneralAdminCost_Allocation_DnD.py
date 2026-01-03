# -*- coding: utf-8 -*-
"""
SellGeneralAdminCost_Allocation_DnD.py

ドラッグ＆ドロップで工数TSVと損益計算書TSVを受け取り、
SellGeneralAdminCost_Allocation_Cmd.py を実行するGUI。

使い方:
  ウィンドウに工数TSVと損益計算書TSVをドラッグ＆ドロップする。

仕様:
  - 入力は以下の2種類のみ:
      工数_yyyy年mm月_step10_各プロジェクトの工数.tsv
      損益計算書_yyyy年mm月_A∪B_プロジェクト名_C∪D_vertical.tsv
  - yyyy年mm月 が一致する工数/損益計算書の組み合わせのみ有効。
  - 有効な組み合わせは yyyy年mm月 の連続した範囲のみ採用する。
    (例: 2025年07月〜2025年10月 が連続していれば有効)
  - 採用された連続範囲はテキストファイルに記録する。
    (例: 採用範囲: 2025年07月〜2025年10月)
  - 有効な組み合わせのみを Cmd 版に渡して実行する。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Dict, List, Optional, Tuple

import win32api
import win32con
import win32gui


def show_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = 0
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONINFORMATION
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def show_error_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = 0
    iMessageBoxType: int = win32con.MB_OK | win32con.MB_ICONERROR
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def append_error_log(pszMessage: str) -> None:
    pszOutputPath: str = os.path.join(
        os.path.dirname(__file__),
        "SellGeneralAdminCost_Allocation_DnD_error.txt",
    )
    with open(pszOutputPath, "a", encoding="utf-8", newline="") as objFile:
        objFile.write(pszMessage + "\n")


def get_temp_output_directory() -> str:
    pszBaseDirectory: str = os.path.dirname(__file__)
    pszOutputDirectory: str = os.path.join(pszBaseDirectory, "temp")
    os.makedirs(pszOutputDirectory, exist_ok=True)
    return pszOutputDirectory


def build_unique_temp_path(pszDirectory: str, pszFileName: str) -> str:
    return os.path.join(pszDirectory, pszFileName)


def move_output_files_to_temp(pszStdOut: str) -> List[str]:
    if pszStdOut.strip() == "":
        return []
    pszTempDirectory: str = get_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    objMoved: List[str] = []
    for pszLine in pszStdOut.splitlines():
        pszLineText: str = pszLine.strip()
        if not pszLineText.startswith("Output:"):
            continue
        pszOutputPath: str = pszLineText.replace("Output:", "", 1).strip()
        if pszOutputPath == "" or not os.path.isfile(pszOutputPath):
            continue
        pszTargetPath: str = build_unique_temp_path(pszTempDirectory, os.path.basename(pszOutputPath))
        shutil.move(pszOutputPath, pszTargetPath)
        objMoved.append(pszTargetPath)
        pszBaseName: str = os.path.basename(pszTargetPath)
        if pszBaseName.startswith("累計_損益計算書_") or pszBaseName.startswith("累計_製造原価報告書_"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszBaseName)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if (
            pszBaseName.startswith("損益計算書_")
            and pszBaseName.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv")
            and "販管費配賦_step" not in pszBaseName
        ):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszBaseName)
            shutil.copy2(pszTargetPath, pszCopyPath)
    return objMoved


def parse_year_month_from_pl_csv(pszFilePath: str) -> Optional[Tuple[int, int]]:
    pszBaseName: str = os.path.basename(pszFilePath)
    objMatch = re.search(r"(\d{2})\.(\d{1,2})\.csv$", pszBaseName)
    if objMatch is None:
        return None
    iYear: int = 2000 + int(objMatch.group(1))
    iMonth: int = int(objMatch.group(2))
    if iMonth < 1 or iMonth > 12:
        return None
    return iYear, iMonth


def move_pl_outputs_to_temp(pszCsvPath: str) -> None:
    objYearMonth = parse_year_month_from_pl_csv(pszCsvPath)
    if objYearMonth is None:
        return
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    objPrefixes: List[str] = [
        f"損益計算書_{iYear}年{pszMonth}月",
        f"製造原価報告書_{iYear}年{pszMonth}月",
    ]
    pszSourceDirectory: str = os.path.dirname(pszCsvPath)
    pszTempDirectory: str = get_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    try:
        objEntries: List[str] = os.listdir(pszSourceDirectory)
    except OSError:
        return
    for pszEntry in objEntries:
        if not pszEntry.endswith(".tsv"):
            continue
        if not any(pszEntry.startswith(pszPrefix) for pszPrefix in objPrefixes):
            continue
        pszSourcePath: str = os.path.join(pszSourceDirectory, pszEntry)
        if not os.path.isfile(pszSourcePath):
            continue
        pszTargetPath: str = build_unique_temp_path(pszTempDirectory, pszEntry)
        shutil.move(pszSourcePath, pszTargetPath)
        if pszEntry.startswith("損益計算書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if pszEntry.startswith("製造原価報告書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)
        if pszEntry.startswith("製造原価報告書_") and pszEntry.endswith("_A∪B_プロジェクト名_C∪D_vertical.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)


def move_manhour_outputs_to_temp(pszCsvPath: str) -> None:
    objYearMonth = parse_year_month_from_pl_csv(pszCsvPath)
    if objYearMonth is None:
        return
    iYear, iMonth = objYearMonth
    pszMonth: str = f"{iMonth:02d}"
    pszPrefix: str = f"工数_{iYear}年{pszMonth}月"
    pszSourceDirectory: str = os.path.dirname(pszCsvPath)
    pszTempDirectory: str = get_temp_output_directory()
    pszCmdDirectory: str = os.path.dirname(__file__)
    try:
        objEntries: List[str] = os.listdir(pszSourceDirectory)
    except OSError:
        return
    for pszEntry in objEntries:
        if not pszEntry.endswith(".tsv"):
            continue
        if not pszEntry.startswith(pszPrefix):
            continue
        pszSourcePath: str = os.path.join(pszSourceDirectory, pszEntry)
        if not os.path.isfile(pszSourcePath):
            continue
        pszTargetPath: str = build_unique_temp_path(pszTempDirectory, pszEntry)
        shutil.move(pszSourcePath, pszTargetPath)
        if pszEntry.startswith("工数_") and pszEntry.endswith("_step10_各プロジェクトの工数.tsv"):
            pszCopyPath: str = os.path.join(pszCmdDirectory, pszEntry)
            shutil.copy2(pszTargetPath, pszCopyPath)


def parse_year_month_from_name(pszBaseName: str) -> Optional[str]:
    iPrefixIndex: int = pszBaseName.find("_")
    if iPrefixIndex < 0:
        return None
    iSecondIndex: int = pszBaseName.find("_", iPrefixIndex + 1)
    if iSecondIndex < 0:
        return None
    pszYearMonth: str = pszBaseName[iPrefixIndex + 1 : iSecondIndex]
    if "年" not in pszYearMonth or "月" not in pszYearMonth:
        return None
    return pszYearMonth


def parse_year_month_value(pszYearMonth: str) -> Optional[Tuple[int, int]]:
    try:
        iYearText: str = pszYearMonth.split("年", 1)[0]
        iMonthText: str = pszYearMonth.split("年", 1)[1].split("月", 1)[0]
        iYear: int = int(iYearText)
        iMonth: int = int(iMonthText)
    except (ValueError, IndexError):
        return None
    if iMonth < 1 or iMonth > 12:
        return None
    return iYear, iMonth


def is_pl_csv_file(pszBaseName: str) -> bool:
    pszNormalized: str = pszBaseName.lower()
    return re.fullmatch(r"損益計算書\d{2}\.\d{1,2}\.csv", pszNormalized) is not None


def is_manhour_csv_file(pszBaseName: str) -> bool:
    pszNormalized: str = pszBaseName.lower()
    return re.fullmatch(r"工数\d{2}\.\d{1,2}\.csv", pszNormalized) is not None


def is_consecutive_months(objYearMonths: List[Tuple[int, int]]) -> bool:
    if not objYearMonths:
        return False
    for iIndex in range(1, len(objYearMonths)):
        iPrevYear, iPrevMonth = objYearMonths[iIndex - 1]
        iNextYear, iNextMonth = objYearMonths[iIndex]
        iExpectedYear: int = iPrevYear
        iExpectedMonth: int = iPrevMonth + 1
        if iExpectedMonth == 13:
            iExpectedMonth = 1
            iExpectedYear += 1
        if iNextYear != iExpectedYear or iNextMonth != iExpectedMonth:
            return False
    return True


def collect_valid_pairs(
    objFilePaths: List[str],
) -> List[Tuple[str, str, Tuple[int, int], str]]:
    objManhourMap: Dict[str, str] = {}
    objPlMap: Dict[str, str] = {}
    for pszFilePath in objFilePaths:
        pszBaseName: str = os.path.basename(pszFilePath)
        if pszBaseName.startswith("工数_"):
            pszYearMonth: Optional[str] = parse_year_month_from_name(pszBaseName)
            if pszYearMonth is None:
                return []
            objManhourMap[pszYearMonth] = pszFilePath
            continue
        if pszBaseName.startswith("損益計算書_"):
            pszYearMonth = parse_year_month_from_name(pszBaseName)
            if pszYearMonth is None:
                return []
            objPlMap[pszYearMonth] = pszFilePath
            continue
        return []

    objPairs: List[Tuple[str, str, Tuple[int, int], str]] = []
    for pszYearMonth, pszManhourPath in objManhourMap.items():
        if pszYearMonth not in objPlMap:
            continue
        objValue: Optional[Tuple[int, int]] = parse_year_month_value(pszYearMonth)
        if objValue is None:
            continue
        objPairs.append((pszManhourPath, objPlMap[pszYearMonth], objValue, pszYearMonth))
    return objPairs


def select_consecutive_pairs(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> List[Tuple[str, str, Tuple[int, int], str]]:
    if not objPairs:
        return []
    objPairsSorted = sorted(objPairs, key=lambda objItem: objItem[2])
    objYearMonths: List[Tuple[int, int]] = [objItem[2] for objItem in objPairsSorted]
    if not is_consecutive_months(objYearMonths):
        return []
    return objPairsSorted


def build_cmd_args(objPairs: List[Tuple[str, str, Tuple[int, int], str]]) -> List[str]:
    objManhourFiles: List[str] = [objItem[0] for objItem in objPairs]
    objPlFiles: List[str] = [objItem[1] for objItem in objPairs]
    objArgs: List[str] = []
    objArgs.extend(objManhourFiles)
    objArgs.extend(objPlFiles)
    return objArgs


def write_selected_range_file(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> Optional[str]:
    if not objPairs:
        return None
    pszStart: str = objPairs[0][3]
    pszEnd: str = objPairs[-1][3]
    pszOutputDirectory: str = os.path.dirname(objPairs[0][1])
    pszOutputFileName: str = "SellGeneralAdminCost_Allocation_DnD_SelectedRange.txt"
    pszOutputPath: str = os.path.join(pszOutputDirectory, pszOutputFileName)
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        objOutputFile.write(f"採用範囲: {pszStart}〜{pszEnd}\n")
    return pszOutputPath


def run_allocation_with_pairs(
    objPairs: List[Tuple[str, str, Tuple[int, int], str]],
) -> int:
    if not objPairs:
        return 1

    pszRangePath: Optional[str] = write_selected_range_file(objPairs)
    objArgs: List[str] = build_cmd_args(objPairs)
    pszScriptPath: str = os.path.join(os.path.dirname(__file__), "SellGeneralAdminCost_Allocation_Cmd.py")
    objCommand: List[str] = [sys.executable, pszScriptPath]
    objCommand.extend(objArgs)

    try:
        objResult = subprocess.run(
            objCommand,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # noqa: BLE001
        pszErrorMessage: str = (
            "Error: unexpected exception while running SellGeneralAdminCost_Allocation_Cmd.py. Detail = "
            + str(exc)
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    if objResult.returncode != 0:
        pszStdErr: str = objResult.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        pszErrorMessage = (
            "Error: SellGeneralAdminCost_Allocation_Cmd.py exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objResult.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return objResult.returncode

    pszStdOut: str = objResult.stdout
    move_output_files_to_temp(pszStdOut)
    if pszStdOut.strip() != "":
        print(pszStdOut)
    pszStdOut = "成功しました！"
    if pszRangePath is not None:
        pszStdOut += "\n\n採用範囲を記録しました: " + pszRangePath
    show_message_box(pszStdOut, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def run_pl_csv_to_tsv(
    objCsvFiles: List[str],
) -> int:
    pszScriptPath: str = os.path.join(os.path.dirname(__file__), "PL_CsvToTsv_Cmd.py")
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: PL_CsvToTsv_Cmd.py not found. Path = " + pszScriptPath
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    objCommand: List[str] = [sys.executable, pszScriptPath] + objCsvFiles
    append_error_log("Running: " + " ".join(objCommand))
    try:
        objResult = subprocess.run(
            objCommand,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # noqa: BLE001
        pszErrorMessage: str = (
            "Error: unexpected exception while running PL_CsvToTsv_Cmd.py. Detail = "
            + str(exc)
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    if objResult.returncode != 0:
        pszStdErr: str = objResult.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        pszErrorMessage = (
            "Error: PL_CsvToTsv_Cmd.py exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objResult.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return objResult.returncode

    pszStdOut: str = objResult.stdout.strip()
    if pszStdOut != "":
        print(pszStdOut)
        move_output_files_to_temp(pszStdOut)

    for pszCsvPath in objCsvFiles:
        move_pl_outputs_to_temp(pszCsvPath)

    pszMessage: str = "PL_CsvToTsv_Cmd.py finished successfully."
    if pszStdOut != "":
        pszMessage = pszStdOut
    show_message_box(pszMessage, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def run_manhour_csv_to_sheet(
    objCsvFiles: List[str],
) -> int:
    pszScriptPath: str = os.path.join(
        os.path.dirname(__file__),
        "make_manhour_to_sheet8_01_0001.py",
    )
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: make_manhour_to_sheet8_01_0001.py not found. Path = "
            + pszScriptPath
        )
        append_error_log(pszErrorMessage)
        show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
        return 1

    objMessages: List[str] = []
    for pszCsvPath in objCsvFiles:
        objCommand: List[str] = [sys.executable, pszScriptPath, pszCsvPath]
        try:
            objResult = subprocess.run(
                objCommand,
                check=False,
                capture_output=True,
                text=True,
            )
        except Exception as exc:  # noqa: BLE001
            pszErrorMessage: str = (
                "Error: unexpected exception while running make_manhour_to_sheet8_01_0001.py. Detail = "
                + str(exc)
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            return 1

        if objResult.returncode != 0:
            pszStdErr: str = objResult.stderr
            if pszStdErr.strip() == "":
                pszStdErr = "Process exited with non-zero return code and no stderr output."
            pszErrorMessage = (
                "Error: make_manhour_to_sheet8_01_0001.py exited with non-zero return code.\n\n"
                + "Return code = "
                + str(objResult.returncode)
                + "\n\n"
                + "stderr:\n"
                + pszStdErr
            )
            append_error_log(pszErrorMessage)
            show_error_message_box(pszErrorMessage, "SellGeneralAdminCost_Allocation_DnD")
            return objResult.returncode

        pszStdOut: str = objResult.stdout.strip()
        if pszStdOut != "":
            print(pszStdOut)
            move_output_files_to_temp(pszStdOut)
        move_manhour_outputs_to_temp(pszCsvPath)

    pszMessage: str = "make_manhour_to_sheet8_01_0001.py finished successfully."
    show_message_box(pszMessage, "SellGeneralAdminCost_Allocation_DnD")
    return 0


def draw_instruction_text(
    iWindowHandle: int,
) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(
        iWindowHandle,
    )
    objClientRect = win32gui.GetClientRect(
        iWindowHandle,
    )

    iMargin: int = 5
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )

    pszInstructionText: str = (
        "工数TSVと損益計算書TSVを、このウィンドウにドラッグ＆ドロップしてください。\n"
        "有効な年月の連続範囲のみ処理されます。\n"
        "採用された年月範囲はテキストファイルに記録します。"
    )

    iDrawTextFormat: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        iDrawTextFormat,
    )
    win32gui.EndPaint(
        iWindowHandle,
        objPaintStruct,
    )


def window_proc(
    iWindowHandle: int,
    iMessage: int,
    iWparam: int,
    iLparam: int,
) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(
            iWindowHandle,
            True,
        )
        return 0

    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        iFileCount: int = win32api.DragQueryFile(
            iDropHandle,
            -1,
        )

        objFiles: List[str] = []
        for iIndex in range(iFileCount):
            pszFilePath: str = win32api.DragQueryFile(
                iDropHandle,
                iIndex,
            )
            objFiles.append(pszFilePath)

        win32api.DragFinish(iDropHandle)

        objCsvFiles: List[str] = []
        objManhourCsvFiles: List[str] = []
        bAllCsv: bool = True
        bAllManhourCsv: bool = True
        for pszFilePath in objFiles:
            pszBaseName: str = os.path.basename(pszFilePath)
            if is_pl_csv_file(pszBaseName):
                objCsvFiles.append(pszFilePath)
            else:
                bAllCsv = False
            if is_manhour_csv_file(pszBaseName):
                objManhourCsvFiles.append(pszFilePath)
            else:
                bAllManhourCsv = False

        if bAllCsv and objCsvFiles:
            run_pl_csv_to_tsv(objCsvFiles)
            return 0
        if bAllManhourCsv and objManhourCsvFiles:
            run_manhour_csv_to_sheet(objManhourCsvFiles)
            return 0

        objPairs = collect_valid_pairs(objFiles)
        objPairs = select_consecutive_pairs(objPairs)
        if not objPairs:
            pszErrorMessage: str = (
                "Error: dropped files are invalid or not consecutive by year/month."
            )
            show_error_message_box(
                pszErrorMessage,
                "SellGeneralAdminCost_Allocation_DnD",
            )
            return 0

        run_allocation_with_pairs(objPairs)
        return 0

    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(
            iWindowHandle,
        )
        return 0

    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    return win32gui.DefWindowProc(
        iWindowHandle,
        iMessage,
        iWparam,
        iLparam,
    )


def register_window_class(
    pszWindowClassName: str,
) -> int:
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1

    iClassAtom: int = win32gui.RegisterClass(objWndClass)
    return iClassAtom


def create_main_window(
    pszWindowClassName: str,
    pszWindowTitle: str,
) -> int:
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowExStyle: int = win32con.WS_EX_ACCEPTFILES

    iWindowPosX: int = win32con.CW_USEDEFAULT
    iWindowPosY: int = win32con.CW_USEDEFAULT
    iWindowHeight: int = 320
    iWindowWidth: int = int(iWindowHeight * 1.618)
    iParentWindowHandle: int = 0
    iMenuHandle: int = 0

    iWindowHandle: int = win32gui.CreateWindowEx(
        iWindowExStyle,
        pszWindowClassName,
        pszWindowTitle,
        iWindowStyle,
        iWindowPosX,
        iWindowPosY,
        iWindowWidth,
        iWindowHeight,
        iParentWindowHandle,
        iMenuHandle,
        iInstanceHandle,
        None,
    )

    win32gui.ShowWindow(
        iWindowHandle,
        win32con.SW_SHOWNORMAL,
    )
    win32gui.UpdateWindow(iWindowHandle)

    iHwndInsertAfter: int = win32con.HWND_TOPMOST
    iFlags: int = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    win32gui.SetWindowPos(
        iWindowHandle,
        iHwndInsertAfter,
        0,
        0,
        0,
        0,
        iFlags,
    )

    win32gui.DragAcceptFiles(
        iWindowHandle,
        True,
    )

    return iWindowHandle


def main() -> None:
    pszWindowClassName: str = "SellGeneralAdminCostAllocationDndWindowClass"
    pszWindowTitle: str = "SellGeneralAdminCost Allocation (Drag & Drop)"

    try:
        register_window_class(pszWindowClassName)
    except Exception as exc:
        pszErrorMessage: str = (
            "Error: failed to register window class. Detail = " + str(exc)
        )
        show_error_message_box(
            pszErrorMessage,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    try:
        create_main_window(
            pszWindowClassName,
            pszWindowTitle,
        )
    except Exception as exc:
        pszErrorMessage: str = (
            "Error: failed to create main window. Detail = " + str(exc)
        )
        show_error_message_box(
            pszErrorMessage,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return

    try:
        win32gui.PumpMessages()
    except Exception as exc:
        pszErrorMessage: str = (
            "Error: unexpected exception in message loop. Detail = " + str(exc)
        )
        show_error_message_box(
            pszErrorMessage,
            "SellGeneralAdminCost_Allocation_DnD",
        )
        return
    return


if __name__ == "__main__":
    main()
