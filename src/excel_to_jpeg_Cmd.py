"""ExcelシートをJPEGへ出力するためのコマンドライン入口。"""

from __future__ import annotations

import os
import sys
import time
import math

import win32com.client


def sanitize_file_component(value: str) -> str:
    """ファイル名に使えない文字を置換する。"""
    for ch in r'\/:*?"<>|':
        value = value.replace(ch, "_")
    return value


def build_used_range(objSheet):
    """最終行・最終列から範囲を組み立てる。"""
    start_row = objSheet.UsedRange.Row
    start_col = objSheet.UsedRange.Column
    last_row_cell = objSheet.Cells.Find(What="*", SearchOrder=1, SearchDirection=2)
    last_col_cell = objSheet.Cells.Find(What="*", SearchOrder=2, SearchDirection=2)
    if last_row_cell is None or last_col_cell is None:
        return objSheet.UsedRange
    return objSheet.Range(
        objSheet.Cells(start_row, start_col),
        objSheet.Cells(last_row_cell.Row, last_col_cell.Column),
    )


def build_tile_ranges(objSheet, objUsedRange, tiles: int):
    """指定したタイル数で範囲を分割する。"""
    tiles = int(tiles)
    start_row = objUsedRange.Row
    start_col = objUsedRange.Column
    end_row = objUsedRange.Row + objUsedRange.Rows.Count - 1
    end_col = objUsedRange.Column + objUsedRange.Columns.Count - 1
    total_rows = end_row - start_row + 1
    total_cols = end_col - start_col + 1
    row_steps = []
    col_steps = []
    for idx in range(tiles):
        row_start = start_row + (total_rows * idx) // tiles
        row_end = start_row + (total_rows * (idx + 1)) // tiles - 1
        row_steps.append((row_start, row_end))
        col_start = start_col + (total_cols * idx) // tiles
        col_end = start_col + (total_cols * (idx + 1)) // tiles - 1
        col_steps.append((col_start, col_end))
    ranges = []
    for row_start, row_end in row_steps:
        for col_start, col_end in col_steps:
            ranges.append(
                objSheet.Range(
                    objSheet.Cells(row_start, col_start),
                    objSheet.Cells(row_end, col_end),
                )
            )
    return ranges


def export_sheet_to_jpeg(objWorkbook, objSheet, pszOutputFilePath: str) -> None:
    """指定したシートをJPEGとして出力する。"""

    max_width = 1600
    max_height = 900
    objUsedRange = build_used_range(objSheet)
    tiles = max(
        math.ceil(objUsedRange.Width / max_width),
        math.ceil(objUsedRange.Height / max_height),
        1,
    )
    tiles = int(tiles)
    ranges = build_tile_ranges(objSheet, objUsedRange, tiles)
    for index, objRange in enumerate(ranges, start=1):
        objRange.CopyPicture(Appearance=1, Format=2)
        time.sleep(0.05)

        chart_width = max(1, min(objRange.Width, max_width))
        chart_height = max(1, min(objRange.Height, max_height))
        objChartObject = objSheet.ChartObjects().Add(
            0, 0, chart_width, chart_height


        )
        objChartObject.Activate()
        objChart = objChartObject.Chart
        objChart.Paste()
        if tiles == 1:
            output_path = pszOutputFilePath
        else:
            base, ext = os.path.splitext(pszOutputFilePath)
            output_path = f"{base}_{index}{ext}"
        objChart.Export(Filename=output_path, FilterName="JPG")
        objChartObject.Delete()


def main() -> int:
    """引数のExcelファイルから全シートをJPEGへ出力する。"""
    if len(sys.argv) != 2:
        print("使い方: python excel_to_jpeg_Cmd.py <Excelファイルパス>")
        return 1

    pszInputFilePath: str = sys.argv[1].strip().strip('"')
    pszInputFilePath = os.path.abspath(os.path.expanduser(pszInputFilePath))
    if not os.path.isfile(pszInputFilePath):
        print(f"入力ファイルが見つかりません: {pszInputFilePath}")
        return 1
    pszInputFileBaseName: str = os.path.splitext(os.path.basename(pszInputFilePath))[0]
    pszInputDirectoryPath: str = os.path.dirname(os.path.abspath(pszInputFilePath))

    objExcel = win32com.client.Dispatch("Excel.Application")
    objExcel.Visible = False
    objExcel.DisplayAlerts = False

    objWorkbook = None
    try:
        objWorkbook = objExcel.Workbooks.Open(pszInputFilePath)
        for objSheet in objWorkbook.Worksheets:
            pszSheetName: str = objSheet.Name
            pszSheetName = sanitize_file_component(pszSheetName)
            pszOutputFileName: str = f"{pszInputFileBaseName}_{pszSheetName}.jpg"
            pszOutputFilePath: str = os.path.join(
                pszInputDirectoryPath, pszOutputFileName
            )
            export_sheet_to_jpeg(objWorkbook, objSheet, pszOutputFilePath)
            export_sheet_to_jpeg(
                objWorkbook, objSheet, pszOutputFilePath, max_width, max_height
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.Close(SaveChanges=False)
        objExcel.Quit()

    return 0


if __name__ == "__main__":
    sys.exit(main())