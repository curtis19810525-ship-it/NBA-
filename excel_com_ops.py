# -*- coding: utf-8 -*-
"""透過本機 Excel COM 重算與複製「紀錄」區塊（僅 Windows + 桌面版 Excel）。"""

from __future__ import annotations

import os
import sys
import time


SHEET_ZONG = "總表"
SHEET_RECORD = "紀錄"


def _require_windows_excel():
    if sys.platform != "win32":
        raise SystemExit("Excel COM 僅支援 Windows 本機（目前環境無法執行）。")
    try:
        import win32com.client  # noqa: F401
    except ImportError as e:
        raise SystemExit(
            "缺少 pywin32。請執行：python -m pip install pywin32\n"
            f"原始錯誤：{e}"
        ) from e


def _abs(path: str) -> str:
    return os.path.abspath(path)


def _open_excel_workbook(path: str):
    import win32com.client

    path = _abs(path)
    if not os.path.exists(path):
        raise SystemExit(f"找不到檔案：{path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)
    except Exception:
        try:
            excel.Quit()
        except Exception:
            pass
        raise
    return excel, wb


def _close_excel(excel, wb=None, *, save: bool = False) -> None:
    try:
        if wb is not None:
            if save:
                wb.Save()
            wb.Close(SaveChanges=1 if save else 0)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    # 給 OneDrive / Excel 一點釋放時間
    time.sleep(0.8)


def archive_record_a_to_k(
    mlb_xlsx: str,
    *,
    src_range: str = "A1:I33",
    dst_range: str = "K1:S33",
) -> None:
    """
    先清空 K1:S33，再把 A1:I33「貼上值」到 K1:S33。
    使用 Excel COM，確保貼的是計算後的值而非公式。
    """
    _require_windows_excel()
    excel, wb = _open_excel_workbook(mlb_xlsx)
    try:
        if SHEET_RECORD not in [sh.Name for sh in wb.Worksheets]:
            raise SystemExit(f"找不到分頁「{SHEET_RECORD}」")
        ws = wb.Worksheets(SHEET_RECORD)
        # 先清空目標區
        ws.Range(dst_range).ClearContents()
        # 複製來源 → 目標貼上值
        ws.Range(src_range).Copy()
        ws.Range(dst_range).PasteSpecial(Paste=-4163)  # xlPasteValues
        try:
            excel.CutCopyMode = False
        except Exception:
            pass
        wb.Save()
        print(f"[Excel] 紀錄已備份：清空 {dst_range} 後，將 {src_range} 貼上值到 {dst_range}")
    finally:
        _close_excel(excel, wb, save=False)


def set_zongbiao_a1_and_recalc(mlb_xlsx: str, yyyymmdd: str) -> None:
    """將「總表」A1 設為 YYYYMMDD，全本重算後存檔。"""
    _require_windows_excel()
    yyyymmdd = str(yyyymmdd).strip()
    if len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        raise SystemExit(f"總表 A1 日期格式錯誤：{yyyymmdd}")

    excel, wb = _open_excel_workbook(mlb_xlsx)
    try:
        names = [sh.Name for sh in wb.Worksheets]
        if SHEET_ZONG not in names:
            raise SystemExit(f"找不到分頁「{SHEET_ZONG}」（現有：{', '.join(names)}）")
        ws = wb.Worksheets(SHEET_ZONG)
        # 寫入文字／數字皆可；以字串寫入避免被当成序號
        ws.Range("A1").NumberFormat = "@"
        ws.Range("A1").Value = yyyymmdd

        # 強制重算
        try:
            excel.CalculateFullRebuild()
        except Exception:
            try:
                excel.CalculateFull()
            except Exception:
                excel.Calculate()

        wb.Save()
        print(f"[Excel] 總表!A1 已設為 {yyyymmdd}，並完成重算存檔")
    finally:
        _close_excel(excel, wb, save=False)
