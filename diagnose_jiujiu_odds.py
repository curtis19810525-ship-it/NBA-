# -*- coding: utf-8 -*-
"""
診斷玖九盤口資料來源（非爬蟲）

本 repo（抓取 MLB 網路資料）的 fill_observation.py 只「讀」
  C:\\Users\\curti\\OneDrive\\MLB26\\玖九盤口變化.xlsx
的「比賽結果」「MLB快照」，再填入盤口觀察.xlsx。

真正把網頁盤口寫進「玖九盤口變化.xlsx」的程式在 sibling：
  ..\\抓取玖九盤口動態資訊
（本 GitHub repo 沒有這份程式；手機捷徑多半打 VPS，例如 108.61.247.23）

本腳本只檢查本機 Excel 裡賽果／快照是否齊全，用來對照：
  - 賽果捷徑成功、盤口捷徑失敗 → 通常「比賽結果」有資料、「MLB快照」缺或過舊
"""

from __future__ import annotations

import argparse
import os
import sys
from collections import Counter
from datetime import datetime

from openpyxl import load_workbook

from config import JIUJIU_XLSX_FILE

SIBLING_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "抓取玖九盤口動態資訊")
)


def _cell_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _parse_ts(s: str) -> datetime | None:
    s = _cell_str(s)
    if not s:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
    ):
        try:
            return datetime.strptime(s[:19], fmt)
        except ValueError:
            continue
    return None


def _kickoff_date_key(v) -> str:
    """開賽時間欄 → YYYYMMDD（能辨就辨，否則空）。"""
    if isinstance(v, datetime):
        return v.strftime("%Y%m%d")
    s = _cell_str(v)
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 8:
        return digits[:8]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return ""


def print_architecture() -> None:
    print("=" * 60)
    print("玖九盤口資料流（檢查結論用）")
    print("=" * 60)
    print("1. 盤口爬蟲（寫入 Excel）")
    print(f"   預期本機 sibling：{SIBLING_DIR}")
    print(f"   目錄存在：{'是' if os.path.isdir(SIBLING_DIR) else '否（本機／此雲端環境都沒有）'}")
    print("   手機捷徑若打 http://108.61.247.23/... → 爬蟲在 VPS，不在本 repo")
    print()
    print("2. 本 repo 消費端（只讀、不爬）")
    print("   fill_observation.py → 讀「比賽結果」建場次、讀「MLB快照」填頭尾盤口")
    print(f"   玖九檔：{JIUJIU_XLSX_FILE}")
    print("=" * 60)
    print()


def diagnose(path: str, focus_start: str | None = None, focus_end: str | None = None) -> int:
    print_architecture()

    if not os.path.exists(path):
        print(f"[失敗] 找不到玖九檔：{path}")
        print("請確認 OneDrive 已同步，或設環境變數 JIUJIU_XLSX_FILE。")
        return 2

    print(f"讀取：{path}")
    mtime = datetime.fromtimestamp(os.path.getmtime(path))
    print(f"檔案修改時間：{mtime:%Y-%m-%d %H:%M:%S}")
    print()

    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = list(wb.sheetnames)
    print(f"分頁：{', '.join(sheets)}")

    has_results = "比賽結果" in wb.sheetnames
    has_snaps = "MLB快照" in wb.sheetnames
    if not has_results:
        print("[失敗] 沒有「比賽結果」分頁（賽果捷徑寫入目標）")
    if not has_snaps:
        print("[失敗] 沒有「MLB快照」分頁（盤口捷徑寫入目標）")

    result_n = 0
    result_dates: Counter[str] = Counter()
    if has_results:
        ws = wb["比賽結果"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            result_n += 1
            key = _kickoff_date_key(row[0])
            if key:
                result_dates[key] += 1
        print(f"比賽結果：{result_n} 列")
        if result_dates:
            top = ", ".join(f"{d}×{n}" for d, n in sorted(result_dates.items())[-5:])
            print(f"  最近有資料的開賽日（最多列 5 天）：{top}")

    snap_n = 0
    latest_ts: datetime | None = None
    latest_raw = ""
    snap_dates: Counter[str] = Counter()
    if has_snaps:
        ws = wb["MLB快照"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            snap_n += 1
            ts = _parse_ts(_cell_str(row[0]))
            if ts and (latest_ts is None or ts > latest_ts):
                latest_ts = ts
                latest_raw = _cell_str(row[0])
            # 開賽時間通常在第 4 欄（0-index 3），與 SNAPSHOT_COLS 對齊
            if len(row) > 3:
                key = _kickoff_date_key(row[3])
                if key:
                    snap_dates[key] += 1
        print(f"MLB快照：{snap_n} 列")
        if latest_ts:
            print(f"  最新抓取時間：{latest_raw}")
        else:
            print("  最新抓取時間：無法解析（或沒有資料）")
        if snap_dates:
            top = ", ".join(f"{d}×{n}" for d, n in sorted(snap_dates.items())[-5:])
            print(f"  最近有快照的開賽日（最多列 5 天）：{top}")

    wb.close()
    print()

    # 區間對照（可選）
    if focus_start and focus_end:
        print(f"指定區間 {focus_start}～{focus_end}：")
        cur = datetime.strptime(focus_start, "%Y%m%d")
        end = datetime.strptime(focus_end, "%Y%m%d")
        missing_r: list[str] = []
        missing_s: list[str] = []
        while cur <= end:
            d = cur.strftime("%Y%m%d")
            if result_dates.get(d, 0) == 0:
                missing_r.append(d)
            if snap_dates.get(d, 0) == 0:
                missing_s.append(d)
            cur = cur.fromordinal(cur.toordinal() + 1)
        print(f"  缺比賽結果的日：{', '.join(missing_r) if missing_r else '無'}")
        print(f"  缺 MLB快照的日：{', '.join(missing_s) if missing_s else '無'}")
        print()

    # 判讀
    print("-" * 60)
    if result_n > 0 and snap_n == 0:
        print("判讀：賽果有、盤口快照空 → 與「手機賽果可、盤口不可」一致。")
        print("原因在盤口爬蟲／VPS（或 sibling 抓取玖九盤口動態資訊），不是 fill_observation.py。")
        print("請在 VPS 或 sibling 專案查盤口捷徑失敗 log；本 repo 無法重現 108.61.247.23。")
        code = 1
    elif result_n > 0 and snap_n > 0:
        print("判讀：賽果與快照都有資料。若盤口觀察頭尾仍空，再查配對（隊名／開賽時間）。")
        print("fill_observation.py 只負責填表；不會對網頁發盤口請求。")
        code = 0
    elif result_n == 0 and snap_n == 0:
        print("判讀：玖九檔兩分頁都空／讀不到 → 先確認存檔與 OneDrive 同步。")
        code = 2
    else:
        print("判讀：快照有、比賽結果空 → 場次來源會失敗（PR #3 起場次靠比賽結果）。")
        code = 1
    print("-" * 60)
    return code


def main() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(description="診斷玖九盤口 Excel 資料來源")
    parser.add_argument(
        "start_date",
        nargs="?",
        help="可選：開始日 YYYYMMDD，對照區間缺資料",
    )
    parser.add_argument(
        "end_date",
        nargs="?",
        help="可選：結束日 YYYYMMDD（可與開始同一天）",
    )
    parser.add_argument(
        "--file",
        default=JIUJIU_XLSX_FILE,
        help="玖九盤口變化.xlsx 路徑",
    )
    args = parser.parse_args()
    start = end = None
    if args.start_date:
        start = args.start_date.strip()
        end = (args.end_date or args.start_date).strip()
        datetime.strptime(start, "%Y%m%d")
        datetime.strptime(end, "%Y%m%d")
    raise SystemExit(diagnose(args.file, start, end))


if __name__ == "__main__":
    main()
