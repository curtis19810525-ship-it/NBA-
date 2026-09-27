# -*- coding: utf-8 -*-
"""
診斷玖九盤口／賽果資料來源（非爬蟲）

本 repo 的 fill_observation.py 只「讀」
  C:\\Users\\curti\\OneDrive\\MLB26\\玖九盤口變化.xlsx
的「比賽結果」「MLB快照」。

真正寫入該檔的程式在 sibling「抓取玖九盤口動態資訊」或手機 VPS
（例如 http://108.61.247.23/…）。捷徑畫面顯示「成功」≠ 指定日一定有新列。

常見症狀：
  - 網站 ti9999 賽果日有資料，但「比賽結果」最後一列停在前一天
  - 「執行紀錄」有跑過，但新增列數為 0 或查詢日不對
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import Counter
from datetime import datetime

from openpyxl import load_workbook

from config import JIUJIU_XLSX_FILE

SIBLING_DIR = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "抓取玖九盤口動態資訊")
)

# 玖九開賽時間常見：09-27 07:10（無年份）
_MMDD_RE = re.compile(r"(?<!\d)(\d{1,2})[-/](\d{1,2})(?!\d)")


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


def _kickoff_mmdd(v) -> str:
    """開賽時間 → MM-DD（玖九常用無年份格式）。"""
    if isinstance(v, datetime):
        return v.strftime("%m-%d")
    s = _cell_str(v)
    if not s:
        return ""
    m = _MMDD_RE.search(s)
    if m:
        return f"{int(m.group(1)):02d}-{int(m.group(2)):02d}"
    return ""


def _kickoff_date_key(v, *, default_year: int | None = None) -> str:
    """開賽時間 → YYYYMMDD；無年份時用 default_year 補上。"""
    if isinstance(v, datetime):
        return v.strftime("%Y%m%d")
    s = _cell_str(v)
    if not s:
        return ""
    # 完整日期優先
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            chunk = s[:10] if fmt != "%Y%m%d" else "".join(ch for ch in s if ch.isdigit())[:8]
            return datetime.strptime(chunk, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    mmdd = _kickoff_mmdd(s)
    if mmdd and default_year:
        return f"{default_year}{mmdd[:2]}{mmdd[3:5]}"
    return ""


def _yyyymmdd_to_mmdd(d: str) -> str:
    return f"{d[4:6]}-{d[6:8]}"


def print_architecture() -> None:
    print("=" * 60)
    print("玖九賽果／盤口資料流（檢查結論用）")
    print("=" * 60)
    print("1. 寫入端（爬蟲）")
    print(f"   預期本機 sibling：{SIBLING_DIR}")
    print(f"   目錄存在：{'是' if os.path.isdir(SIBLING_DIR) else '否（本機／此雲端環境都沒有）'}")
    print("   手機「玖九補賽果／盤口」→ 多半是 http://108.61.247.23/…（VPS）")
    print()
    print("2. 本 repo 消費端（只讀、不爬）")
    print("   fill_observation.py → 讀「比賽結果」「MLB快照」")
    print(f"   玖九檔：{JIUJIU_XLSX_FILE}")
    print("=" * 60)
    print()


def _dump_run_log(wb) -> None:
    """若有「執行紀錄」分頁，印出最後幾列（對照捷徑成功但沒寫入）。"""
    name = None
    for cand in ("執行紀錄", "執行記錄", "RunLog", "log"):
        if cand in wb.sheetnames:
            name = cand
            break
    if not name:
        print("執行紀錄：無此分頁（無法從 Excel 對照捷徑寫入筆數）")
        return

    ws = wb[name]
    rows: list[tuple] = []
    for row in ws.iter_rows(min_row=1, values_only=True):
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        rows.append(tuple(_cell_str(v) for v in row[:8]))
    print(f"執行紀錄（分頁「{name}」）最近 {min(12, len(rows))} 列：")
    if not rows:
        print("  （空白）")
        return
    for row in rows[-12:]:
        line = " | ".join(c for c in row if c)
        print(f"  {line}")
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

    # 無年份的 MM-DD 用檔案年或指定區間年推斷
    default_year = mtime.year
    if focus_start:
        default_year = int(focus_start[:4])

    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = list(wb.sheetnames)
    print(f"分頁：{', '.join(sheets)}")
    print()
    _dump_run_log(wb)

    has_results = "比賽結果" in wb.sheetnames
    has_snaps = "MLB快照" in wb.sheetnames
    if not has_results:
        print("[失敗] 沒有「比賽結果」分頁（賽果捷徑寫入目標）")
    if not has_snaps:
        print("[失敗] 沒有「MLB快照」分頁（盤口捷徑寫入目標）")

    result_n = 0
    result_dates: Counter[str] = Counter()
    result_mmdd: Counter[str] = Counter()
    last_kickoffs: list[str] = []
    if has_results:
        ws = wb["比賽結果"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            result_n += 1
            raw = row[0]
            key = _kickoff_date_key(raw, default_year=default_year)
            mmdd = _kickoff_mmdd(raw)
            if key:
                result_dates[key] += 1
            if mmdd:
                result_mmdd[mmdd] += 1
            last_kickoffs.append(_cell_str(raw))
        print(f"比賽結果：{result_n} 列")
        if result_mmdd:
            top = ", ".join(f"{d}×{n}" for d, n in sorted(result_mmdd.items())[-8:])
            print(f"  最近有資料的開賽日 MM-DD：{top}")
        if last_kickoffs:
            print("  最後 8 筆開賽時間（原始值）：")
            for s in last_kickoffs[-8:]:
                print(f"    {s}")

    snap_n = 0
    latest_ts: datetime | None = None
    latest_raw = ""
    snap_dates: Counter[str] = Counter()
    snap_mmdd: Counter[str] = Counter()
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
            if len(row) > 3:
                key = _kickoff_date_key(row[3], default_year=default_year)
                mmdd = _kickoff_mmdd(row[3])
                if key:
                    snap_dates[key] += 1
                if mmdd:
                    snap_mmdd[mmdd] += 1
        print(f"MLB快照：{snap_n} 列")
        if latest_ts:
            print(f"  最新抓取時間：{latest_raw}")
        else:
            print("  最新抓取時間：無法解析（或沒有資料）")
        if snap_mmdd:
            top = ", ".join(f"{d}×{n}" for d, n in sorted(snap_mmdd.items())[-8:])
            print(f"  最近有快照的開賽日 MM-DD：{top}")

    wb.close()
    print()

    missing_r: list[str] = []
    missing_s: list[str] = []
    if focus_start and focus_end:
        print(f"指定區間 {focus_start}～{focus_end}：")
        cur = datetime.strptime(focus_start, "%Y%m%d")
        end = datetime.strptime(focus_end, "%Y%m%d")
        while cur <= end:
            d = cur.strftime("%Y%m%d")
            mmdd = _yyyymmdd_to_mmdd(d)
            # 同時用 YYYYMMDD 與 MM-DD 對（Excel 常無年份）
            has_r = result_dates.get(d, 0) > 0 or result_mmdd.get(mmdd, 0) > 0
            has_s = snap_dates.get(d, 0) > 0 or snap_mmdd.get(mmdd, 0) > 0
            n_r = result_dates.get(d, 0) or result_mmdd.get(mmdd, 0)
            n_s = snap_dates.get(d, 0) or snap_mmdd.get(mmdd, 0)
            print(f"  {d}（{mmdd}）：比賽結果 {n_r} 列，MLB快照 {n_s} 列")
            if not has_r:
                missing_r.append(d)
            if not has_s:
                missing_s.append(d)
            cur = cur.fromordinal(cur.toordinal() + 1)
        print(f"  缺比賽結果的日：{', '.join(missing_r) if missing_r else '無'}")
        print(f"  缺 MLB快照的日：{', '.join(missing_s) if missing_s else '無'}")
        print()

    # 判讀
    print("-" * 60)
    if missing_r:
        print("判讀：捷徑／網站可能有該日賽果，但 Excel「比賽結果」沒有對應開賽日。")
        print(f"  缺日：{', '.join(missing_r)}")
        if last_kickoffs:
            print(f"  表尾停在：{last_kickoffs[-1]}")
        print("原因在 VPS「玖九補賽果」寫入端（成功畫面仍可能 0 筆新增），不是 fill_observation。")
        print("請查 Excel「執行紀錄」該次的查詢日／寫入筆數；並在 VPS 確認是否帶對賽果日。")
        print("本 GitHub repo 沒有 VPS 爬蟲原始碼，無法在此直接修寫入。")
        code = 1
    elif result_n > 0 and snap_n == 0:
        print("判讀：賽果有、盤口快照空 → 盤口捷徑／VPS 問題。")
        code = 1
    elif result_n > 0 and snap_n > 0:
        print("判讀：賽果與快照都有資料。若盤口觀察頭尾仍空，再查配對（隊名／開賽時間）。")
        print("fill_observation.py 只負責填表；不會對網頁發盤口／賽果請求。")
        code = 0
    elif result_n == 0 and snap_n == 0:
        print("判讀：玖九檔兩分頁都空／讀不到 → 先確認存檔與 OneDrive 同步。")
        code = 2
    else:
        print("判讀：快照有、比賽結果空 → 場次來源會失敗（場次靠比賽結果）。")
        code = 1
    print("-" * 60)
    return code


def main() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(description="診斷玖九賽果／盤口 Excel 資料來源")
    parser.add_argument(
        "start_date",
        nargs="?",
        help="可選：開始日 YYYYMMDD（例：20260927）",
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
