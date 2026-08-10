# -*- coding: utf-8 -*-
"""
回家日常一鍵流程

例：昨天=20260731、今天=20260801
  1) 填盤口觀察（昨天賽果）
  2) 爬蟲昨天步驟 1～3
  3) 紀錄 A1:I33 → 清空 K1:S33 後貼上值（Excel COM）
  4) 爬蟲今天步驟 1～8
  5) 總表 A1 改為今天並 Excel COM 重算存檔
  6) 填盤口觀察（今天：可建頭盤；有比賽結果再填結果欄）
  7) 匯出昨天／今天「紀錄」給 NotebookLM（剪貼簿預設今天）
"""

from __future__ import annotations

import os
import sys

from config import JIUJIU_XLSX_FILE, MLB_XLSX_FILE, OBSERVATION_XLSX_FILE
from excel_com_ops import archive_record_a_to_k, set_zongbiao_a1_and_recalc
from fill_observation import (
    ensure_excel_files_closed,
    parse_yyyymmdd,
    prompt_date,
)
from fill_observation import build as fill_observation_build
from notebooklm_export import export_yesterday_and_today
from pipeline_lock import acquire_pipeline_lock, release_pipeline_lock
from run_all import STEP_NAMES_ZH, _run_steps


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def _prompt_yesterday_today() -> tuple[str, str]:
    print("請輸入「昨天」與「今天」（YYYYMMDD）。")
    while True:
        yesterday = prompt_date("請輸入昨天日期 (YYYYMMDD): ")
        today = prompt_date("請輸入今天日期 (YYYYMMDD): ")
        y = parse_yyyymmdd(yesterday)
        t = parse_yyyymmdd(today)
        if t < y:
            print("錯誤：今天不可早於昨天，請重新輸入。")
            continue
        if (t - y).days != 1:
            print(f"警告：今天不是昨天的隔天（間隔 {(t - y).days} 天）。")
            try:
                ans = input("按 Enter 繼續，或輸入 Q 重輸：").strip().upper()
            except (EOFError, KeyboardInterrupt):
                raise SystemExit(1)
            if ans == "Q":
                continue
        return yesterday, today


def _section(title: str) -> None:
    print()
    print("=" * 60)
    print(title)
    print("=" * 60)


def _run_fill(start: str, end: str) -> None:
    ensure_excel_files_closed(
        [MLB_XLSX_FILE, JIUJIU_XLSX_FILE, OBSERVATION_XLSX_FILE]
    )
    fill_observation_build(parse_yyyymmdd(start), parse_yyyymmdd(end))


def main() -> int:
    _configure_stdio()
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    print("=" * 60)
    print("回家日常（一鍵）")
    print("=" * 60)
    print("流程：補昨天盤口觀察 → 爬昨天1～3 → 紀錄A→K →")
    print("      爬今天1～8 → 總表A1改今天並重算 → 填今天盤口觀察 →")
    print("      匯出紀錄給 NotebookLM（昨天＋今天；剪貼簿＝今天）")
    print()
    print("【重要】執行期間請關閉 Excel（MLB／玖九／盤口觀察）。")
    print("白天手機抓的玖九資料請已同步到本機 OneDrive。")
    print()

    yesterday, today = _prompt_yesterday_today()
    print()
    print(f"昨天：{yesterday}　今天：{today}")

    acquire_pipeline_lock(base_dir, owner="回家日常")
    results_log: list[tuple[str, bool, str]] = []

    try:
        ensure_excel_files_closed(
            [MLB_XLSX_FILE, JIUJIU_XLSX_FILE, OBSERVATION_XLSX_FILE]
        )

        # 1) 昨天盤口觀察（賽果）
        _section(f"步驟 1/7：填入盤口觀察（昨天 {yesterday}）")
        try:
            _run_fill(yesterday, yesterday)
            results_log.append(("填盤口觀察(昨天)", True, "OK"))
        except SystemExit as e:
            results_log.append(("填盤口觀察(昨天)", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("填盤口觀察(昨天)", False, str(e)))
            print(f"[失敗] {e}")

        # 2) 爬昨天 1～3
        _section(f"步驟 2/7：爬蟲正負各隊（昨天 {yesterday}，步驟 1～3）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        step_results = _run_steps(base_dir, [1, 2, 3], yesterday, yesterday)
        ok = all(code == 0 for _, _, code in step_results)
        detail = ", ".join(
            f"{STEP_NAMES_ZH.get(s, s)}={'OK' if c == 0 else 'FAIL'}"
            for s, _, c in step_results
        )
        results_log.append(("爬蟲昨天1～3", ok, detail))
        if not ok:
            print("[警告] 昨天步驟有失敗，仍繼續後續流程。")

        # 3) 紀錄 A→K
        _section("步驟 3/7：紀錄備份 A1:I33 → 清空 K1:S33 後貼上值")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        try:
            archive_record_a_to_k(MLB_XLSX_FILE)
            results_log.append(("紀錄A→K", True, "OK"))
        except SystemExit as e:
            results_log.append(("紀錄A→K", False, str(e)))
            print(f"[失敗] {e}")
            print("後續總表重算／今天盤口觀察可能受影響。")
        except Exception as e:
            results_log.append(("紀錄A→K", False, str(e)))
            print(f"[失敗] {e}")

        # 4) 爬今天 1～8
        _section(f"步驟 4/7：爬蟲正負各隊（今天 {today}，步驟 1～8）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        step_results = _run_steps(
            base_dir, [1, 2, 3, 4, 5, 6, 7, 8], today, today
        )
        ok = all(code == 0 for _, _, code in step_results)
        detail = ", ".join(
            f"{s}={'OK' if c == 0 else 'FAIL'}" for s, _, c in step_results
        )
        results_log.append(("爬蟲今天1～8", ok, detail))
        if not ok:
            print("[警告] 今天步驟有失敗，仍繼續後續流程。")

        # 5) 總表 A1 + 重算
        _section(f"步驟 5/7：總表 A1 → {today}（Excel COM 重算）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        try:
            set_zongbiao_a1_and_recalc(MLB_XLSX_FILE, today)
            results_log.append(("總表A1重算", True, "OK"))
        except SystemExit as e:
            results_log.append(("總表A1重算", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("總表A1重算", False, str(e)))
            print(f"[失敗] {e}")

        # 6) 今天盤口觀察
        _section(f"步驟 6/7：填入盤口觀察（今天 {today}，可建頭盤）")
        try:
            _run_fill(today, today)
            results_log.append(("填盤口觀察(今天)", True, "OK"))
        except SystemExit as e:
            results_log.append(("填盤口觀察(今天)", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("填盤口觀察(今天)", False, str(e)))
            print(f"[失敗] {e}")

        # 7) NotebookLM 匯出
        _section("步驟 7/7：匯出紀錄給 NotebookLM（昨天＋今天）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        try:
            info = export_yesterday_and_today(yesterday, today, mlb_xlsx=MLB_XLSX_FILE)
            print(f"昨天：{info['yesterday_path']}（{info['yesterday_games']} 場）")
            print(f"今天：{info['today_path']}（{info['today_games']} 場）")
            print("已將「今天」內容複製到剪貼簿 → 可直接到 NotebookLM Ctrl+V 貼上資料來源。")
            print("昨天請開啟對應 txt 手動複製。")
            results_log.append(
                (
                    "NotebookLM匯出",
                    True,
                    f"昨{info['yesterday_games']}場/今{info['today_games']}場",
                )
            )
        except SystemExit as e:
            results_log.append(("NotebookLM匯出", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("NotebookLM匯出", False, str(e)))
            print(f"[失敗] {e}")

    finally:
        release_pipeline_lock()

    print()
    print("=" * 60)
    print("回家日常 Summary")
    print("=" * 60)
    all_ok = True
    for name, ok, detail in results_log:
        status = "OK" if ok else "FAILED"
        if not ok:
            all_ok = False
        print(f"{status}  {name}" + (f"  ({detail})" if detail and not ok else ""))
    print("=" * 60)
    if all_ok:
        print("全部步驟完成。可將剪貼簿內容貼到 NotebookLM 資料來源。")
    else:
        print("有步驟失敗：請依上方 FAILED 項目重跑對應段落。")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
