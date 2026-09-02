# -*- coding: utf-8 -*-
"""
回家日常一鍵流程

日期語意（勿再用「昨天／今天」誤解）：
  - 賽果日／歸檔日：已打完、要補盤口觀察賽果、爬 1～3、紀錄 A→K 的那一天
  - 關注日／總表 A1：接下來要分析的日子；爬 1～8、寫入總表!A1、建頭盤、NotebookLM「今天」剪貼簿

例：日曆 8/11 晚間執行 → 賽果日=20260811、關注日=20260812
  1) 填盤口觀察（賽果日）
  2) 爬蟲賽果日步驟 1～3
  3) 總表先切回賽果日並重算 → 紀錄 A→K 貼上值，K1 強制＝賽果日
  4) 爬蟲關注日步驟 1～8
  5) 總表 A1 改為關注日並 Excel COM 重算存檔
  6) 填盤口觀察（關注日：可建頭盤；有比賽結果再填結果欄）
  7) 匯出整合 txt：{賽果日}賽後結果、{關注日}賽前分析（剪貼簿）
"""

from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta

from config import JIUJIU_XLSX_FILE, MLB_XLSX_FILE, OBSERVATION_XLSX_FILE
from excel_com_ops import archive_record_a_to_k, set_zongbiao_a1_and_recalc
from fill_observation import (
    ensure_excel_files_closed,
    parse_yyyymmdd,
    prompt_date,
)
from fill_observation import build as fill_observation_build
from integrated_export import export_integrated_for_pipeline
from pipeline_lock import acquire_pipeline_lock, release_pipeline_lock
from run_all import STEP_NAMES_ZH, _run_steps


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def _local_yyyymmdd(offset_days: int = 0) -> str:
    return (datetime.now() + timedelta(days=offset_days)).strftime("%Y%m%d")


def _prompt_results_and_focus() -> tuple[str, str]:
    """
    回傳 (賽果日, 關注日)。
    預設：賽果日＝本機今天、關注日＝明天（晚間跑完後總表 A1 指向隔天）。
    """
    default_results = _local_yyyymmdd(0)
    default_focus = _local_yyyymmdd(1)
    print("請輸入「賽果日」與「關注日」（YYYYMMDD）。")
    print("  賽果日＝已打完要歸檔／補賽果的日子（爬 1～3、紀錄 A→K）")
    print("  關注日＝接下來要分析的日子（爬 1～8、總表 A1、剪貼簿）")
    print(
        f"例：日曆今天跑完 → 賽果日={default_results}、"
        f"關注日={default_focus} → 總表 A1 會變成 {default_focus}"
    )
    print("直接按 Enter 可採用括號內預設。")
    while True:
        results_day = prompt_date(
            f"請輸入賽果日／歸檔日 (YYYYMMDD) [{default_results}]: ",
            default=default_results,
        )
        focus_day = prompt_date(
            f"請輸入關注日／總表A1 (YYYYMMDD) [{default_focus}]: ",
            default=default_focus,
        )
        r = parse_yyyymmdd(results_day)
        f = parse_yyyymmdd(focus_day)
        if f < r:
            print("錯誤：關注日不可早於賽果日，請重新輸入。")
            continue
        if (f - r).days != 1:
            print(f"警告：關注日不是賽果日的隔天（間隔 {(f - r).days} 天）。")
            try:
                ans = input("按 Enter 繼續，或輸入 Q 重輸：").strip().upper()
            except (EOFError, KeyboardInterrupt):
                raise SystemExit(1)
            if ans == "Q":
                continue
        return results_day, focus_day


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
    print("流程：補賽果日盤口觀察 → 爬賽果日1～3 → 總表切賽果日並A→K（K1＝賽果日） →")
    print("      爬關注日1～8 → 總表A1改關注日並重算 → 填關注日盤口觀察 →")
    print("      匯出 exports：賽後結果.txt ＋ 賽前分析.txt（剪貼簿＝賽前分析）")
    print()
    print("【重要】執行期間請關閉 Excel（MLB／玖九／盤口觀察）。")
    print("白天手機抓的玖九資料請已同步到本機 OneDrive。")
    print()

    results_day, focus_day = _prompt_results_and_focus()
    print()
    print(f"賽果日：{results_day}　關注日：{focus_day}")
    print(f"→ 歸檔後紀錄!K1 應為 {results_day}；步驟 5 總表!A1／紀錄!A1 應為 {focus_day}")

    acquire_pipeline_lock(base_dir, owner="回家日常")
    results_log: list[tuple[str, bool, str]] = []

    try:
        ensure_excel_files_closed(
            [MLB_XLSX_FILE, JIUJIU_XLSX_FILE, OBSERVATION_XLSX_FILE]
        )

        # 1) 賽果日盤口觀察
        _section(f"步驟 1/7：填入盤口觀察（賽果日 {results_day}）")
        try:
            _run_fill(results_day, results_day)
            results_log.append(("填盤口觀察(賽果日)", True, "OK"))
        except SystemExit as e:
            results_log.append(("填盤口觀察(賽果日)", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("填盤口觀察(賽果日)", False, str(e)))
            print(f"[失敗] {e}")

        # 2) 爬賽果日 1～3
        _section(f"步驟 2/7：爬蟲正負各隊（賽果日 {results_day}，步驟 1～3）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        step_results = _run_steps(base_dir, [1, 2, 3], results_day, results_day)
        ok = all(code == 0 for _, _, code in step_results)
        detail = ", ".join(
            f"{STEP_NAMES_ZH.get(s, s)}={'OK' if c == 0 else 'FAIL'}"
            for s, _, c in step_results
        )
        results_log.append(("爬蟲賽果日1～3", ok, detail))
        if not ok:
            print("[警告] 賽果日步驟有失敗，仍繼續後續流程。")

        # 3) 先讓「紀錄」A 區顯示賽果日，再歸檔到 K（K1＝賽果日）
        _section(
            f"步驟 3/7：總表→{results_day} 重算後，紀錄 A→K（K1={results_day}）"
        )
        ensure_excel_files_closed([MLB_XLSX_FILE])
        try:
            # 歸檔前必須讓 A 區是賽果日內容；否則若總表已是關注日，
            # 貼上值會讓 K1 也變成關注日（例如兩邊都是 20260812）。
            set_zongbiao_a1_and_recalc(MLB_XLSX_FILE, results_day)
            archive_record_a_to_k(MLB_XLSX_FILE, archive_date=results_day)
            results_log.append(("紀錄A→K", True, f"K1={results_day}"))
        except SystemExit as e:
            results_log.append(("紀錄A→K", False, str(e)))
            print(f"[失敗] {e}")
            print("後續總表重算／關注日盤口觀察可能受影響。")
        except Exception as e:
            results_log.append(("紀錄A→K", False, str(e)))
            print(f"[失敗] {e}")

        # 4) 爬關注日 1～8
        _section(f"步驟 4/7：爬蟲正負各隊（關注日 {focus_day}，步驟 1～8）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        step_results = _run_steps(
            base_dir, [1, 2, 3, 4, 5, 6, 7, 8], focus_day, focus_day
        )
        ok = all(code == 0 for _, _, code in step_results)
        detail = ", ".join(
            f"{s}={'OK' if c == 0 else 'FAIL'}" for s, _, c in step_results
        )
        results_log.append(("爬蟲關注日1～8", ok, detail))
        if not ok:
            print("[警告] 關注日步驟有失敗，仍繼續後續流程。")

        # 5) 總表 A1 = 關注日 + 重算
        _section(f"步驟 5/7：總表 A1 → {focus_day}（Excel COM 重算）")
        ensure_excel_files_closed([MLB_XLSX_FILE])
        try:
            set_zongbiao_a1_and_recalc(MLB_XLSX_FILE, focus_day)
            results_log.append(("總表A1重算", True, f"A1={focus_day}"))
        except SystemExit as e:
            results_log.append(("總表A1重算", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("總表A1重算", False, str(e)))
            print(f"[失敗] {e}")

        # 6) 關注日盤口觀察
        _section(f"步驟 6/7：填入盤口觀察（關注日 {focus_day}，可建頭盤）")
        try:
            _run_fill(focus_day, focus_day)
            results_log.append(("填盤口觀察(關注日)", True, "OK"))
        except SystemExit as e:
            results_log.append(("填盤口觀察(關注日)", False, str(e)))
            print(f"[失敗] {e}")
        except Exception as e:
            results_log.append(("填盤口觀察(關注日)", False, str(e)))
            print(f"[失敗] {e}")

        # 7) NotebookLM 整合匯出
        _section("步驟 7/7：匯出整合分析給 NotebookLM")
        ensure_excel_files_closed(
            [MLB_XLSX_FILE, OBSERVATION_XLSX_FILE]
        )
        try:
            info = export_integrated_for_pipeline(
                results_day,
                focus_day,
                mlb_xlsx=MLB_XLSX_FILE,
                observation_xlsx=OBSERVATION_XLSX_FILE,
            )
            print(f"檔案目錄：{info['export_dir']}")
            print(
                f"賽後結果：{info['results_path']}（{info['results_games']} 場）"
            )
            print(
                f"賽前分析：{info['pregame_path']}（{info['pregame_games']} 場）"
            )
            print(
                "已將「賽前分析」複製到剪貼簿 → 可直接到 NotebookLM Ctrl+V 貼上資料來源。"
            )
            print("賽後結果請在 exports\\ 手動上傳或複製。")
            results_log.append(
                (
                    "NotebookLM匯出",
                    True,
                    f"賽後{info['results_games']}場/賽前{info['pregame_games']}場",
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
        # 成功時也顯示 A1= 等簡短 detail
        extra = ""
        if detail and (
            not ok
            or detail.startswith("A1=")
            or detail.startswith("K1=")
            or "場" in detail
        ):
            extra = f"  ({detail})"
        print(f"{status}  {name}{extra}")
    print("=" * 60)
    if all_ok:
        print(
            f"全部步驟完成。紀錄 K1 應為 {results_day}；"
            f"總表／紀錄 A1 應為 {focus_day}；剪貼簿＝賽前分析。"
        )
        print(
            f"匯出：exports\\{results_day}賽後結果.txt、"
            f"exports\\{focus_day}賽前分析.txt"
        )
    else:
        print("有步驟失敗：請依上方 FAILED 項目重跑對應段落。")
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
