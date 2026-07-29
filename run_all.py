# -*- coding: utf-8 -*-
"""
Crawler + OX + Team Stats (All-in-One)

模式：
  - 無參數：進階步驟選單（爬蟲正負各隊.bat）
  - --preset today_results：抓取今日賽果（只跑步驟 1）
  - --preset full_range：完整抓取今日及明日（步驟 1～8 + 盤口觀察，起迄日由使用者輸入）
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime

from pipeline_lock import acquire_pipeline_lock, release_pipeline_lock

STEP_NAMES = {
    1: "Crawler (執行範圍爬蟲)",
    2: "Batch Convert OX (批次轉換正負盤)",
    3: "Batch Convert Team Stats (批次轉換各隊正負)",
    4: "ESPN MLB Batting + Team Pitchers",
    5: "ESPN MLB Team Total",
    6: "ESPN MLB Player Injuries",
    7: "Rotowire MLB Starting Lineups",
    8: "Incremental Name Mapping Update",
}

STEP_NAMES_ZH = {
    1: "爬蟲（執行範圍爬蟲）",
    2: "批次轉換正負盤",
    3: "批次轉換各隊正負",
    4: "ESPN 球員打擊＋各隊投手",
    5: "ESPN 各隊 Total",
    6: "ESPN 球員傷兵",
    7: "Rotowire 先發名單",
    8: "更新姓名對照表",
}

PRESET_TODAY = "today_results"
PRESET_FULL = "full_range"
OBS_STEP_ID = 90
OBS_STEP_NAME = "填入盤口觀察"


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr, sys.stdin):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def _render_menu(selected):
    marker = lambda n: "x" if n in selected else " "
    print("\n" + "=" * 60)
    print("MLB Crawler Step Menu（進階）")
    print("=" * 60)
    print("Toggle by typing step number (e.g. 1 or 1,3,5)")
    print("A=Select all, N=Select none, 0=Start, Q=Quit")
    print()
    print(f"[{marker(1)}] 1. 爬蟲（執行範圍爬蟲.py）")
    print(f"[{marker(2)}] 2. 批次轉換正負盤（批次轉換正負盤.py）")
    print(f"[{marker(3)}] 3. 批次轉換各隊正負（批次轉換各隊正負.py）")
    print(f"[{marker(4)}] 4. 爬取 ESPN MLB 球員打擊＋各隊投手（球員狀態、各隊投手）")
    print(f"[{marker(5)}] 5. 爬取 ESPN MLB 各隊 Total（各隊total）")
    print(f"[{marker(6)}] 6. 爬取 ESPN MLB 球員傷兵（球員傷兵）")
    print(f"[{marker(7)}] 7. 爬取 Rotowire MLB 先發名單（先發名單）")
    print(f"[{marker(8)}] 8. 更新姓名對照表（增量，只補 C/D）")
    print("=" * 60)


def _prompt_steps():
    selected = set()
    while True:
        _render_menu(selected)
        try:
            cmd = input("請輸入指令：").strip().upper()
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            return None

        if cmd == "Q":
            print("Cancelled.")
            return None
        if cmd == "A":
            selected = {1, 2, 3, 4, 5, 6, 7, 8}
            continue
        if cmd == "N":
            selected = set()
            continue
        if cmd == "0":
            if not selected:
                print("錯誤：尚未勾選任何步驟，請至少選 1 個步驟。")
                continue
            return sorted(selected)

        tokens = [t.strip() for t in cmd.split(",") if t.strip()]
        invalid = []
        for token in tokens:
            if token.isdigit():
                step = int(token)
                if step in STEP_NAMES:
                    if step in selected:
                        selected.remove(step)
                    else:
                        selected.add(step)
                else:
                    invalid.append(token)
            else:
                invalid.append(token)
        if invalid:
            print(f"無效輸入：{', '.join(invalid)}")


def _prompt_date(prompt_text: str) -> str:
    while True:
        try:
            value = input(prompt_text).strip()
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            raise SystemExit(1)
        try:
            datetime.strptime(value, "%Y%m%d")
            return value
        except ValueError:
            print("日期格式錯誤，請使用 YYYYMMDD（例如 20251022）。")


def _prompt_date_range(
    *,
    warn_if_span: bool = False,
    hint: str = "",
) -> tuple[str, str]:
    """兩行輸入起迄日；結束早於開始則整段重輸（不對調）。"""
    while True:
        start = _prompt_date(f"請輸入開始日期 (YYYYMMDD){hint}: ")
        end = _prompt_date(f"請輸入結束日期 (YYYYMMDD){hint}: ")
        start_dt = datetime.strptime(start, "%Y%m%d")
        end_dt = datetime.strptime(end, "%Y%m%d")
        if end_dt < start_dt:
            print("錯誤：結束日期早於開始日期，請重新輸入起迄兩行。")
            continue
        if warn_if_span and start != end:
            print()
            print("警告：此快捷建議開始日與結束日為同一天（通常都填今天）。")
            print(f"目前輸入：{start} ～ {end}")
            try:
                ans = input("按 Enter 繼續，或輸入 Q 重新輸入日期：").strip().upper()
            except (EOFError, KeyboardInterrupt):
                print("\nCancelled.")
                raise SystemExit(1)
            if ans == "Q":
                continue
        return start, end


def _prompt_date_range_if_needed(steps):
    if not any(step in {1, 2, 3, 7} for step in steps):
        return None, None
    hint = ""
    if 7 in steps:
        hint = "（台灣日；步驟7 Rotowire 會換算美東日曆再抓取）"
    return _prompt_date_range(warn_if_span=False, hint=hint)


def _build_command(step, start, end):
    if step == 1:
        return [sys.executable, "執行範圍爬蟲.py", start, end]
    if step == 2:
        return [sys.executable, "批次轉換正負盤.py", f"{start}~{end}"]
    if step == 3:
        return [sys.executable, "批次轉換各隊正負.py", f"{start}~{end}"]
    if step == 4:
        return [sys.executable, "espn_player_stats.py"]
    if step == 5:
        return [sys.executable, "espn_team_total.py"]
    if step == 6:
        return [sys.executable, "espn_player_stats.py", "--injuries-only"]
    if step == 7:
        return [sys.executable, "mlb_lineups.py", start, end]
    if step == 8:
        return [sys.executable, "更新姓名對照表.py"]
    raise ValueError(f"Unsupported step: {step}")


def _run_steps(base_dir: str, steps: list[int], start: str | None, end: str | None) -> list[tuple]:
    """依序執行；失敗不中斷（乙）。回傳 (step_id, name, returncode) 列表。"""
    results = []
    for idx, step in enumerate(steps, start=1):
        name = STEP_NAMES_ZH.get(step, STEP_NAMES[step])
        print()
        print("=" * 60)
        print(f"Step {idx}/{len(steps)}: {name}")
        print("=" * 60)
        cmd = _build_command(step, start, end)
        proc = subprocess.run(cmd, cwd=base_dir)
        status = "OK" if proc.returncode == 0 else "FAILED"
        print(f"Result: {status} (exit code={proc.returncode})")
        results.append((step, name, proc.returncode))
    return results


def _run_fill_observation(base_dir: str, start: str, end: str) -> tuple:
    """盤口觀察：使用完整起迄（定案乙）。子行程帶 --no-lock（父行程已持鎖）。"""
    print()
    print("=" * 60)
    print(f"最後步驟: {OBS_STEP_NAME}（{start} ～ {end}）")
    print("=" * 60)
    cmd = [
        sys.executable,
        "fill_observation.py",
        start,
        end,
        "--no-lock",
    ]
    proc = subprocess.run(cmd, cwd=base_dir)
    status = "OK" if proc.returncode == 0 else "FAILED"
    print(f"Result: {status} (exit code={proc.returncode})")
    return (OBS_STEP_ID, OBS_STEP_NAME, proc.returncode)


def _print_summary(results: list[tuple]) -> bool:
    print()
    print("=" * 60)
    print("Summary（成功／失敗一覽）")
    print("=" * 60)
    all_ok = True
    failed = []
    for step, name, code in results:
        status = "OK" if code == 0 else "FAILED"
        label = f"Step {step}" if step != OBS_STEP_ID else "盤口觀察"
        print(f"{label} - {name}: {status}")
        if code != 0:
            all_ok = False
            failed.append(step if step != OBS_STEP_ID else "盤口觀察")
    print("=" * 60)
    if failed:
        print("建議：失敗項目可用「爬蟲正負各隊.bat」進階選單勾選重跑；")
        print("　　　若僅盤口觀察失敗，可執行「填入盤口觀察.bat」。")
        print(f"失敗項目：{', '.join(str(x) for x in failed)}")
    else:
        print("全部步驟成功。")
    print("=" * 60)
    return all_ok


def _run_preset_today_results(base_dir: str) -> int:
    print("=" * 60)
    print("抓取今日賽果")
    print("=" * 60)
    print("固定只跑步驟 1（爬蟲）。請先關閉 Excel。")
    print()
    start, end = _prompt_date_range(warn_if_span=True)
    print()
    print(f"將執行：步驟 1，日期 {start} ～ {end}")
    results = _run_steps(base_dir, [1], start, end)
    _print_summary(results)
    print()
    print("下一步：")
    print("  1. 開啟 MLB26-27「紀錄」，複製今天有比賽區（約 A1:I33）")
    print("  2. 貼到 NotebookLM「資料來源」並改檔名")
    print("  3. 關閉 Excel 後，執行「完整抓取今日及明日」")
    return 0 if all(c == 0 for _, _, c in results) else 1


def _run_preset_full_range(base_dir: str) -> int:
    print("=" * 60)
    print("完整抓取今日及明日")
    print("=" * 60)
    print("將依序執行步驟 1～8，最後填入盤口觀察（完整起迄日；場次依玖九比賽結果）。")
    print("某步失敗不中斷，全部跑完再彙報。")
    print("請先存檔並關閉 Excel（MLB／玖九／盤口觀察）。")
    print()
    start, end = _prompt_date_range(
        warn_if_span=False,
        hint="（建議：開始=今天、結束=明天；休賽可自行調整）",
    )
    steps = [1, 2, 3, 4, 5, 6, 7, 8]
    print()
    print(f"將執行：步驟 1～8 + 盤口觀察，日期 {start} ～ {end}")
    results = _run_steps(base_dir, steps, start, end)
    # 無門閘：即使前面失敗，仍跑盤口觀察
    results.append(_run_fill_observation(base_dir, start, end))
    ok = _print_summary(results)
    print()
    print("下一步：")
    print("  1. 開啟 MLB26-27「總表」，A1 填明天日期並檢查")
    print("  2. 「紀錄」複製明天相關區塊 → NotebookLM 對話框要推薦")
    return 0 if ok else 1


def _run_menu(base_dir: str) -> int:
    print("=" * 60)
    print("MLB Crawler + OX + Team Stats（進階選單）")
    print("=" * 60)
    print()
    print("Please close Excel before running.")
    print()

    steps = _prompt_steps()
    if not steps:
        return 1

    start, end = _prompt_date_range_if_needed(steps)
    print("\n" + "=" * 60)
    print("Selected steps: " + ", ".join(str(s) for s in steps))
    if start and end:
        print(f"Date range: {start}～{end}")
    print("=" * 60)

    results = _run_steps(base_dir, steps, start, end)
    return 0 if _print_summary(results) else 1


def main() -> int:
    _configure_stdio()
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    parser = argparse.ArgumentParser(description="MLB 管線：快捷預設或進階選單")
    parser.add_argument(
        "--preset",
        choices=(PRESET_TODAY, PRESET_FULL),
        help="today_results=抓取今日賽果；full_range=完整抓取今日及明日",
    )
    args = parser.parse_args()

    owner = {
        PRESET_TODAY: "抓取今日賽果",
        PRESET_FULL: "完整抓取今日及明日",
    }.get(args.preset, "爬蟲正負各隊（進階選單）")

    acquire_pipeline_lock(base_dir, owner=owner)
    try:
        if args.preset == PRESET_TODAY:
            return _run_preset_today_results(base_dir)
        if args.preset == PRESET_FULL:
            return _run_preset_full_range(base_dir)
        return _run_menu(base_dir)
    finally:
        release_pipeline_lock()


if __name__ == "__main__":
    sys.exit(main())
