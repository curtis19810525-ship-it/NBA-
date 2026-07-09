"""
選項 A：每日數據轉正負盤
從「每日數據」分頁讀取資料，計算 O/X，寫入「正負盤」分頁
支援單日轉換或指定日期範圍
"""

import os
import sys
import re
import shutil
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# 引入設定檔
try:
    from config import NBA_XLSX_FILE
except ImportError:
    NBA_XLSX_FILE = r"C:\Users\curti\OneDrive\MLB26\MLB26-27數據.xlsx"

# 備份目錄
BACKUP_DIR = os.path.join(os.path.dirname(__file__), "回朔點備分")


def create_backup():
    """建立還原點備份"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)
    
    files_to_backup = [
        "每日數據轉正負盤.py",
        "批次轉換正負盤.py"
    ]
    
    backed_up = []
    for filename in files_to_backup:
        source_path = os.path.join(os.path.dirname(__file__), filename)
        if os.path.exists(source_path):
            backup_filename = f"{filename}_備份_{timestamp}"
            backup_path = os.path.join(BACKUP_DIR, backup_filename)
            shutil.copy2(source_path, backup_path)
            backed_up.append(backup_filename)
    
    if backed_up:
        print(f"✓ 已建立還原點：{timestamp}")
        print(f"  備份檔案：{', '.join(backed_up)}")
    return timestamp


def calculate_ox(away_handicap, home_handicap, away_score, home_score, away_ratio, home_ratio):
    """
    計算 O/X 結果
    
    規則：O = 高比例方過盤，X = 高比例方沒過盤。
    雙方皆 50% 時，以讓分方過盤與否為準。
    過盤條件：實際分差 >= X+1（即大於 X 分）。
    
    Args:
        away_handicap: 客隊讓分（如「6分輸」或空）
        home_handicap: 主隊讓分（如「13分贏」或空）
        away_score: 客隊分數
        home_score: 主隊分數
        away_ratio: 客隊比例（百分比數字，如 44）
        home_ratio: 主隊比例（百分比數字，如 56）
    
    Returns:
        'O' 或 'X' 或 ''（空白）
    """
    # 檢查讓分和分數（注意：0 分是有效分數，不能視為空值）
    def _is_missing_score(v):
        if v is None:
            return True
        if isinstance(v, str) and not v.strip():
            return True
        return False

    if (not away_handicap and not home_handicap) or _is_missing_score(away_score) or _is_missing_score(home_score):
        return ''
    
    # 轉換比例為數字（百分比數字，如 54 代表 54%）
    try:
        if away_ratio is None or home_ratio is None:
            return ''
        
        if isinstance(away_ratio, (int, float)):
            away_ratio_num = float(away_ratio)
        elif isinstance(away_ratio, str):
            away_ratio_str = away_ratio.replace('%', '').strip()
            away_ratio_num = float(away_ratio_str)
        else:
            return ''
        
        if isinstance(home_ratio, (int, float)):
            home_ratio_num = float(home_ratio)
        elif isinstance(home_ratio, str):
            home_ratio_str = home_ratio.replace('%', '').strip()
            home_ratio_num = float(home_ratio_str)
        else:
            return ''
    except Exception as e:
        return ''
    
    # 轉換分數為整數
    try:
        away_score_int = int(away_score)
        home_score_int = int(home_score)
    except:
        return ''
    
    # 判斷讓分方是否過盤（過盤條件：實際分差 >= X+1）
    handicap_covered = False  # 讓分方是否過盤
    if away_handicap:
        handicap_match = re.search(r'(\d+)分(輸|贏)', str(away_handicap))
        if handicap_match:
            handicap_points = int(handicap_match.group(1))
            actual_diff = away_score_int - home_score_int
            handicap_covered = (actual_diff >= handicap_points + 1)
    elif home_handicap:
        handicap_match = re.search(r'(\d+)分(輸|贏)', str(home_handicap))
        if handicap_match:
            handicap_points = int(handicap_match.group(1))
            actual_diff = home_score_int - away_score_int
            handicap_covered = (actual_diff >= handicap_points + 1)
    
    # 正負盤規則：高比例方過盤→O，高比例方沒過盤→X；雙方 50% 時以讓分方為準
    if away_ratio_num > home_ratio_num:
        # 高比例方 = 客隊；高比例方過盤 = 客隊讓分時客過盤、主隊讓分時主沒過盤（客相對過盤）
        if away_handicap:
            high_ratio_covered = handicap_covered  # 客隊讓分，讓分方=客=高比例方
        else:
            high_ratio_covered = not handicap_covered  # 主隊讓分，高比例方=客，客過盤 = 主沒過盤
    elif home_ratio_num > away_ratio_num:
        # 高比例方 = 主隊
        if home_handicap:
            high_ratio_covered = handicap_covered  # 主隊讓分，讓分方=主=高比例方
        else:
            high_ratio_covered = not handicap_covered  # 客隊讓分，高比例方=主，主過盤 = 客沒過盤
    else:
        # 雙方皆 50%，以讓分方過盤與否為準
        high_ratio_covered = handicap_covered
    
    return 'O' if high_ratio_covered else 'X'


def read_games_from_daily_data(ws_daily, target_date):
    """
    從「每日數據」分頁讀取指定日期的所有比賽（一列一場，A-J）
    A=日期, B=時間, C=客分, D=客隊, E=客讓分, F=客比例, G=主分, H=主隊, I=主讓分, J=主比例
    
    Args:
        ws_daily: 「每日數據」工作表
        target_date: 目標日期（YYYYMMDD 格式字串）
    
    Returns:
        list: 比賽資料列表，每個元素包含 {date, time, away_handicap, home_handicap,
              away_score, home_score, away_ratio, home_ratio}
    """
    games = []
    for row in range(2, ws_daily.max_row + 1):
        date_value = ws_daily[f'A{row}'].value
        if not date_value:
            continue
        if isinstance(date_value, datetime):
            date_str = date_value.strftime("%Y%m%d")
        else:
            date_str = str(date_value).strip()
        if date_str != target_date:
            continue

        time_value = ws_daily[f'B{row}'].value
        away_score = ws_daily[f'C{row}'].value
        away_team = ws_daily[f'D{row}'].value
        away_handicap = ws_daily[f'E{row}'].value
        away_ratio_raw = ws_daily[f'F{row}'].value
        home_score = ws_daily[f'G{row}'].value
        home_team = ws_daily[f'H{row}'].value
        home_handicap = ws_daily[f'I{row}'].value
        home_ratio_raw = ws_daily[f'J{row}'].value

        if not away_handicap and away_team:
            m = re.search(r'(\d+)分(輸|贏)', str(away_team))
            if m:
                away_handicap = f"{m.group(1)}分{m.group(2)}"
        if not home_handicap and home_team:
            m = re.search(r'(\d+)分(輸|贏)', str(home_team))
            if m:
                home_handicap = f"{m.group(1)}分{m.group(2)}"

        away_ratio = None
        if away_ratio_raw is not None:
            if isinstance(away_ratio_raw, (int, float)):
                away_ratio = away_ratio_raw * 100 if away_ratio_raw <= 1 else away_ratio_raw
            elif isinstance(away_ratio_raw, str):
                match = re.search(r'(\d+(?:\.\d+)?)', str(away_ratio_raw))
                if match:
                    away_ratio = float(match.group(1))
        home_ratio = None
        if home_ratio_raw is not None:
            if isinstance(home_ratio_raw, str) and home_ratio_raw.startswith('='):
                if away_ratio is not None:
                    home_ratio = 100 - away_ratio
            elif isinstance(home_ratio_raw, (int, float)):
                home_ratio = home_ratio_raw * 100 if home_ratio_raw <= 1 else home_ratio_raw
            elif isinstance(home_ratio_raw, str):
                match = re.search(r'(\d+(?:\.\d+)?)', str(home_ratio_raw))
                if match:
                    home_ratio = float(match.group(1))
        if home_ratio is None and away_ratio is not None:
            home_ratio = 100 - away_ratio

        time_str = ''
        if time_value:
            time_str = time_value.strftime("%H:%M") if isinstance(time_value, datetime) else str(time_value).strip()

        games.append({
            'date': date_str,
            'time': time_str,
            'away_handicap': away_handicap,
            'home_handicap': home_handicap,
            'away_score': away_score,
            'home_score': home_score,
            'away_ratio': away_ratio,
            'home_ratio': home_ratio
        })

    games.sort(key=lambda x: x['time'] or '00:00')
    return games


def find_date_column(ws_ox, target_date_mmddyy):
    """
    在「正負盤」分頁中找到對應日期的欄位
    
    Args:
        ws_ox: 「正負盤」工作表
        target_date_mmddyy: 目標日期（MM/DD 格式，如 "10/26"）
    
    Returns:
        str: 欄位字母（如 "B"），如果找不到則返回 None
    """
    # 在第4列（日期列）中查找
    for col_idx in range(2, ws_ox.max_column + 1):  # 從B欄開始
        col_letter = get_column_letter(col_idx)
        cell_value = ws_ox[f'{col_letter}4'].value
        
        if cell_value:
            # 轉換為字串並比較
            if isinstance(cell_value, datetime):
                cell_date_str = cell_value.strftime("%m/%d")
            else:
                cell_date_str = str(cell_value).strip()
            
            if cell_date_str == target_date_mmddyy:
                return col_letter
    
    return None


def convert_date_to_mmddyy(date_str):
    """將 YYYYMMDD 轉換為 MM/DD"""
    try:
        date_obj = datetime.strptime(date_str, "%Y%m%d")
        return date_obj.strftime("%m/%d")
    except:
        return None


def convert_daily_to_ox(excel_file, target_date, show_progress=True):
    """
    將指定日期的「每日數據」轉換到「正負盤」
    
    Args:
        excel_file: Excel 檔案路徑
        target_date: 目標日期（YYYYMMDD 格式字串）
        show_progress: 是否顯示詳細進度（預設：True）
    
    Returns:
        bool: 是否成功
    """
    if not os.path.exists(excel_file):
        print(f"✗ 檔案不存在：{excel_file}")
        return False
    
    try:
        wb = load_workbook(excel_file, read_only=False)
        
        # 檢查「每日數據」分頁
        if "每日數據" not in wb.sheetnames:
            print("✗ 找不到「每日數據」分頁")
            wb.close()
            return False
        
        ws_daily = wb["每日數據"]
        
        # 檢查「正負盤」分頁
        if "正負盤" not in wb.sheetnames:
            print("✗ 找不到「正負盤」分頁")
            print("  提示：請先執行「建立NBA統計表.py」建立「正負盤」分頁")
            wb.close()
            return False
        
        ws_ox = wb["正負盤"]
        
        # 讀取該日期的所有比賽
        if show_progress:
            print(f"\n正在讀取 {target_date} 的比賽資料...")
        games = read_games_from_daily_data(ws_daily, target_date)
        
        if not games:
            if show_progress:
                print(f"  ⚠ 在「每日數據」中找不到 {target_date} 的比賽資料")
            wb.close()
            return False
        
        if show_progress:
            print(f"  ✓ 找到 {len(games)} 場比賽")
        
        # 轉換日期格式（YYYYMMDD → MM/DD）
        target_date_mmddyy = convert_date_to_mmddyy(target_date)
        if not target_date_mmddyy:
            print(f"✗ 日期格式錯誤：{target_date}")
            wb.close()
            return False
        
        # 找到對應的欄位
        target_col = find_date_column(ws_ox, target_date_mmddyy)
        if not target_col:
            if show_progress:
                print(f"  ⚠ 在「正負盤」中找不到日期 {target_date_mmddyy}")
                print(f"  提示：請確認「正負盤」分頁的第4列是否包含該日期")
            wb.close()
            return False
        
        if show_progress:
            print(f"  ✓ 找到對應欄位：{target_col} 欄（日期：{target_date_mmddyy}）")
        
        # 計算每場比賽的 O/X
        ox_results = []
        for idx, game in enumerate(games, 1):
            # 調試資訊（僅在 show_progress 時顯示）
            if show_progress:
                print(f"    第 {idx} 場（{game['time']}）：")
                print(f"      讓分：客={game['away_handicap']}, 主={game['home_handicap']}")
                print(f"      分數：客={game['away_score']}, 主={game['home_score']}")
                print(f"      比例：客={game['away_ratio']}, 主={game['home_ratio']}")
            
            ox = calculate_ox(
                game['away_handicap'],
                game['home_handicap'],
                game['away_score'],
                game['home_score'],
                game['away_ratio'],
                game['home_ratio']
            )
            ox_results.append(ox)
            if show_progress:
                print(f"      結果：{ox if ox else '(空白)'}")
        
        # 寫入「正負盤」（第 6~20 行）
        if show_progress:
            print(f"\n正在寫入「正負盤」{target_col} 欄...")
        for idx, ox_value in enumerate(ox_results):
            row_num = 6 + idx
            if row_num > 20:  # 最多15場比賽
                break
            cell = ws_ox[f'{target_col}{row_num}']
            cell.value = ox_value
            cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 清空剩餘的儲存格（如果有）
        for idx in range(len(ox_results), 15):
            row_num = 6 + idx
            cell = ws_ox[f'{target_col}{row_num}']
            cell.value = ''
        
        # 儲存檔案
        try:
            wb.save(excel_file)
            if show_progress:
                print(f"  ✓ 已將 {len(ox_results)} 場比賽的 O/X 結果寫入「正負盤」{target_col} 欄（第 6~{5+len(ox_results)} 行）")
                print(f"  ✓ 檔案已儲存")
            return True
        except PermissionError:
            print(f"  ✗ 無法儲存檔案：檔案可能正在被其他程式使用")
            print(f"  提示：請關閉 Excel 後再試")
            wb.close()
            return False
    
    except Exception as e:
        print(f"✗ 轉換時發生錯誤：{e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            wb.close()
        except:
            pass


def main():
    """主函數"""
    # 建立還原點
    create_backup()
    
    print("=" * 60)
    print("每日數據轉正負盤")
    print("=" * 60)
    
    # 檢查參數
    if len(sys.argv) < 2:
        print("用法：python 每日數據轉正負盤.py YYYYMMDD")
        print("範例：python 每日數據轉正負盤.py 20251026")
        sys.exit(1)
    
    target_date = sys.argv[1].strip()
    
    # 驗證日期格式
    if not re.match(r'^\d{8}$', target_date):
        print("✗ 日期格式錯誤！")
        print("  請使用格式：YYYYMMDD（例如：20251026）")
        sys.exit(1)
    
    # 執行轉換
    print(f"\n目標日期：{target_date}")
    print(f"Excel 檔案：{NBA_XLSX_FILE}")
    print("=" * 60)
    
    if convert_daily_to_ox(NBA_XLSX_FILE, target_date):
        print("\n" + "=" * 60)
        print("✓ 轉換完成！")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("✗ 轉換失敗！")
        print("=" * 60)
        sys.exit(1)


if __name__ == "__main__":
    main()
