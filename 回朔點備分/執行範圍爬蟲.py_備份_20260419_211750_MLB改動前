"""
批量執行爬蟲程式，處理日期範圍
"""

import sys
import os
import subprocess
from datetime import datetime, timedelta

def run_spider_for_date(game_date):
    """執行指定日期的爬蟲"""
    print(f"\n{'='*60}")
    print(f"正在處理日期：{game_date}")
    print(f"{'='*60}")
    
    try:
        # 執行爬蟲腳本
        script_path = os.path.join(os.path.dirname(__file__), "使用Selenium完整版.py")
        result = subprocess.run(
            [sys.executable, script_path, game_date],
            cwd=os.path.dirname(__file__),
            capture_output=False,  # 顯示輸出
            text=True
        )
        
        if result.returncode == 0:
            print(f"✓ {game_date} 執行成功")
            return True
        else:
            print(f"✗ {game_date} 執行失敗（返回碼：{result.returncode}）")
            return False
    except Exception as e:
        print(f"✗ {game_date} 執行時發生錯誤：{e}")
        return False


def main():
    """主函數：處理日期範圍"""
    if len(sys.argv) < 3:
        print("錯誤：請提供起始日期和結束日期")
        print("用法：python 執行範圍爬蟲.py YYYYMMDD YYYYMMDD")
        print("範例：python 執行範圍爬蟲.py 20251022 20251215")
        sys.exit(1)
    
    start_date_str = sys.argv[1].strip()
    end_date_str = sys.argv[2].strip()
    
    # 驗證日期格式
    try:
        start_date = datetime.strptime(start_date_str, "%Y%m%d")
        end_date = datetime.strptime(end_date_str, "%Y%m%d")
    except ValueError:
        print("錯誤：日期格式不正確！")
        print("請使用格式：YYYYMMDD（例如：20251022）")
        sys.exit(1)
    
    # 驗證日期順序
    if start_date > end_date:
        print("錯誤：起始日期不能晚於結束日期！")
        sys.exit(1)
    
    # 計算總天數
    total_days = (end_date - start_date).days + 1
    current_date = start_date
    
    print("\n" + "="*60)
    print("批量執行爬蟲程式")
    print("="*60)
    print(f"起始日期：{start_date_str} ({start_date.strftime('%Y/%m/%d')})")
    print(f"結束日期：{end_date_str} ({end_date.strftime('%Y/%m/%d')})")
    print(f"總共需要處理：{total_days} 天")
    print("="*60)
    
    # 統計資訊
    success_count = 0
    fail_count = 0
    failed_dates = []
    
    # 逐日執行爬蟲
    day_num = 0
    while current_date <= end_date:
        day_num += 1
        date_str = current_date.strftime("%Y%m%d")
        
        print(f"\n進度：{day_num}/{total_days} ({day_num*100//total_days}%)")
        
        if run_spider_for_date(date_str):
            success_count += 1
        else:
            fail_count += 1
            failed_dates.append(date_str)
        
        # 移至下一天
        current_date += timedelta(days=1)
    
    # 顯示最終結果
    print("\n" + "="*60)
    print("執行結果統計")
    print("="*60)
    print(f"總共處理：{total_days} 天")
    print(f"成功：{success_count} 天")
    print(f"失敗：{fail_count} 天")
    
    if failed_dates:
        print(f"\n失敗的日期：")
        for date in failed_dates:
            print(f"  - {date}")
    
    print("="*60)


if __name__ == "__main__":
    main()





































