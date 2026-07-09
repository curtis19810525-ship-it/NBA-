"""
強制結束所有 Excel 程序
"""
import subprocess
import sys

print("=" * 60)
print("強制結束所有 Excel 程序")
print("=" * 60)
print()

try:
    # 使用 taskkill 強制結束所有 Excel 程序
    print("正在結束所有 Excel 程序...")
    result = subprocess.run(
        ['taskkill', '/F', '/IM', 'EXCEL.EXE'],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='ignore'
    )
    
    if result.returncode == 0:
        print("✓ 已成功結束所有 Excel 程序")
    else:
        output = result.stdout + result.stderr
        if "找不到處理程序" in output or "not found" in output.lower():
            print("✓ 沒有 Excel 程序在執行")
        else:
            print(f"⚠ 執行結果: {output}")
    
    print("\n請等待 2 秒，讓系統釋放檔案鎖定...")
    import time
    time.sleep(2)
    
except Exception as e:
    print(f"❌ 發生錯誤: {e}")
    import traceback
    traceback.print_exc()

input("\n按 Enter 鍵結束...")














