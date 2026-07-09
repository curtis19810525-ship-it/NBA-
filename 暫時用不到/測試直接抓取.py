"""
測試是否可以使用 requests 直接抓取 Playsport 網頁
（不使用 Selenium，速度會快很多）
"""

import requests
from bs4 import BeautifulSoup
import re

def test_direct_fetch(date_str="20260223"):
    """
    測試直接使用 requests 抓取網頁
    """
    url = f"https://www.playsport.cc/predict/scale?allianceid=3&gametime={date_str}&sid=0"
    
    print(f"測試網址: {url}")
    print("=" * 60)
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }
    
    try:
        print("方法1: 使用 requests 直接抓取...")
        response = requests.get(url, headers=headers, timeout=10)
        
        print(f"HTTP 狀態碼: {response.status_code}")
        print(f"回應長度: {len(response.text)} 字元")
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 檢查關鍵元素是否存在
            battle_links = soup.find_all('a', href=re.compile(r'/gamesData/battle'))
            print(f"找到對戰資訊連結數量: {len(battle_links)}")
            
            # 檢查表格是否存在
            tables = soup.find_all('table')
            print(f"找到表格數量: {len(tables)}")
            
            # 檢查是否有 JavaScript 動態載入的內容
            scripts = soup.find_all('script')
            print(f"找到 script 標籤數量: {len(scripts)}")
            
            # 檢查是否有包含比賽資料的 div 或 table
            team_info_cells = soup.find_all('td', class_='td-teaminfo')
            print(f"找到 td.td-teaminfo 數量: {len(team_info_cells)}")
            
            if battle_links:
                print("\n✓ 成功！可以使用 requests 直接抓取")
                print("建議：改用 requests 方法，速度會快很多（約 1-2 秒 vs 15-20 秒）")
                
                # 顯示第一個連結的父元素結構
                if battle_links[0]:
                    row = battle_links[0].find_parent('tr')
                    if row:
                        print("\n範例：第一個比賽的 HTML 結構：")
                        print(row.prettify()[:500])  # 只顯示前 500 字元
                
                return True
            else:
                print("\n⚠ 警告：找不到對戰資訊連結")
                print("可能原因：")
                print("1. 該日期沒有比賽")
                print("2. 內容需要 JavaScript 動態載入（需要 Selenium）")
                print("3. 需要登入或特殊權限")
                
                # 檢查是否有 JavaScript 載入提示
                if 'loading' in response.text.lower() or 'javascript' in response.text.lower():
                    print("\n發現 JavaScript 相關內容，可能需要 Selenium")
                
                return False
        else:
            print(f"\n✗ HTTP 錯誤: {response.status_code}")
            return False
            
    except Exception as e:
        print(f"\n✗ 發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("測試 Playsport 網頁是否可以直接用 requests 抓取\n")
    result = test_direct_fetch()
    
    print("\n" + "=" * 60)
    if result:
        print("結論：可以使用 requests 直接抓取，建議改用此方法")
    else:
        print("結論：可能需要繼續使用 Selenium（內容需要 JavaScript 載入）")
