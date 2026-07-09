# -*- coding: utf-8 -*-
"""抓取 MLB 網路資料：活頁簿與路徑集中設定。"""

import os

# MLB 26-27 賽季主檔（與 OneDrive 資料夾 MLB26 內檔名一致）
MLB_XLSX_FILE = os.environ.get(
    "MLB_XLSX_FILE",
    r"C:\Users\curti\OneDrive\MLB26\MLB26-27數據.xlsx",
)

# 舊腳本慣用變數名，指向同一本活頁簿，減少改動範圍
NBA_XLSX_FILE = MLB_XLSX_FILE
NBA_STATS_FILE = MLB_XLSX_FILE

# 各隊正負（MLB）日期軸設定：每天一欄
TEAM_STATS_START_DATE = "20260326"
TEAM_STATS_END_DATE = "20261102"
TEAM_STATS_START_COLUMN = 16  # P 欄

# 玖九盤口（本機 OneDrive 同步路徑）
JIUJIU_XLSX_FILE = os.environ.get(
    "JIUJIU_XLSX_FILE",
    r"C:\Users\curti\OneDrive\MLB26\玖九盤口變化.xlsx",
)

# 分析母版（openpyxl 僅支援 xlsx；若需 VBA 請於 Excel 手動另存為 .xlsm）
ANALYSIS_XLSX_FILE = os.environ.get(
    "ANALYSIS_XLSX_FILE",
    r"C:\Users\curti\OneDrive\MLB26\MLB26-分析母版.xlsx",
)

# 盤口觀察（長期紀錄：模型欄 + 玖九頭尾盤口）
OBSERVATION_XLSX_FILE = os.environ.get(
    "OBSERVATION_XLSX_FILE",
    r"C:\Users\curti\OneDrive\MLB26\盤口觀察.xlsx",
)
# 舊變數名相容
ANALYSIS_XLSM_FILE = ANALYSIS_XLSX_FILE
