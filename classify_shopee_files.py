"""
Shopee 檔案自動分類腳本
========================

功能說明：
    此腳本用於自動分類 Shopee 平台的訂單資料與廣告資料檔案。
    從「00_Unclassified」資料夾讀取檔案，根據檔名格式自動分類至對應的店家資料夾。
    
    程式啟動時會自動檢查伺服器連線狀態，並提供詳細的錯誤診斷與解決建議。

支援的檔案類型：
    1. 訂單檔案 (.xlsx)
       - 檔名格式：店名_ID_報表類型_YYYYMMDD_...
       - 範例：萌寵要當家_SH0001_訂單_20240115.xlsx
       - 識別條件：副檔名為 .xlsx 且檔名包含 "_sh" 或 "order" (不分大小寫)
       - 分類至：01_Order_Data/{店家資料夾}/{年份}/{月份}/
    
    2. 廣告檔案 (.csv)
       - 檔名格式：店名-蝦皮廣告_YYYY_MM_DD.csv 或 店名_蝦皮廣告_YYYY_MM_DD.csv
       - 範例：萌寵要當家-蝦皮廣告_2024_01_15.csv
       - 識別條件：副檔名為 .csv 且檔名包含 "蝦皮廣告" (不分大小寫)
       - 分類至：02_Ads_Data/{店家資料夾}/{年份}/{月份}/

目錄結構：
    BASE_DIR (\\\\Ycpl-server\\a02_電商部\\03_平台管理\\Online_Platform_Data\\01_Shopee)
    ├── 00_Unclassified/          # 待分類檔案來源資料夾
    ├── 01_Order_Data/            # 訂單資料分類目標
    │   └── {店家資料夾}/
    │       └── {年份}/
    │           └── {月份}/
    ├── 02_Ads_Data/              # 廣告資料分類目標
    │   └── {店家資料夾}/
    │       └── {年份}/
    │           └── {月份}/
    ├── 99_無法辨識/              # 無法分類的檔案
    └── latest_sort_log.txt       # 分類日誌檔案

處理邏輯：
    1. 檢查伺服器連線狀態（自動偵測防毒阻擋、網路問題、權限問題）
    2. 讀取「00_Unclassified」資料夾中的所有檔案
    3. 根據副檔名與檔名關鍵字判斷檔案類型
    4. 解析檔名中的店家資訊（ID 或名稱）
    5. 解析檔名中的日期資訊（年份、月份）
    6. 移動檔案至對應的目標資料夾
    7. 無法辨識的檔案移至「99_無法辨識」資料夾

特殊處理：
    - 訂單檔案日期解析：使用 re.findall 抓取所有符合的日期（格式：20xxxxxx），
      若有多個日期則取最後一個，適用於檔名如 "...20251104_20251204.xlsx" 的情況
    - 廣告檔案分隔符號：支援 "-蝦皮廣告" 或 "_蝦皮廣告" 兩種格式
    - 廣告檔案店家識別：會嘗試將店名部分以底線分割，逐一檢查每個部分是否在 Mapping 中

伺服器連線檢查：
    - 自動偵測防毒軟體阻擋 UNC 路徑：檢查伺服器根目錄是否可存取
    - 網路連線診斷：解析伺服器名稱，判斷網路或權限問題
    - 權限檢查：驗證路徑是否存在且可讀取
    - 錯誤分類：將錯誤分為防毒阻擋、網路問題、權限不足、未知錯誤等類型
    - 提供針對性的解決建議：根據錯誤類型顯示具體的排查步驟

錯誤處理：
    - 伺服器連線失敗：顯示詳細錯誤訊息和解決建議，程式無法繼續執行
    - 檔案已存在：跳過，不覆蓋（記錄 [SKIP]）
    - 檔案被佔用：記錄錯誤（通常為 Excel 檔案開啟中，記錄 [LOCKED]）
    - 格式不符：移至「99_無法辨識」資料夾（記錄 [UNKNOWN]）
    - 日期解析失敗：移至「99_無法辨識」資料夾（記錄 [UNKNOWN]）
    - 店家 Mapping 不存在：移至「99_無法辨識」資料夾（記錄 [UNKNOWN]）
    - Log 檔初始化失敗：顯示錯誤訊息並結束程式
    - stdin 不可用：使用 safe_input() 函數安全處理，避免在非互動式環境中出錯

特殊功能：
    - safe_input() 函數：安全處理使用者輸入，當 stdin 不可用時（如被重定向、關閉，
      或在非互動式環境中執行）不會產生 RuntimeError，程式可正常執行
    - 非互動式執行支援：程式可在被其他程式呼叫、批次檔執行等非互動式環境中正常運作

日誌記錄：
    - 所有操作記錄於 latest_sort_log.txt
    - 同時顯示於命令列視窗
    - 每次執行會覆蓋舊的日誌檔案
    - 程式完成時會顯示日誌檔案位置

使用方法：
    Python 腳本模式：
        python classify_shopee_files.py
    
    Exe 執行檔模式（推薦）：
        1. 直接雙擊執行 classify_shopee_files.exe
        2. 或在命令列視窗中執行 classify_shopee_files.exe
        3. 程式會自動檢查伺服器連線並顯示執行結果

注意事項：
    - 執行前請確保關閉所有 Excel 檔案，避免檔案被佔用
    - 建議先備份重要檔案
    - 支援的店家清單請參考 SHOP_MAPPING 字典
    - 若檔名中包含多個日期，系統會自動選取最後一個日期進行分類
    - 若遇到防毒軟體阻擋，請將伺服器路徑加入防毒軟體白名單
    - 確保您的帳號有權限存取伺服器資料夾
    - 程式會在執行完成後暫停，方便查看結果（exe 模式，且 stdin 可用時）
    - 程式支援非互動式執行，可在批次檔、排程任務等環境中正常運作
    - 若在非互動式環境中執行，程式會自動跳過等待使用者輸入的步驟
"""

import os
import shutil
import re
import sys
import socket
from datetime import datetime

# ================================
# Shopee 店家 Mapping
# ================================
SHOP_MAPPING = {
    "SH0001": "SH0001_萌寵要當家", "萌寵要當家": "SH0001_萌寵要當家",
    "SH0002": "SH0002_汪喵日總匯", "汪喵日總匯": "SH0002_汪喵日總匯",
    "SH0003": "SH0003_毛寵星人樂園", "毛寵星人樂園": "SH0003_毛寵星人樂園",
    "SH0004": "SH0004_大尾巴寵物店", "大尾巴寵物店": "SH0004_大尾巴寵物店",
    "SH0005": "SH0005_有才寵物商店", "有才寵物商店": "SH0005_有才寵物商店",
    "SH0006": "SH0006_火箭貓狗星際站", "火箭貓狗星際站": "SH0006_火箭貓狗星際站",
    "SH0007": "SH0007_驕傲貓狗王國", "驕傲貓狗王國": "SH0007_驕傲貓狗王國",
    "SH0008": "SH0008_毛一堆寵物超市", "毛一堆寵物超市": "SH0008_毛一堆寵物超市",
    "SH0009": "SH0009_貓狗好時光", "貓狗好時光": "SH0009_貓狗好時光",
    "SH0010": "SH0010_貓狗超有事", "貓狗超有事": "SH0010_貓狗超有事",
    "SH0011": "SH0011_貓狗得來速", "貓狗得來速": "SH0011_貓狗得來速",
    "SH0012": "SH0012_維尼寵物店", "維尼寵物店": "SH0012_維尼寵物店",
    "SH0013": "SH0013_熊好命寵物店", "熊好命寵物店": "SH0013_熊好命寵物店",
    "SH0014": "SH0014_毛宇宙", "毛宇宙": "SH0014_毛宇宙",
    "SH0015": "SH0015_想寵貓狗", "想寵貓狗": "SH0015_想寵貓狗",
    "SH0016": "SH0016_LULUMI寵物館", "LULUMI寵物館": "SH0016_LULUMI寵物館",
    "SH0017": "SH0017_愛喵樂寵物", "愛喵樂寵物": "SH0017_愛喵樂寵物",
    "SH0018": "SH0018_妮卡寵物館", "妮卡寵物館": "SH0018_妮卡寵物館",
    "SH0019": "SH0019_優格小喵.", "優格小喵.": "SH0019_優格小喵.",
    "SH0020": "SH0020_寵物執行長", "寵物執行長": "SH0020_寵物執行長",
    "SH0021": "SH0021_寵物有喵病", "寵物有喵病": "SH0021_寵物有喵病",
    "SH0022": "SH0022_優選寵物", "優選寵物": "SH0022_優選寵物",
    "SH0023": "SH0023_皇家狗貓", "皇家狗貓": "SH0023_皇家狗貓",
    "SH0024": "SH0024_寵物の達人", "寵物の達人": "SH0024_寵物の達人",
    "SH0025": "SH0025_毛孩超市", "毛孩超市": "SH0025_毛孩超市",
    "SH0026": "SH0026_寵物來了~", "寵物來了~": "SH0026_寵物來了~",
    "SH0027": "SH0027_貓狗派對", "貓狗派對": "SH0027_貓狗派對",
    "SH0028": "SH0028_寵貓寵狗", "寵貓寵狗": "SH0028_寵貓寵狗",
}

# ================================
# 路徑與 Log 設定
# ================================
def check_unc_path_access(path):
    """
    檢查 UNC 路徑是否可存取，並偵測是否被防毒軟體阻擋
    
    返回: (is_accessible, error_type, error_message)
    - is_accessible: bool, 路徑是否可存取
    - error_type: str, 錯誤類型 ('network', 'permission', 'antivirus', 'unknown')
    - error_message: str, 錯誤訊息
    """
    try:
        # 方法 1: 使用 os.path.exists 檢查
        if os.path.exists(path):
            # 進一步檢查是否真的可以讀取目錄
            try:
                os.listdir(path)
                return True, None, None
            except PermissionError:
                return False, 'permission', '路徑存在但無讀取權限'
            except Exception as e:
                return False, 'unknown', f'無法讀取目錄: {str(e)}'
        
        # 方法 2: 嘗試存取父目錄來判斷是否為網路問題
        parent_path = os.path.dirname(path)
        if parent_path and parent_path != path:
            try:
                if os.path.exists(parent_path):
                    # 父目錄存在，但目標目錄不存在
                    return False, 'network', '目標資料夾不存在，但網路連線正常'
            except Exception:
                pass
        
        # 方法 3: 嘗試建立連線來偵測防毒阻擋
        try:
            # 嘗試存取伺服器根目錄
            server_root = r"\\Ycpl-server"
            if not os.path.exists(server_root):
                # 可能是防毒軟體阻擋 UNC 路徑
                return False, 'antivirus', '可能被防毒軟體阻擋 UNC 路徑存取'
        except Exception:
            pass
        
        # 方法 4: 檢查是否為 UNC 路徑格式
        if path.startswith('\\\\'):
            # 嘗試解析伺服器名稱（簡單檢查）
            try:
                server_name = path.split('\\')[2]  # 取得伺服器名稱
                socket.gethostbyname(server_name)
                # 伺服器名稱可以解析，但路徑無法存取
                return False, 'permission', '伺服器連線正常，但無權限存取該路徑'
            except socket.gaierror:
                return False, 'network', '無法解析伺服器名稱，請檢查網路連線'
            except Exception:
                pass
        
        return False, 'unknown', '無法存取路徑，原因不明'
        
    except PermissionError:
        return False, 'permission', '權限不足，無法存取該路徑'
    except OSError as e:
        if 'network' in str(e).lower() or 'network path' in str(e).lower():
            return False, 'network', '網路路徑無法存取，請檢查網路連線'
        elif 'access is denied' in str(e).lower() or '拒絕存取' in str(e):
            return False, 'permission', '存取被拒絕，請確認帳號權限'
        else:
            return False, 'unknown', f'系統錯誤: {str(e)}'
    except Exception as e:
        return False, 'unknown', f'發生未預期的錯誤: {str(e)}'

def get_base_dir():
    """
    取得基礎目錄路徑，並提供詳細的錯誤訊息
    """
    path = r"\\Ycpl-server\a02_電商部\03_平台管理\Online_Platform_Data\01_Shopee"
    
    is_accessible, error_type, error_message = check_unc_path_access(path)
    
    if is_accessible:
        return path
    
    # 顯示詳細的錯誤訊息
    print("\n" + "="*60)
    print("❌ 無法存取伺服器資料夾")
    print("="*60)
    print(f"\n目標路徑：")
    print(f"  {path}")
    print(f"\n錯誤類型：{error_type}")
    print(f"錯誤訊息：{error_message}")
    print("\n" + "-"*60)
    
    # 根據錯誤類型提供解決建議
    if error_type == 'antivirus':
        print("🔒 偵測到可能被防毒軟體阻擋 UNC 路徑")
        print("\n建議解決方法：")
        print("  1. 檢查防毒軟體設定，將以下路徑加入白名單：")
        print(f"     {path}")
        print("  2. 暫時關閉防毒軟體的即時防護功能")
        print("  3. 聯絡 IT 部門協助設定防毒軟體例外規則")
        print("  4. 確認 Windows Defender 或其他安全軟體未阻擋網路路徑")
    elif error_type == 'network':
        print("🌐 網路連線問題")
        print("\n建議解決方法：")
        print("  1. 檢查網路連線是否正常")
        print("  2. 確認是否可以存取其他網路資料夾")
        print("  3. 嘗試在檔案總管中手動連線到：\\\\Ycpl-server")
        print("  4. 確認是否在同一網域或內網環境")
    elif error_type == 'permission':
        print("🔐 權限不足")
        print("\n建議解決方法：")
        print("  1. 確認您的帳號是否有權限存取該資料夾")
        print("  2. 聯絡系統管理員確認權限設定")
        print("  3. 嘗試以系統管理員身分執行此程式")
        print("  4. 確認資料夾分享權限設定")
    else:
        print("❓ 未知錯誤")
        print("\n建議解決方法：")
        print("  1. 確認伺服器是否正常運作")
        print("  2. 檢查路徑是否正確")
        print("  3. 聯絡 IT 部門尋求協助")
    
    print("\n" + "="*60)
    print("程式無法繼續執行，請解決上述問題後再試。")
    print("="*60 + "\n")
    
    return None

# 全域變數，將在 main() 中初始化
BASE_DIR = None
UNCLASSIFIED_DIR = None
ORDER_DIR = None
ADS_DIR = None
UNKNOWN_DIR = None
LOG_FILE = None

# ================================
# Log 輔助函式
# ================================
def log(message):
    print(message)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(message + "\n")
    except Exception as e:
        print(f"!! Log寫入失敗: {e}")

def safe_input(prompt=""):
    """
    安全的輸入函數，當 stdin 不可用時不會出錯
    
    返回: bool, 是否成功等待使用者輸入
    """
    try:
        # 檢查 stdin 是否可用
        if not sys.stdin.isatty():
            # stdin 不可用（例如被重定向或關閉）
            return False
        
        # 嘗試讀取輸入
        input(prompt)
        return True
    except (RuntimeError, OSError, EOFError):
        # stdin 不可用或已關閉
        return False
    except Exception:
        # 其他未預期的錯誤
        return False

# ================================
# 檔案與目錄操作
# ================================
def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def safe_move(src_path, dest_folder, filename):
    ensure_dir(dest_folder)
    dest_path = os.path.join(dest_folder, filename)
    
    # 1. 檢查檔案是否已存在
    if os.path.exists(dest_path):
        log(f"[SKIP] 檔案已存在，跳過: {filename}")
        return

    # 2. 嘗試移動
    try:
        shutil.move(src_path, dest_path)
        log(f"[OK] 移動成功: {filename} -> {dest_folder}")
    except PermissionError:
        log(f"[LOCKED] 檔案被佔用 (請關閉 Excel): {filename}")
    except Exception as e:
        log(f"[ERROR] 移動失敗: {filename}, 錯誤: {e}")

# ================================
# 訂單分類
# ================================
def classify_order_file(filepath, filename):
    parts = filename.split("_")
    
    if len(parts) < 2:
        log(f"[UNKNOWN] 格式不符 (parts<2): {filename}")
        return move_to_unknown(filepath, filename)

    shop_name = parts[0]
    shop_id = parts[1]

    # 判斷邏輯：檢查 mapping
    folder_name = None
    if shop_id in SHOP_MAPPING:
        folder_name = SHOP_MAPPING[shop_id]
    elif shop_name in SHOP_MAPPING:
        folder_name = SHOP_MAPPING[shop_name]
    else:
        log(f"[UNKNOWN] 無對應店鋪 ID ({shop_id}) 亦無對應名稱 ({shop_name}): {filename}")
        return move_to_unknown(filepath, filename)

    # =========================================================================
    # 修改點：使用 re.findall 抓取所有符合的日期，並取「最後一個」
    # 針對檔名如 "...20251104_20251204.xlsx"，這會抓到 ['20251104', '20251204']
    # 我們取最後一個 [-1]，即 20251204
    # =========================================================================
    date_matches = re.findall(r"20\d{6}", filename)
    
    if date_matches:
        # 取最後一個日期
        target_date = date_matches[-1]
        year = target_date[:4]
        month = target_date[4:6]
        
        if int(month) < 1 or int(month) > 12:
            log(f"[UNKNOWN] 日期解析失敗 (月份無效: {month}): {filename}")
            move_to_unknown(filepath, filename)
        else:
            target_dir = os.path.join(ORDER_DIR, folder_name, year, month)
            safe_move(filepath, target_dir, filename)
    else:
        # 備用解析邏輯 (如果正則完全沒抓到)
        try:
            if len(parts) >= 4:
                 date_str = parts[3][:8]
                 if len(date_str) >= 8 and date_str.isdigit() and date_str.startswith("20"):
                    year, month = date_str[:4], date_str[4:6]
                    if int(month) < 1 or int(month) > 12:
                        log(f"[UNKNOWN] 日期解析失敗 (月份無效): {filename}")
                        move_to_unknown(filepath, filename)
                    else:
                        target_dir = os.path.join(ORDER_DIR, folder_name, year, month)
                        safe_move(filepath, target_dir, filename)
                        return

            log(f"[UNKNOWN] 日期解析失敗 (Regex 與固定位置皆無效): {filename}")
            move_to_unknown(filepath, filename)
        except (IndexError, ValueError) as e:
            log(f"[UNKNOWN] 解析日期發生錯誤: {filename}, {e}")
            move_to_unknown(filepath, filename)

# ================================
# 廣告分類
# ================================
def classify_ads_file(filepath, filename):
    separator = None
    if "-蝦皮廣告" in filename:
        separator = "-蝦皮廣告"
    elif "_蝦皮廣告" in filename:
        separator = "_蝦皮廣告"
    
    if not separator:
        log(f"[UNKNOWN] 廣告檔名格式錯誤 (找不到 '蝦皮廣告' 關鍵字或分隔符號不符): {filename}")
        return move_to_unknown(filepath, filename)
    
    try:
        prefix_part = filename.split(separator)[0]
    except IndexError:
        log(f"[UNKNOWN] 廣告檔名分割失敗: {filename}")
        return move_to_unknown(filepath, filename)

    if not prefix_part:
        log(f"[UNKNOWN] 廣告檔名格式錯誤 (店名為空): {filename}")
        return move_to_unknown(filepath, filename)

    folder_name = None
    if prefix_part in SHOP_MAPPING:
        folder_name = SHOP_MAPPING[prefix_part]
    else:
        sub_parts = prefix_part.split("_")
        for part in sub_parts:
            if part in SHOP_MAPPING:
                folder_name = SHOP_MAPPING[part]
                break
    
    if not folder_name:
        log(f"[UNKNOWN] 無對應店鋪名稱: {prefix_part} (檔名: {filename})")
        return move_to_unknown(filepath, filename)

    date_match = re.search(r"(20\d{2})_(\d{2})_\d{2}", filename)
    if date_match:
        year = date_match.group(1)
        month = date_match.group(2)
        try:
            if int(month) < 1 or int(month) > 12:
                log(f"[UNKNOWN] 廣告日期解析失敗 (月份無效): {filename}")
                move_to_unknown(filepath, filename)
            else:
                target_dir = os.path.join(ADS_DIR, folder_name, year, month)
                safe_move(filepath, target_dir, filename)
        except ValueError:
            log(f"[UNKNOWN] 廣告日期解析失敗 (月份格式錯誤): {filename}")
            move_to_unknown(filepath, filename)
    else:
        log(f"[UNKNOWN] 廣告日期解析失敗: {filename}")
        move_to_unknown(filepath, filename)

def move_to_unknown(filepath, filename):
    safe_move(filepath, UNKNOWN_DIR, filename)

# ================================
# 主程式
# ================================
def main():
    global BASE_DIR, UNCLASSIFIED_DIR, ORDER_DIR, ADS_DIR, UNKNOWN_DIR, LOG_FILE
    
    # 顯示程式啟動訊息
    print("\n" + "="*60)
    print("Shopee 檔案自動分類程式")
    print("="*60)
    print("正在檢查伺服器連線...\n")
    
    BASE_DIR = get_base_dir()
    if BASE_DIR is None:
        # 等待使用者按鍵後結束（如果是 exe 模式且 stdin 可用）
        if getattr(sys, 'frozen', False):
            safe_input("\n按 Enter 鍵結束程式...")
        return
    
    UNCLASSIFIED_DIR = os.path.join(BASE_DIR, "00_Unclassified")
    ORDER_DIR = os.path.join(BASE_DIR, "01_Order_Data")
    ADS_DIR = os.path.join(BASE_DIR, "02_Ads_Data")
    UNKNOWN_DIR = os.path.join(BASE_DIR, "99_無法辨識")
    LOG_FILE = os.path.join(BASE_DIR, "latest_sort_log.txt")
    
    print("✓ 伺服器連線成功！\n")
    
    try:
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.write(f"=== Shopee 自動分類日誌 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===\n")
    except Exception as e:
        print("❌ Log 檔初始化失敗")
        print(f"   錯誤訊息: {e}")
        print("   可能原因: Log 檔案正在被其他程式開啟")
        print("   建議: 請關閉可能開啟該檔案的程式後重試\n")
        if getattr(sys, 'frozen', False):
            safe_input("按 Enter 鍵結束程式...")
        return

    log("=== Shopee 自動分類開始 ===")
    log(f"BASE_DIR: {BASE_DIR}")
    log(f"UNCLASSIFIED_DIR: {UNCLASSIFIED_DIR}")

    if not os.path.exists(UNCLASSIFIED_DIR):
        print("\n❌ 錯誤: 找不到來源資料夾")
        print(f"   路徑: {UNCLASSIFIED_DIR}")
        print("   可能原因:")
        print("     1. 資料夾不存在")
        print("     2. 資料夾名稱或路徑已變更")
        print("     3. 權限不足無法存取該資料夾")
        print("\n請確認資料夾是否存在，或聯絡系統管理員。\n")
        log(f"錯誤: 找不到來源資料夾: {UNCLASSIFIED_DIR}")
        if getattr(sys, 'frozen', False):
            safe_input("按 Enter 鍵結束程式...")
        return

    file_list = os.listdir(UNCLASSIFIED_DIR)
    if not file_list:
        log("來源資料夾為空，無須處理。")
        return
    
    for filename in file_list:
        filepath = os.path.join(UNCLASSIFIED_DIR, filename)

        if not os.path.isfile(filepath) or filename.startswith("~$"):
            continue

        filename_lower = filename.lower()

        # 放寬條件，支援 '_sh' 或 'order' 關鍵字
        is_order_file = filename_lower.endswith(".xlsx") and (
            "_sh" in filename_lower or "order" in filename_lower
        )
        
        is_ads_file = filename_lower.endswith(".csv") and "蝦皮廣告" in filename_lower

        if is_order_file:
            classify_order_file(filepath, filename)
        
        elif is_ads_file:
            classify_ads_file(filepath, filename)
        
        else:
            log(f"[UNKNOWN] 未知類型，移至無法辨識: {filename}")
            move_to_unknown(filepath, filename)

    log("=== Shopee 自動分類完成 ===")
    
    # 顯示完成訊息
    print("\n" + "="*60)
    print("✓ 檔案分類完成！")
    print("="*60)
    print(f"\n日誌檔案位置：")
    print(f"  {LOG_FILE}")
    print("\n您可以開啟日誌檔案查看詳細處理結果。\n")
    
    # 如果是 exe 模式，等待使用者按鍵（如果 stdin 可用）
    if getattr(sys, 'frozen', False):
        safe_input("按 Enter 鍵結束程式...")

if __name__ == "__main__":
    main()