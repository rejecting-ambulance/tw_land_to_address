import os
import json
import time
import requests
import math
from typing import Optional, Tuple
import pandas as pd

# 假設 location2latlng 模組已存在且包含 location2lat 函式
from location2latlng import location2lat

CONFIG_PATH = "config.json"

## --------------------------- API Key 載入 ---------------------------

def load_api_key() -> Optional[str]:
    """
    優先從環境變數讀取 GOOGLE_API_KEY,如果沒有,則嘗試讀取 config.json。
    """
    api_key = os.getenv("GOOGLE_API_KEY")
    if api_key:
        return api_key

    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data.get("google_api_key")
        except Exception as e:
            print(f"讀取 {CONFIG_PATH} 發生錯誤: {e}")
    return None

## --------------------------- 反向地理編碼 (Reverse Geocoding: 經緯度 -> 地址) ---------------------------

def reverse_geocode_google(lat: float, lng: float, api_key: str, language: str = "zh-TW", timeout: int = 10) -> Optional[str]:
    """呼叫 Google Geocoding API,回傳 formatted_address 或 None"""
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "latlng": f"{lat},{lng}",
        "key": api_key,
        "language": language
    }
    try:
        resp = requests.get(url, params=params, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        status = data.get("status")
        if status == "OK":
            results = data.get("results", [])
            if results:
                addr = results[0].get("formatted_address")
                # 移除地址開頭的國家名稱
                addr = addr[5:]
                return addr
            return None
        else:
            print(f"Google API 回傳狀態: {status},訊息: {data.get('error_message')}")
            return None
    except requests.RequestException as e:
        print(f"呼叫 Google Geocoding API 發生網路錯誤: {e}")
        return None

def reverse_geocode_nominatim(lat: float, lng: float, timeout: int = 10) -> Optional[dict]:
    """
    備援:使用 OpenStreetMap Nominatim 服務
    回傳字典包含原始地址和格式化地址
    """
    url = "https://nominatim.openstreetmap.org/reverse"
    params = {
        "lat": lat,
        "lon": lng,
        "format": "jsonv2",
        "addressdetails": 1,
        "zoom": 18,
        "accept-language": "zh-TW"
    }
    headers = {
        "User-Agent": "my-reverse-geocode-app/1.0 (zhandezhonghenry@gmail.com)"
    }
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        original_addr = data.get("display_name")
        formatted_addr = reverse_foreign_address(original_addr)
        return {
            "original": original_addr,  # 保留原始地址用於回轉查詢
            "formatted": formatted_addr  # 格式化地址用於顯示
        }
    except requests.RequestException as e:
        print(f"Nominatim 呼叫失敗: {e}")
        return None

def reverse_geocode_both(lat: float, lng: float, api_key: Optional[str], language: str = "zh-TW") -> dict:
    """同時呼叫 Google 和 Nominatim 反向地理編碼,回傳結果字典。"""
    results = {"google": None, "nominatim": None}

    if api_key:
        results["google"] = reverse_geocode_google(lat, lng, api_key, language)
    else:
        print("找不到 Google API Key,Google 反向地理編碼結果將為 None。")

    results["nominatim"] = reverse_geocode_nominatim(lat, lng)

    return results

def reverse_foreign_address(address: str) -> str:
    """將外國格式地址(Nominatim 回傳)轉換為台灣常見格式。"""
    if not address:
        return ""
        
    parts = [p.strip() for p in address.split(',') if p.strip()]

    # 移除最後兩個部分(郵遞區號與國家)
    if len(parts) >= 2:
        parts = parts[:-2]

    # 反轉順序
    parts.reverse()

    # 組合為台灣格式
    result = ""
    for part in parts:
        # 若為純數字(門牌號碼)
        if part.isdigit() and result: # 確保不是開頭的數字
            result += f"{part}號"
        elif part.isdigit() and not result:
            result = f"{part}" # 如果開頭是數字,先不加 '號'
        else:
            result += part

    return result

## --------------------------- 地理編碼 (Geocoding: 地址 -> 經緯度) ---------------------------

def geocode_google(address: str, api_key: str, language: str = "zh-TW", timeout: int = 10) -> Optional[Tuple[float, float]]:
    """呼叫 Google Geocoding API,將地址轉換回 (緯度, 經度) 或 None"""
    if not address: return None
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": api_key,
        "language": language
    }
    try:
        resp = requests.get(url, params=params, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") == "OK":
            results = data.get("results", [])
            if results:
                location = results[0]["geometry"]["location"]
                # 回傳 (緯度, 經度)
                return location["lat"], location["lng"]
        return None
    except requests.RequestException as e:
        print(f"呼叫 Google Geocoding API 發生網路錯誤: {e}")
        return None

def geocode_nominatim(address: str, timeout: int = 10) -> Optional[Tuple[float, float]]:
    """使用 Nominatim(OpenStreetMap)將地址轉換回 (緯度, 經度) 或 None"""
    if not address: return None
    url = "https://nominatim.openstreetmap.org/search"
    headers = {
        "User-Agent": "my-reverse-geocode-app/1.0 (zhandezhonghenry@gmail.com)"
    }
    params = {
        "q": address,
        "format": "jsonv2",
        "limit": 1,
        "accept-language": "zh-TW"
    }
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        if data:
            # Nominatim 回傳的 lat/lon 是字串,需轉成浮點數
            return float(data[0]["lat"]), float(data[0]["lon"])
        return None
    except requests.RequestException as e:
        print(f"Nominatim Geocoding 呼叫失敗: {e}")
        return None
    except (ValueError, KeyError) as e:
        print(f"Nominatim Geocoding 解析資料錯誤: {e}")
        return None

## --------------------------- 距離計算 ---------------------------

def haversine_distance(lat1: float, lng1: float, lat2: float, lng2: float) -> float:
    """
    使用 Haversine 公式計算兩組經緯度之間的距離。
    回傳單位:公尺 (m)
    """
    # 地球半徑 (km)
    R = 6371.0

    # 將經緯度從角度轉換為弧度
    lat1_rad = math.radians(lat1)
    lng1_rad = math.radians(lng1)
    lat2_rad = math.radians(lat2)
    lng2_rad = math.radians(lng2)

    # 經緯度差值
    d_lat = lat2_rad - lat1_rad
    d_lng = lng2_rad - lng1_rad

    # Haversine 公式
    a = math.sin(d_lat / 2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(d_lng / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    distance_km = R * c
    return distance_km * 1000 # 轉換為公尺 (m)

## --------------------------- 主程式執行區塊 ---------------------------

if __name__ == "__main__":

    excel_file = "locatoin2address.xlsx"
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"❌ 錯誤:找不到檔案 {excel_file}。請確認檔案是否存在。")
        exit()

    # 讀取 API Key (只需讀取一次)
    google_api_key = load_api_key()

    # 建立新的欄位
    df["原始_經度"] = None
    df["原始_緯度"] = None
    df["Google地址"] = None
    df["Nominatim地址"] = None
    df["Nominatim地址_原始"] = None  # 新增:儲存 Nominatim 原始地址
    df["Google地址_迴轉緯度"] = None
    df["Google地址_迴轉經度"] = None
    df["Nominatim地址_迴轉緯度"] = None
    df["Nominatim地址_迴轉經度"] = None
    df["Google_誤差_m"] = None      # Google 地址轉回經緯度與原經緯度的距離 (公尺)
    df["Nominatim_誤差_m"] = None   # Nominatim 地址轉回經緯度與原經緯度的距離 (公尺)


    # 從 Excel 組 data_list
    data_list = []
    # 檢查必要的欄位是否存在
    required_cols = ["縣市", "區", "段", "地號"]
    if not all(col in df.columns for col in required_cols):
        print(f"❌ 錯誤:Excel 檔案中缺少必要的欄位 ({', '.join(required_cols)})。")
        exit()

    for _, row in df.iterrows():
        data_list.append({
            "city": row["縣市"],
            "area": row["區"],
            "section": row["段"],
            "landcode": str(row["地號"])
        })
    
    # 呼叫 location2lat 取得原始經緯度
    print("⏳ 正在進行地號轉換經緯度...")
    results = location2lat(data_list)
    print("✅ 地號轉換經緯度完成。")

    # 填回 Excel
    for idx, r in enumerate(results):
        
        original_lat = r.get("緯度_WGS84")
        original_lng = r.get("經度_WGS84")
        
        # 檢查經緯度是否成功取得
        if original_lat is None or original_lng is None:
            print(f"\n跳過處理 {df.at[idx,'地號']}:無法取得經緯度。")
            continue
            
        df.at[idx, "原始_緯度"] = original_lat
        df.at[idx, "原始_經度"] = original_lng

        print(f"\n處理 {df.at[idx,'地號']} → 原始經緯度: {original_lat}, {original_lng}")
        
        # 1. 反向地理編碼 (Reverse Geocoding: 經緯度 -> 地址)
        addresses = reverse_geocode_both(original_lat, original_lng, google_api_key)
        google_addr = addresses["google"]
        nominatim_data = addresses["nominatim"]
        
        df.at[idx, "Google地址"] = google_addr
        
        # 儲存 Nominatim 的原始地址和格式化地址
        if nominatim_data:
            df.at[idx, "Nominatim地址"] = nominatim_data["formatted"]
            df.at[idx, "Nominatim地址_原始"] = nominatim_data["original"]
            nominatim_addr_for_geocoding = nominatim_data["original"]  # 使用原始地址進行回轉
        else:
            nominatim_addr_for_geocoding = None

        # 2. 地理編碼 (Geocoding: 地址 -> 經緯度) 與距離計算
        
        # --- Google 地址處理 ---
        if google_addr and google_api_key:
            re_latlng_g = geocode_google(google_addr, google_api_key)
            if re_latlng_g:
                re_lat_g, re_lng_g = re_latlng_g
                df.at[idx, "Google地址_迴轉緯度"] = re_lat_g
                df.at[idx, "Google地址_迴轉經度"] = re_lng_g
                
                # 計算距離 (公尺)
                distance_g = haversine_distance(original_lat, original_lng, re_lat_g, re_lng_g)
                df.at[idx, "Google_誤差_m"] = round(distance_g, 2)
                print(f"  > Google 地址回轉誤差: {round(distance_g, 2)} 公尺")

        time.sleep(1) # Google API 速率限制

        # --- Nominatim 地址處理 (使用原始地址) ---
        if nominatim_addr_for_geocoding:
            re_latlng_n = geocode_nominatim(nominatim_addr_for_geocoding)
            if re_latlng_n:
                re_lat_n, re_lng_n = re_latlng_n
                df.at[idx, "Nominatim地址_迴轉緯度"] = re_lat_n
                df.at[idx, "Nominatim地址_迴轉經度"] = re_lng_n

                # 計算距離 (公尺)
                distance_n = haversine_distance(original_lat, original_lng, re_lat_n, re_lng_n)
                df.at[idx, "Nominatim_誤差_m"] = round(distance_n, 2)
                print(f"  > Nominatim 地址回轉誤差: {round(distance_n, 2)} 公尺")
            else:
                print(f"  > Nominatim 無法將原始地址轉回經緯度")

        time.sleep(1) # Nominatim 速率限制

    # 調整欄位順序: 縣市、區、段、地號、Google地址、Google_誤差_m、Nominatim地址、Nominatim_誤差_m，其他欄位放後面
    desired_first = [
        "縣市",
        "區",
        "段",
        "地號",
        "Google地址",
        "Google_誤差_m",
        "Nominatim地址",
        "Nominatim_誤差_m",
        "原始_經度",
        "原始_緯度",
        "Google地址_迴轉經度",
        "Google地址_迴轉緯度",
        "Nominatim地址_原始",
        "Nominatim地址_迴轉經度",
        "Nominatim地址_迴轉緯度"
    ]

    # 只保留存在於 df 中的欄位，避免 KeyError
    desired_existing = [c for c in desired_first if c in df.columns]
    remaining = [c for c in df.columns if c not in desired_existing]
    df = df[desired_existing + remaining]

    df.to_excel(excel_file, index=False)
    print(f"\n✅ 結果已輸出到 {excel_file},包含地址迴轉誤差分析。")
    os.startfile(excel_file)