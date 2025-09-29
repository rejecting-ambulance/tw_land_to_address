import os
import json
import time
import requests
from typing import Optional, Tuple

from location2lat import location2lat

CONFIG_PATH = "config.json"  # 若你用檔案存 key，放這裡

def load_api_key() -> Optional[str]:
    """
    優先從環境變數讀取 GOOGLE_API_KEY，
    如果沒有，則嘗試讀取 config.json 的 google_api_key。
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

def reverse_geocode_google(lat: float, lng: float, api_key: str, language: str = "zh-TW", timeout: int = 10) -> Optional[str]:
    """呼叫 Google Geocoding API，回傳 formatted_address 或 None"""
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
                return results[0].get("formatted_address")
            return None
        else:
            # 常見 status: ZERO_RESULTS, OVER_QUERY_LIMIT, REQUEST_DENIED, INVALID_REQUEST
            print(f"Google API 回傳狀態: {status}，訊息: {data.get('error_message')}")
            return None
    except requests.RequestException as e:
        print(f"呼叫 Google Geocoding API 發生網路錯誤: {e}")
        return None

def reverse_geocode_nominatim(lat: float, lng: float, timeout: int = 10) -> Optional[str]:
    """
    備援：使用 OpenStreetMap Nominatim（注意使用政策，不要大量呼叫）
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
        "User-Agent": "my-reverse-geocode-app/1.0 (your@email.example)"
    }
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        resp.raise_for_status()
        data = resp.json()
        return data.get("display_name")
    except requests.RequestException as e:
        print(f"Nominatim 呼叫失敗: {e}")
        return None

def reverse_geocode_both(lat: float, lng: float, language: str = "zh-TW") -> dict:
    """
    同時呼叫 Google 和 Nominatim 反向地理編碼，回傳結果字典。
    範例回傳：
    {
        "google": "台北市信義區信義路五段7號",
        "nominatim": "人行空橋, 西村里, 信義區, 信義商圈, 臺北市, 11049, 臺灣"
    }
    """
    results = {"google": None, "nominatim": None}

    # 讀取 Google API Key
    api_key = load_api_key()
    if api_key:
        results["google"] = reverse_geocode_google(lat, lng, api_key, language)
    else:
        print("找不到 Google API Key，Google 結果將為 None。")

    # 呼叫 Nominatim
    results["nominatim"] = reverse_geocode_nominatim(lat, lng)

    return results


if __name__ == "__main__":
    # 範例經緯度：台北 101
    # lat = 24.980500
    # lng = 121.201819

    data_list = [
    {"city": "桃園市", "area": "中壢區", "section": "大路段", "landcode": "815"},
    {"city": "桃園市", "area": "中壢區", "section": "中原段", "landcode": "1115"}
    ]
    results = location2lat(data_list)
    # 只取經緯度
    latlng_only = [{'lon': r['經度_WGS84'], 'lat': r['緯度_WGS84']} for r in results]

    for item in latlng_only:
        lat = item['lat']
        lng = item['lon']
        print(f"\n經緯度: {lat}, {lng}")
        addresses = reverse_geocode_both(lat, lng)
        print("=== 反向地理編碼結果 ===")
        print(f"Google:    {addresses['google']}")
        print(f"Nominatim: {addresses['nominatim']}")
