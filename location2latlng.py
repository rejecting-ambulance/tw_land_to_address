import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re


def set_chrome_options(headless: bool = False):
    chrome_opts = Options()
    if headless:
        chrome_opts.add_argument("--headless=new")  # Selenium 4.12+ 建議 new headless
        chrome_opts.add_argument("--disable-gpu")
    chrome_opts.add_argument("--start-maximized")
    # 可在此加入更多 options (如 user-agent, disable-infobars...)

    # 自動下載並啟用 chromedriver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_opts)
    return driver

def initialize_web(driver, url: str):
    try:
        driver.get(url)
        # 等待 1 秒（題目要求）
        time.sleep(1)

        # 使用 ActionChains 傳送 ESC 鍵（按 5 次）
        actions = ActionChains(driver)
        for _ in range(5):
            actions.send_keys(Keys.ESCAPE).perform()
            time.sleep(0.08)  # 秒數可調，避免按太快

        # 如果你想確認頁面狀態，可加些檢查 (例如抓 title)
        # print("完成：按下 Esc 5 次。頁面 title:", driver.title)
    except Exception as e:
        print("發生錯誤：", e)



def location2lat_chrome(driver, data_list):
    """
    data_list: list of dict, 每個 dict 包含 city、area、section、landcode
    範例: [{"city":"桃園市","area":"中壢區","section":"大路段","landcode":"815"}]

    回傳: list of dict，每個 dict 是 parse_land_info 的結果
    """
    wait = WebDriverWait(driver, 10)
    results = []  # 用來收集每筆查詢結果的 dict

    try:
        # 打開查詢頁面
        driver.find_element(By.XPATH, '//*[@id="map_header"]/div[4]/ul[3]/li/a').click()
        time.sleep(0.5)

        for data in data_list:
            city_name = data.get("city", "")
            area_name = data.get("area", "")
            section_name = data.get("section", "")
            landcode_val = data.get("landcode", "")

            # 選擇縣市
            county_select_elem = driver.find_element(By.ID, "city")
            county_select = Select(county_select_elem)
            county_select.select_by_visible_text(city_name)
            time.sleep(0.5)

            # 選擇區域
            city_select_elem = driver.find_element(By.ID, "area_office")
            city_select = Select(city_select_elem)
            city_select.select_by_visible_text(area_name)
            time.sleep(0.5)

            # 選擇段名
            location_select_elem = driver.find_element(By.XPATH, '//*[@id="submenu_pos"]/table/tbody/tr[2]/td[2]/span/span[1]/span')
            location_select_elem.click()
            search_field = driver.find_element(By.CLASS_NAME, "select2-search__field")
            search_field.clear()
            search_field.send_keys(section_name)
            search_field.send_keys(Keys.ENTER)
            time.sleep(0.5)

            # 輸入地號
            landcode_elem = driver.find_element(By.ID, "landcode")
            landcode_elem.click()
            landcode_elem.clear()
            landcode_elem.send_keys(landcode_val)
            time.sleep(0.5)

            # 按送出
            submit_btn = driver.find_element(By.ID, "div_cross_query")
            submit_btn.click()
            print(f"查詢 {city_name} {area_name} {section_name} {landcode_val}")
            time.sleep(1)

            # 等待查詢結果
            wait.until(EC.presence_of_element_located((By.ID, "DMAPS_Info")))
            actions = ActionChains(driver)
            for _ in range(3):
                actions.send_keys(Keys.ESCAPE).perform()
                time.sleep(0.08)

            query_exist(driver)  # 你的檢查函式

            # 找所有 div_cross(詳細按鈕)
            div_cross_list = driver.find_elements(By.XPATH, '//*[@id="div_cross"]')
            if div_cross_list:
                last_div = div_cross_list[-1]  # 取最後一個
                try:
                    # 按下 ESC 鍵關閉可能的彈跳視窗
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.08)

                    # 點擊最後一個 div_cross 裡的第一個按鈕
                    button = last_div.find_element(By.XPATH, './/input[@type="button"][1]')
                    button.click()

                    # 等待詳細資訊出現
                    div_imfo = wait.until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="qryLand_tab1"]/table/tbody/tr[1]/td'))
                    )
                    # 解析文字成 dict
                    div_imfo_dict = parse_land_info(div_imfo.text)
                    results.append(div_imfo_dict)  # 收集結果
                    time.sleep(1)

                except Exception as e:
                    print(f"最後一個 div_cross 找不到按鈕，錯誤: {e}")

    except Exception as e:
        print("發生錯誤：", e)
    finally:
        # input("按 Enter 鍵關閉瀏覽器...")
        driver.quit()
    return results  # 回傳整理好的 dict 列表

def query_exist(driver):
    try:
        driver.find_element(By.XPATH, '//*[@id="map_header"]/div[4]/ul[3]/li/a').click()
        wait = WebDriverWait(driver, 3)  # 最多等 3 秒
        wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="city"]')))
        # print("查詢框存在")
    except Exception as e:
        driver.find_element(By.XPATH, '//*[@id="map_header"]/div[4]/ul[3]/li/a').click()
        # print("查詢框不存在，重新開啟查詢視窗")
        
import re

def parse_land_info(text: str) -> dict:
    """
    將查詢結果文字解析成欄位字典
    
    text: 查詢結果整段文字
    回傳: dict，欄位如下：
        行政區、經度_WGS84、緯度_WGS84、經緯度_DMS、
        國土利用、TWD97_E、TWD97_N、所屬所、地段、地號
    """
    result = {}

    # 行政區
    match = re.search(r'行政區:(.+)', text)
    if match:
        result['行政區'] = match.group(1).strip()

    # WGS84
    match = re.search(r'經緯度WGS84:(\d+\.\d+),(\d+\.\d+)', text)
    if match:
        result['經度_WGS84'] = float(match.group(1))
        result['緯度_WGS84'] = float(match.group(2))

    # 度分秒
    match = re.search(r'經緯度:(.+)', text)
    if match:
        result['經緯度_DMS'] = match.group(1).strip()

    # 國土利用
    match = re.search(r'國土利用現況調查:(.+)', text)
    if match:
        result['國土利用'] = match.group(1).strip()

    # TWD97
    match = re.search(r'TWD97坐標 E:(\d+\.\d+) N:(\d+\.\d+)', text)
    if match:
        result['TWD97_E'] = float(match.group(1))
        result['TWD97_N'] = float(match.group(2))

    # 所、地段、地號
    match = re.search(r'(.+所 .+)\s+(.+)\s+(\d+)地號', text)
    if match:
        result['所屬所'] = match.group(1).strip()
        result['地段'] = match.group(2).strip()
        result['地號'] = match.group(3).strip()

    return result

def location2lat(data_list):
    url = "https://maps.nlsc.gov.tw/T09/mapshow.action#"
    driver = set_chrome_options(headless=False)
    initialize_web(driver, url)
    location_dict = location2lat_chrome(driver, data_list)
    return location_dict


if __name__ == "__main__":
    
    data_list = [
    {"city": "桃園市", "area": "中壢區", "section": "大路段", "landcode": "815"},
    {"city": "桃園市", "area": "中壢區", "section": "中原段", "landcode": "1115"}
    ]

    url = "https://maps.nlsc.gov.tw/T09/mapshow.action#"
    driver = set_chrome_options(headless=False)
    initialize_web(driver, url)
    location_dict = location2lat_chrome(driver, data_list)
    print(location_dict)


    