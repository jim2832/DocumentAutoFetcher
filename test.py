from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time
import re
from bs4 import BeautifulSoup
import html
from datetime import datetime, timedelta
import holidays
import pandas as pd
import os
from openpyxl import load_workbook

# params
START_DATE = "112/04/01"
END_DATE = "112/04/30"

# 台灣國定假日設定
tw_holidays = holidays.TW()

# 民國年轉換函式
def roc_to_ad(roc_date_str):
    year, month, day = map(int, roc_date_str.split('/'))
    return datetime(year + 1911, month, day)

# 計算工作日差距（排除例假日與國定假日）
def working_days_diff(start_date, end_date):
    day_count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in tw_holidays:
            day_count += 1
        current += timedelta(days=1)
    return day_count - 1  # 排除起始日

# 初始化瀏覽器
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

#-----------------------------------------------------------------------------------------------------------------

# STEP 1：登入頁面
driver.get("https://eip.taichung.gov.tw/myPortal.do")

# 按下重新跳轉
wait = WebDriverWait(driver, 5)
button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/h6/a')))
button.click()

username_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[1]/input')))
password_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[2]/input')))

username_input.send_keys("") # 帳號
password_input.send_keys("") # 密碼

# TODO：這裡需手動輸入帳號、密碼與驗證碼，再點擊登入

input("🔐 登入完成後請按下 Enter 繼續...")

#-----------------------------------------------------------------------------------------------------------------

# STEP 2：回首頁後，點擊「公文整合資訊系統」圖示按鈕
driver.get("https://eip.taichung.gov.tw/myPortal.do")
wait = WebDriverWait(driver, 5)

# 等待該按鈕出現
gongwen_button = wait.until(EC.element_to_be_clickable(
    (By.XPATH, '/html/body/div[1]/div[2]/div[1]/div[2]/div/div[1]/div/div[2]/div/ul/li[5]/a/div[1]')
))

# 移動滑鼠並點擊
ActionChains(driver).move_to_element(gongwen_button).click().perform()

#-----------------------------------------------------------------------------------------------------------------

# STEP 3：點選「公文資料查詢」介面
# 模擬點擊左側選單「查詢 > 公文資料查詢」
# 切換到新開的分頁
original_window = driver.current_window_handle
for handle in driver.window_handles:
    if handle != original_window:
        driver.switch_to.window(handle)
        break

driver.get("https://odis.taichung.gov.tw/HOME/Home_Splitter.aspx")
wait = WebDriverWait(driver, 20)

# STEP 3-1：滑鼠移到「查　詢」按鈕，觸發下拉選單
menu = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div[1]/table/tbody/tr/td/table/tbody/tr[1]/td[3]/table')))
menu.click()
print("✅ 已點選『公文資料查詢』")

#-----------------------------------------------------------------------------------------------------------------

# STEP 4：輸入日期、按存查、按查詢
driver.switch_to.default_content()
driver.switch_to.frame("main")
time.sleep(1)
start = driver.find_element(By.XPATH, "/html/body/form/table[3]/tbody/tr/td/table/tbody/tr[3]/td[2]/span[1]/input")
end = driver.find_element(By.XPATH, "/html/body/form/table[3]/tbody/tr/td/table/tbody/tr[3]/td[2]/span[3]/input")

# 清除原本的值（如果有）
start.clear()
end.clear()

# 輸入新的日期
start.send_keys(START_DATE)
end.send_keys(END_DATE)

# 按下存查
label = driver.find_element(By.XPATH, '//label[contains(text(), "存查")]')
label.click()

# 按下查詢
search_button = driver.find_element(By.XPATH, "/html/body/form/table[4]/tbody/tr/td/input[3]")
search_button.click()

print("小夫我要進去了")

#-----------------------------------------------------------------------------------------------------------------

# STEP 5：每頁100筆資料
driver.switch_to.default_content()
driver.switch_to.frame("main")
time.sleep(1)
select_element = driver.find_element(By.XPATH, "/html/body/form/table[2]/tbody/tr/td/table/tbody/tr/td[2]/select")
select = Select(select_element)
select.select_by_value("100")
print("✅ 成功選擇 100 筆資料")
time.sleep(1)  # 等待畫面刷新

#-----------------------------------------------------------------------------------------------------------------

# 印出所有frame (不是iframe)
frames = driver.find_elements(By.TAG_NAME, "frame")
print("目前頁面所有 frame：")
for i, frame in enumerate(frames):
    print(f"Frame {i}: {frame.get_attribute('name')}")

# STEP 6：排序後找到電子檔(點兩下icon)
try:
    driver.switch_to.default_content()
    driver.switch_to.frame("main")
    icon = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@id, 'img_SignHead')]"))
    )
    icon.click()
    print("✅ 成功點擊 icon!")
    time.sleep(0.5)

    # 再次等待元素變可點擊（確保頁面沒變動）
    icon = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@id, 'img_SignHead')]"))
    )

    icon.click()
    print("✅ 第二次點擊成功")
    time.sleep(0.5)

except Exception as e:
    print(f"❌ 點擊 icon 失敗：{e}")


#-----------------------------------------------------------------------------------------------------------------

# STEP 7：走訪每一頁，並點擊流程視窗
page_number = 0
cur_page_number = 0
dec = 0

output = []

while True:
    print(f"\n🔍 處理第 {page_number + 1} 頁")
    driver.switch_to.default_content()
    driver.switch_to.frame("main")
    file_icons = driver.find_elements(By.XPATH, '//a[img[contains(@src, "../IMAGE/GDOCSIGN_1.gif")]]')

    for i, a_tag in enumerate(file_icons):
        original_window = driver.current_window_handle
        original_windows = driver.window_handles  # 正確儲存目前所有視窗
        popup_opened = False

        a_tag.click()
        print(f"正在處理第 {cur_page_number + 1} 個檔案")
        cur_page_number += 1

        try:
            wait.until(EC.number_of_windows_to_be(len(original_windows) + 1))
            new_window = [w for w in driver.window_handles if w not in original_windows][0]
            driver.switch_to.window(new_window)

            # 切換到上方 menu 的 frame
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "GDOCDetailTopMenu")))

            # 點擊「流程」按鈕
            process = driver.find_element(By.XPATH, "//td[@class='MenuChild' and contains(@url, 'TableTitle=流程')]")
            process.click()

            # 切換到流程資料的主 frame
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "GDOCDetailMain")))

            # 上面的資料
            data_cells = driver.find_elements(By.CSS_SELECTOR, 'td.cssData')
            if data_cells:
                doc_number = data_cells[0].text.strip()
                unit = data_cells[2].text.strip()
                person = data_cells[3].text.strip()
            
            # 下面的資料
            table_cells = driver.find_elements(By.XPATH, "//tr[@class='cssGridItem' or @class='cssGridAlter']")
            decision_date = None
            archive_date = None

            for row in table_cells:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 6:
                    process_date_str = cells[0].text.strip().split()[0]
                    process_type = cells[4].text.strip()

                    if process_type == "決行(存查)":
                        decision_date = roc_to_ad(process_date_str)
                    elif process_type == "存查":
                        archive_date = roc_to_ad(process_date_str)

            # 判斷大於五個工作天
            if decision_date and archive_date:
                diff = working_days_diff(decision_date, archive_date)
                if diff > 5:
                    print(f"⚠️ 超過五個工作天，需記錄：{doc_number}")
                    dict = {
                        "公文文號": doc_number,
                        "主辦單位": unit,
                        "主辦人員": person,
                        "決行(存查)日期": decision_date.strftime("%Y/%m/%d"),
                        "存查日期": archive_date.strftime("%Y/%m/%d"),
                    }
                    output.append(dict)

        except Exception as e:
            print("❌ 發生錯誤：", e)

        finally:
            driver.close()  # 關閉新視窗

            driver.switch_to.window(original_window)
            driver.switch_to.default_content()

            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "main")))


    # 換頁
    try:
        page_links = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//a[contains(@href, '__doPostBack') and contains(text(), '[')]")
        ))

        # 只保留像 [ 1 ]、[ 2 ] 的按鈕，排除 [>]、[ ... ] 等
        half = len(page_links) // 2
        page_numbers = [
            link for link in page_links[:half]
            if re.fullmatch(r"\[\s*\d+\s*\]", link.text.strip())
        ]

        # 額外找到 [ ... ] 的按鈕
        try:
            dots = driver.find_elements(By.XPATH, "//a[contains(@href, '__doPostBack') and contains(text(), '[ ... ]')]") #[[ ... ], [ ... ]]
        except Exception as e:
            print("❌ 找不到 [ ... ] 的按鈕")
            dots = []

        print(f"目前頁碼：{page_number + 1 + dec * 10}")

        # 如果 index 為 8，則點擊 [ ... ]
        try:
            if(page_number == 9):
                dots[-1].click()
                print("✅ 點擊 [ ... ]")
                page_number = 0
                dec += 1
            else:
                next_page = page_numbers[page_number]
                print(f"換頁：{next_page.text}")
                next_page.click()
                page_number += 1
            time.sleep(1)

        except Exception as e:
            print("❌ 換頁按鈕不存在，可能已經到最後一頁")
            break

    except Exception as e:
        print(f"❌ 換頁失敗：{e}")
        break

#-----------------------------------------------------------------------------------------------------------------

# step8: 把 output 的資料寫進 excel
output_file = "存查結果.xlsx"

# 檢查檔案是否已存在
file_exists = os.path.exists(output_file)

# 將 output 寫入 Excel
df = pd.DataFrame(output)

if not df.empty:
    if not file_exists:
        # 第一次建立檔案，寫入欄位名稱
        df.to_excel(output_file, index=False)
        print("✅ 首次建立 Excel 並寫入資料")
    else:
        # 檔案已存在，開啟並附加資料（不寫欄位）
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # 取得現有行數，從最後一列繼續寫入
            book = writer.book
            sheet = book.active
            startrow = sheet.max_row

            df.to_excel(writer, index=False, header=False, startrow=startrow)
            print(f"✅ 已將新資料附加至現有 Excel(從第 {startrow + 1} 列開始）")
else:
    print("⚠️ 本次沒有新資料要寫入 Excel")

#-----------------------------------------------------------------------------------------------------------------

input("完成後按enter結束")
driver.quit()