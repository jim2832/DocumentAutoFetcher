from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime, timedelta
import holidays
import pandas as pd
import os
import time
import re

# 參數設定
START_DATE = "112/04/01"
END_DATE = "112/04/30"
TW_HOLIDAYS = holidays.TW()
OUTPUT_FILE = "存查結果.xlsx"

def roc_to_ad(roc_date_str):
    year, month, day = map(int, roc_date_str.split('/'))
    return datetime(year + 1911, month, day)

def working_days_diff(start_date, end_date):
    day_count = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5 and current not in TW_HOLIDAYS:
            day_count += 1
        current += timedelta(days=1)
    return day_count - 1

# 啟動 Chrome
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# 登入 EIP
driver.get("https://eip.taichung.gov.tw/myPortal.do")
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/h6/a'))).click()
wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[1]/input'))).send_keys("")
driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[2]/input').send_keys("")
input("🔐 請手動輸入驗證碼後按 Enter 繼續...")

# 點選公文整合資訊系統
driver.get("https://eip.taichung.gov.tw/myPortal.do")
gongwen_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[5]/a/div[1]')))
ActionChains(driver).move_to_element(gongwen_button).click().perform()

# 切換新視窗並進入查詢介面
original_window = driver.current_window_handle
for handle in driver.window_handles:
    if handle != original_window:
        driver.switch_to.window(handle)
        break

driver.get("https://odis.taichung.gov.tw/HOME/Home_Splitter.aspx")
wait.until(EC.element_to_be_clickable((By.XPATH, '//td[3]/table'))).click()
print("✅ 已點選『公文資料查詢』")

# 輸入查詢條件
driver.switch_to.default_content()
driver.switch_to.frame("main")
time.sleep(1)
driver.find_element(By.XPATH, "//input[contains(@name,'txtStartDate')]").clear()
driver.find_element(By.XPATH, "//input[contains(@name,'txtEndDate')]").clear()
driver.find_element(By.XPATH, "//input[contains(@name,'txtStartDate')]").send_keys(START_DATE)
driver.find_element(By.XPATH, "//input[contains(@name,'txtEndDate')]").send_keys(END_DATE)
driver.find_element(By.XPATH, '//label[contains(text(), "存查")]').click()
driver.find_element(By.XPATH, '//input[@value="查詢"]').click()

# 切換每頁 100 筆
time.sleep(1)
select = Select(driver.find_element(By.XPATH, "//select"))
select.select_by_value("100")
time.sleep(1)

# 開始走訪所有頁面與流程比對
output = []
page_number = 0
dec = 0

while True:
    print(f"\n🔍 處理第 {page_number + 1 + dec * 10} 頁")
    driver.switch_to.default_content()
    driver.switch_to.frame("main")
    file_icons = driver.find_elements(By.XPATH, '//a[img[contains(@src, "../IMAGE/GDOCSIGN_1.gif")]]')

    for i, a_tag in enumerate(file_icons):
        original_window = driver.current_window_handle
        original_windows = driver.window_handles
        a_tag.click()
        try:
            wait.until(EC.number_of_windows_to_be(len(original_windows) + 1))
            new_window = [w for w in driver.window_handles if w not in original_windows][0]
            driver.switch_to.window(new_window)

            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "GDOCDetailTopMenu")))
            driver.find_element(By.XPATH, "//td[@class='MenuChild' and contains(@url, 'TableTitle=流程')]").click()
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "GDOCDetailMain")))

            data_cells = driver.find_elements(By.CSS_SELECTOR, 'td.cssData')
            if data_cells:
                doc_number = data_cells[0].text.strip()
                unit = data_cells[2].text.strip()
                person = data_cells[3].text.strip()

            table_rows = driver.find_elements(By.XPATH, "//tr[@class='cssGridItem' or @class='cssGridAlter']")
            decision_date = None
            archive_date = None
            for row in table_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 6:
                    date_str = cells[0].text.strip().split()[0]
                    ptype = cells[4].text.strip()
                    if ptype == "決行(存查)":
                        decision_date = roc_to_ad(date_str)
                    elif ptype == "存查":
                        archive_date = roc_to_ad(date_str)

            if decision_date and archive_date:
                diff = working_days_diff(decision_date, archive_date)
                if diff > 5:
                    output.append({
                        "公文文號": doc_number,
                        "主辦單位": unit,
                        "主辦人員": person,
                        "決行(存查)日期": decision_date.strftime("%Y/%m/%d"),
                        "存查日期": archive_date.strftime("%Y/%m/%d")
                    })
        except Exception as e:
            print("❌ 發生錯誤：", e)
        finally:
            driver.close()
            driver.switch_to.window(original_window)
            driver.switch_to.default_content()
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "main")))

    # 換頁
    try:
        page_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '__doPostBack') and contains(text(), '[')]")))
        half = len(page_links) // 2
        page_numbers = [link for link in page_links[:half] if re.fullmatch(r"\[\s*\d+\s*\]", link.text.strip())]
        dots = driver.find_elements(By.XPATH, "//a[contains(text(), '[ ... ]')]")

        if page_number == 9 and dots:
            dots[-1].click()
            page_number = 0
            dec += 1
        else:
            page_numbers[page_number].click()
            page_number += 1
        time.sleep(1)
    except:
        print("✅ 已到最後一頁")
        break

# 匯出 Excel
df = pd.DataFrame(output)
if not df.empty:
    if not os.path.exists(OUTPUT_FILE):
        df.to_excel(OUTPUT_FILE, index=False)
        print("✅ 已建立新 Excel 檔案")
    else:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            book = writer.book
            sheet = book.active
            df.to_excel(writer, index=False, header=False, startrow=sheet.max_row)
            print("✅ 資料已附加到現有 Excel")
else:
    print("⚠️ 本次沒有找到超過五個工作天的資料")

input("📦 執行完畢，按 Enter 關閉瀏覽器...")
driver.quit()