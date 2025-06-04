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

"""
ahya0201
gkzBBca0440@
"""

# 初始化瀏覽器
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

# STEP 1：登入頁面
driver.get("https://eip.taichung.gov.tw/myPortal.do")

# 按下重新跳轉
wait = WebDriverWait(driver, 5)
button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/h6/a')))
button.click()

username_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[1]/input')))
password_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[1]/div[2]/form/ul[1]/li[2]/input')))

username_input.send_keys("ahya0201")
password_input.send_keys("gkzBBca0440@")

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
time.sleep(2)

#-----------------------------------------------------------------------------------------------------------------

# STEP 3：點選「公文資料查詢」介面
# 模擬點擊左側選單「查詢 > 公文資料查詢」
# 切換到新開的分頁
original_window = driver.current_window_handle
for handle in driver.window_handles:
    if handle != original_window:
        driver.switch_to.window(handle)
        break

print("目前網址：", driver.current_url)

driver.get("https://odis.taichung.gov.tw/HOME/Home_Splitter.aspx")
wait = WebDriverWait(driver, 20)

# STEP 3-1：滑鼠移到「查　詢」按鈕，觸發下拉選單
menu = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/div[3]/div[1]/table/tbody/tr/td/table/tbody/tr[1]/td[3]/table')))
menu.click()
print("✅ 已點選『公文資料查詢』")

#-----------------------------------------------------------------------------------------------------------------

# STEP 4：手動輸入日期、按存查、按查詢
input("🔐 查詢完成後請按下 Enter 繼續...")
print("小夫我要進去了")

#-----------------------------------------------------------------------------------------------------------------

# STEP 5：每頁100筆資料
driver.switch_to.default_content()
driver.switch_to.frame("main")
time.sleep(1)
select_element = driver.find_element(By.ID, "PAGE_HEAD1_ddlPAGE_COUNT")
select = Select(select_element)
select.select_by_value("100")
print("✅ 成功選擇 100 筆資料")
time.sleep(2)  # 等待畫面刷新

#-----------------------------------------------------------------------------------------------------------------

# STEP 6：排序後找到電子檔(點兩下icon)
try:
    driver.switch_to.default_content()
    driver.switch_to.frame("main")
    icon = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@id, 'img_SignHead')]"))
    )
    icon.click()
    print("✅ 成功點擊 icon!")

    time.sleep(2)

    # 再次等待元素變可點擊（確保頁面沒變動）
    icon = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='image' and contains(@id, 'img_SignHead')]"))
    )
    icon.click()
    print("✅ 第二次點擊成功")

except Exception as e:
    print(f"❌ 點擊 icon 失敗：{e}")

time.sleep(1)

#-----------------------------------------------------------------------------------------------------------------

# step 7: 走訪每一個頁面
wait = WebDriverWait(driver, 10)
driver.switch_to.default_content()
driver.switch_to.frame("main")
data_list = []

# 只點 1 到 10 頁
current_index = 0
max_pages = 9

while True:
    # 找出所有 <a><img src="../IMAGE/GDOCSIGN_1.gif"></a>
    elements = driver.find_elements(By.XPATH, '//a[img[contains(@src, "../IMAGE/GDOCSIGN_1.gif")]]')
    print(f"🔍 第 {current_index+1} 頁：找到 {len(elements)} 個電子檔")

    for i, a_tag in enumerate(elements):
        print(f"✅ 點擊第 {i+1} 個 <a>")
        a_tag.click()
        time.sleep(1)  # 給 iframe 彈窗載入時間

        try:
            # 切到彈窗 iframe（實際名稱依你找到的為主）
            driver.switch_to.default_content()
            driver.switch_to.frame("Popup_Page")  # ← 改成你實際的 iframe 名稱

            print(f"👁️ 正在瀏覽第 {i+1} 個彈窗內容")
            time.sleep(1)  # 模擬查看

        except Exception as e:
            print(f"❌ 找不到彈窗 iframe：{e}")

        # 等待該 td 元素可點擊
        wait = WebDriverWait(driver, 10)
        target_td = wait.until(EC.element_to_be_clickable((By.ID, "td_a")))

        # 點擊該 td，執行 funOpen(this)
        target_td.click()

        # 切回主頁面 iframe（例如 main）
        driver.switch_to.default_content()
        driver.switch_to.frame("main")
        print(f"🔙 回到主畫面")

        time.sleep(1)

    # 換頁處理
    try:
        page_links = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//a[contains(@href, '__doPostBack') and contains(text(), '[')]")
        ))
        number_pages = [link for link in page_links if link.text.strip() != "[ ... ]"]

        if current_index >= len(number_pages) or current_index >= max_pages:
            print("✅ 完成所有頁面")
            break

        # 點擊下一頁
        link_to_click = number_pages[current_index]
        print(f"➡️ 點擊頁碼 {link_to_click.text}")
        link_to_click.click()
        time.sleep(2)

        driver.switch_to.default_content()
        driver.switch_to.frame("main")

        current_index += 1

    except Exception as e:
        print(f"❌ 換頁失敗：{e}")
        break

#-----------------------------------------------------------------------------------------------------------------

input("完成後按enter結束")
driver.quit()


"""
112/03/20
112/03/31
"""

# https://odis.taichung.gov.tw/USER/USER_SHOW_FRAMESET.aspx?GDOC_PK=lnw7xLfcZrYPymRdp7uo2g%3d%3d&GOV_FK=rS2AuNUdjdme1Ja3xcKJDA%3d%3d;TableTitle=%E6%B5%81%E7%A8%8B
# https://odis.taichung.gov.tw/USER_SHOW_GDOCFLOW.aspx?GDOC_PK=lnw7xLfcZrY7gl%2bo3AZSkA%3d%3d&amp;GOV_FK=rS2AuNUdjdme1Ja3xcKJDA%3d%3d&amp;TableTitle=流程