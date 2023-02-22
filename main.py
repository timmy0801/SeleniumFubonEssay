from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
driver =''
import time
import openpyxl

def run_webdriver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('detach', True)
    options.add_argument('--start-maximized')  # 視窗最大化
    s = Service("/Users/timmylinlin/PycharmProjects/GetFubonEssay/chromedriver_mac64/chromedriver.exe")
    driver = webdriver.Chrome(service=s,options=options)
    driver.get("https://www.cmathesis.org.tw/awarded_paper")
    # 等待list-row 出現 網頁載入較慢
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "list-row"))
    )

    res = []

    total = driver.find_elements(By.CLASS_NAME,"openBtn")
    count = len(total)
    for i in range(1, count):
        total[i].click()
        time.sleep(1)

    best = driver.find_elements(By.CSS_SELECTOR, "a[class^='nav-link award_category_li']")
    for b in best:
        if b.text == "優勝論文" or b.text=="優等" or b.text=="特優":
            b.click()
            time.sleep(5)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "list-row"))
            )
            titles = driver.find_elements(By.CLASS_NAME, "list-row")
            for title in titles:
                res.append(title.text)
    wb = openpyxl.Workbook()  # 建立空白的 Excel 活頁簿物件
    s1 = wb.create_sheet("essay")
    s1.append(['姓名', '學校系所', '論文名稱', '指導教授'])
    a = []
    for d in res:
        if d != "":
            a.append(d.split("\n", 4))
    for d in a:
        s1.append(d)
    wb.save('Essay.xlsx')
    print(res)






if __name__ == '__main__':
    run_webdriver()



# total = driver.find_elements_by_class_name("openBtn")
# count = len(total)
# for i in range(1,count):
#     total[i].click()
#     time.sleep(1)

# link = driver.find_element_by_css_selector("a[data-id='14']")
# link.click()
# link = driver.find_elements_by_css_selector("a[data-id]")
# for lin in link:
#     if lin.text == "優勝論文":
#         lin.click()
#         time.sleep(1)
# time.sleep(5)
