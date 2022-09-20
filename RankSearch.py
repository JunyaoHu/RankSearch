from time import sleep

from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.by import By

import pandas as pd

options = EdgeOptions()
options.use_chromium = True
options.add_argument("headless")
driver = Edge(executable_path='src/msedgedriver.exe', options=options)
# driver.implicitly_wait(2)
# driver.set_window_size(width=300, height=600, windowHandle='current')

# -----------修改查询网站 -------------------
search_link = 'http://cumtcs19.chafenba.com/'
# -----------修改查询列表序号（下标从1开始） ----
index = 2
# -----------读取姓名学号表格-----------------
info_file = "jike.xlsx"
# -----------成绩输出文件名-------------------
score_file = "result.csv"

df = pd.read_excel(info_file, dtype={'学号': str})
length = len(df)
first = True
stu_info_list = []
title = ""
for i in range(length):
    name = df.iloc[i, 0]
    number = df.iloc[i, 1]

    driver.get(search_link)

    part = driver.find_element(By.CSS_SELECTOR, f"#time > option:nth-child({index})")
    part.click()
    stu_num = driver.find_element(By.CSS_SELECTOR, "#sfzh0")
    stu_num.send_keys(number)
    stu_name = driver.find_element(By.CSS_SELECTOR, "#sfzh1")
    stu_name.send_keys(name)
    driver.find_element(By.CSS_SELECTOR, "#sub").click()

    score_info = driver.find_elements(By.CSS_SELECTOR, "#main > div:nth-child(2) > table > tbody > tr > td")
    if first:
        title = [s.get_attribute("data-label") for s in score_info]
        stu_info_list.append(title)
        # print(title)
        first = False

    score = [s.text for s in score_info]
    stu_info_list.append(score)

    print(i, length, name + " ok.")

data = pd.DataFrame(stu_info_list)
data.to_csv(score_file, index=False, encoding="GBK")
