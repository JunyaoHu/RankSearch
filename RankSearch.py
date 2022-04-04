from time import sleep

from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.by import By
import numpy as np
import pandas as pd

def search(name, number):
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument("headless")
    driver = Edge(executable_path='src/msedgedriver.exe', options=options)
    driver.implicitly_wait(10)
    driver.get(r'http://cumtcs19.chafenba.com/')
    driver.set_window_size(width=300, height=600, windowHandle='current')
    # 找搜索框，输入关键词
    part = driver.find_element(By.CSS_SELECTOR, "#time > option:nth-child(4)")
    part.click()
    stu_num = driver.find_element(By.CSS_SELECTOR, "#sfzh0")
    stu_num.send_keys(number)
    stu_name = driver.find_element(By.CSS_SELECTOR, "#sfzh1")
    stu_name.send_keys(name)
    driver.find_element(By.CSS_SELECTOR, "#sub").click()

    stu_info = driver.find_elements(By.CSS_SELECTOR,"#main > div:nth-child(2) > table > tbody > tr > td")
    stu_info_list =[]
    for i in stu_info:
        stu_info_list.append(i.text)

    data = pd.DataFrame(stu_info_list)
    data = pd.DataFrame(data.values.T, index=data.columns, columns=data.index)
    data.to_csv('result.csv', index=False, encoding="GBK", mode='a', header=False)

file = u"rank_jike.xlsx"
df = pd.read_excel(file,  dtype={'学号': str})
length = len(df)
for i in range(length):
    name = df.iloc[i, 0]
    number = df.iloc[i, 1]
    search(name, number)
    print(i, length, name + " ok.")