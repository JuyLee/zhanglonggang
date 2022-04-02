from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import xlrd
import pandas as pd

import re


class zhihu_anser():

    def __init__(self):
        self.driver = webdriver.Chrome()
        self.wait = WebDriverWait(self.driver, 30)  # 设置超时时间

    def get_pagesource(self, url):
        self.driver.get(url=url)
        self.driver.maximize_window()
        # time.sleep(5)
        self.driver.find_element(By.XPATH, '/html/body/div[4]/div/div/div/div[2]/button').click()
        content_click = '#root > div > main > div > div > div.Question-main > div.ListShortcut > div > div.Card.AnswerCard.css-0 > div > div > div > div.RichContent.RichContent--unescapable > div.RichContent-inner > span'
        complete_content = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, content_click)))
        time.sleep(5)
        return complete_content.text

    def handle_content(self, content_info):
        content_f = "".join(content_info)
        content_reg = re.compile(r'<[^>]+>', re.S)
        content = content_reg.sub('', content_f)
        return content


def opt_excel():
    pass


if __name__ == '__main__':
    filename = './rank_check_data_36_.xls'
    wb = xlrd.open_workbook(filename)
    result = []
    table = wb.sheets()[0]
    data = list(enumerate(table.col(3)))
    data.pop(0)
    for i, url in data:
        if isinstance(url.value, str) and str.format(url.value).__contains__('http'):
            z = zhihu_anser()
            data = z.get_pagesource(url.value)
            z.driver.close()
            result.append(data)
        else:
            result.append(url.value)
        if i == 5:
            break

    df = pd.read_excel(filename)
    df.insert(8, 'answer', pd.Series(result))

    df.to_excel('test.xls')
