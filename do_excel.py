import xlrd
import requests
from bs4 import BeautifulSoup


def opt_excel():
    data = xlrd.open_workbook('rank_check_data_36_.xls')
    headers = {
        'Host': 'www.zhihu.com',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.114 Safari/537.36'
    }
    table = data.sheets()[0]
    for _, url in enumerate(table.col(3)):
        if isinstance(url.value, str) and str.format(url.value).__contains__('http'):
            print(url.value)
            resp = requests.get(url.value, headers)
            print(resp.text)
        else:
            print(url.value)





