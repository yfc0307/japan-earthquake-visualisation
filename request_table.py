import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
from datetime import datetime

"""
page_num = 1

if page_num == 0:
    data_index = None
else:
    data_index = page_num * 10
source_url = f"https://typhoon.yahoo.co.jp/weather/jp/earthquake/list/?sort=1&key=1&b={data_index}1"
"""


def request_url(url):
    with requests.get(url) as r:
        soup = BeautifulSoup(r.content, "html.parser")
        print(f"Requesting {url}")
        return soup


def find_eqhist_div(soup):
    eqhist_text = []
    div = soup.find("div", id="eqhist")
    table = div.find("table")
    if table:
        for row in table.find_all("tr", attrs={"bgcolor": "#ffffff"}):
            columns = row.find_all(["th", "td"])
            eqhist_text.append([column.text for column in columns])
        print("Getting Text")
    else:
        eqhist_text.append("")

    return eqhist_text


eqhist_text_list = []
try:
    for page_num in range(1000):
        # time.sleep(1)
        if page_num == 0:
            data_index = None
        else:
            data_index = page_num * 10
        source_url = f"https://typhoon.yahoo.co.jp/weather/jp/earthquake/list/?sort=1&key=1&b={data_index}1"

        eqhist_text_list = eqhist_text_list + find_eqhist_div(request_url(source_url))
        print(eqhist_text_list[-5:])
except Exception as e:
    pass

df = pd.DataFrame(
    eqhist_text_list, columns=["Time", "Location", "Magnitude", "Seismic Intensity"]
)

df.to_excel("jp_eqhist_data.xlsx", index=False)
print(
    f"Successfully downloaded jp_eqhist_data.xlsx, contains {len(eqhist_text_list)} rows of data."
)
