import json

import requests
from bs4 import BeautifulSoup

from directory import *

URL = 'http://indicators.miccedu.ru/monitoring/2019/?m=vpo'
ID = 'tregion'
HEADERS = {
    'user-agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/81.0.4044.129 Safari/537.36 OPR/68.0.3618.63',
    'accept': '*/*'}


def parse_main(url, teg):
    r = requests.get(url, headers=HEADERS)
    r.encoding = 'cp1251'
    html = r
    if html.status_code == 200:
        soup = BeautifulSoup(html.text, 'html.parser')
        table = soup.find("table", id=teg)
        table_rows = table.find_all('tr')
        global listRegion
        listRegion = []
        global listCity
        listCity = []
        global listCityHaos
        allList = {}
        listCityHaos = []
        for tr in table_rows:
            td = tr.find_all('p', class_='MsoListParagraph')
            for i in td:
                listRegion.append(i.get_text(strip=True))
        for tr_p in table_rows:
            td_p = tr_p.find_all('p')
            for i in td_p:
                listCityHaos.append(i.get_text(strip=True))
                if listRegion.count(i.get_text(strip=True)) < 1:
                    listCity.append(i.get_text(strip=True))
        for tr_a in table_rows:
            td = tr_a.find_all('a')
            for i in td:
                if str(i.get_text(strip=True)) in listRegion:
                    allList[str(i.get_text(strip=True))] = {
                        'link': URL_MAIN + str(i.get('href')),
                        'type': '1'
                    }
                else:
                    allList[str(i.get_text(strip=True))] = {
                        'link': URL_MAIN + str(i.get('href')),
                        'type': '0'
                    }
        return allList
    else:
        print('Error parse_main()')


def toJSON(filename, var):
    with open('src/' + filename, 'w', encoding="utf-8") as fp:
        json.dump(var, fp, ensure_ascii=False)


toJSON('cityAndRegions.json', parse_main(URL, ID))

