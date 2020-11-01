from datetime import datetime
import json
import logging
import os
import random
import string
import shutil
from tqdm import tqdm
import requests
from bs4 import BeautifulSoup

URL = 'http://indicators.miccedu.ru/monitoring/2019/_vpo/material.php?type=2&id=10204'
ID = 'an'
HEADERS = {
    'user-agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/81.0.4044.129 Safari/537.36 OPR/68.0.3618.63',
    'accept': '*/*'}

logging.basicConfig(filename="vkbot.log", level=logging.INFO)
logging.info("Start Analika: " + str(datetime.now()))


def randomString(size):
    return ''.join(random.choice(string.ascii_letters + string.digits) for _ in range(size))


def parse_vuz():
    global folderName
    shutil.rmtree('src/json/vuzs/')
    shutil.rmtree('src/json/by_city/')
    shutil.rmtree('src/json/by_fo/')
    with open('src/cityAndRegions.json', encoding="utf-8") as json_file:
        dataVuzs = json.load(json_file)
    intVuzIdAll = []
    linkVuzList = []
    for x in dataVuzs:
        link = str(dataVuzs[x]['link'])
        intVuzIdAll.append(int(link.split('&id=')[1]))
        linkVuzList.append(str(dataVuzs[x]['link']))
    os.mkdir('src/json/vuzs/')
    os.mkdir('src/json/by_city/')
    os.mkdir('src/json/by_fo/')
    with tqdm(total=len(linkVuzList)) as pbar:
        for y in linkVuzList:
            try:
                r = requests.get(y, headers=HEADERS)
                r.encoding = 'cp1251'
                html = r
                vuzsList = {}
                if html.status_code == 200:
                    soup = BeautifulSoup(html.text, 'html.parser')
                    table = soup.find_all("table")
                    table_1 = table[1]
                    table_rows = table_1.find_all('tr')
                    for tr in table_rows:
                        td = tr.find_all('a')
                        for i in td:
                            if i.get('href') == 'http://www.miccedu.ru/' or i.get(
                                    'href') == 'mailto:general@miccedu.ru':
                                pass
                            else:
                                if (int(y.split('&id=')[1])) / 100 < 1:
                                    folderName = 'by_fo'
                                    vuzLink = i.get('href')
                                    vuzsList['vuz_' + vuzLink.split('?id=')[1]] = {
                                        'vuz_name': str(i.get_text()),
                                        'vuz_link': str(i.get('href')),
                                        'region': (str(y.split('&id=')[1])),
                                        'vuz_id': vuzLink.split('?id=')[1]
                                    }
                                else:
                                    folderName = 'by_city'
                                    vuzLink = i.get('href')
                                    vuzsList['vuz_' + vuzLink.split('?id=')[1]] = {
                                        'vuz_name': str(i.get_text()),
                                        'vuz_link': str(i.get('href')),
                                        'region': (str(y.split('&id=')[1])),
                                        'vuz_id': vuzLink.split('?id=')[1]
                                    }
                    with open(('src/json/' + folderName + '/' + randomString(9) + '.json'), 'w',
                              encoding="utf-8") as fp:
                        json.dump(vuzsList, fp, ensure_ascii=False)
                    pbar.update(1)
                else:
                    print('Error parse_main()')
            except Exception as e:
                logging.error(str(datetime.now()) + " " + str(e) + str(" : ") + str(y))
                pbar.update(1)


parse_vuz()
