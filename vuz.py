# -*- coding: utf-8 -*-
import json
import logging
from datetime import datetime

import pandas as pd
from colorama import Fore
from tqdm import tqdm

URL_FO = 'http://indicators.miccedu.ru/monitoring/2019/_vpo/material.php?type=1&id=4'
ID = 'an'
HEADERS = {
    'user-agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/81.0.4044.129 Safari/537.36 OPR/68.0.3618.63',
    'accept': '*/*'}
logging.basicConfig(filename="vkbot.log", level=logging.INFO)
logging.info("Start Analika: " + str(datetime.now()))


def parse_toExcel():
    with open('src/vuz_merge.json', encoding="utf-8") as json_file:
        data = json.load(json_file)
    intIdAll = []
    listFileExcelE = []
    listFileExcelOd = []
    listFileExcelMd = []
    listFileExcelNd = []
    listFileExcelFe = []
    listFileExcelInf = []
    listFileExcelKadr = []
    listFileExcelRole = []
    listFileExcelForRegion = []
    listFileExcelAllInRegion = []
    listFileExcelStudentInVuz = []
    listFileExcelProcentStudent = []
    listFileExcelDolyaInRegion = []
    listFileExcelPoPerechnyam = []
    listFileExcelDop = []

    linkList = []

    for x in data:
        link = str(data[x]['vuz_link'])
        intIdAll.append(int(link.split('id=')[1]))
        linkList.append(str(data[x]['vuz_link']))
    print(Fore.GREEN + 'Парсинг Html-страниц в Excel...')
    with tqdm(total=len(linkList)) as pbar:
        for r in range(len(linkList)):
            try:
                url = linkList[r]
                table_e = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[3]
                table_e.to_excel('src/excel/vuzs/e/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelE.append('src/excel/vuzs/e/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_od = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[4]
                table_od.to_excel('src/excel/vuzs/od/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelOd.append('src/excel/vuzs/od/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_nd = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[5]
                table_nd.to_excel('src/excel/vuzs/nd/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelNd.append('src/excel/vuzs/nd/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_md = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[6]
                table_md.to_excel('src/excel/vuzs/md/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelMd.append('src/excel/vuzs/md/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_fe = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[7]
                table_fe.to_excel('src/excel/vuzs/fe/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelFe.append('src/excel/vuzs/fe/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_inf = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[8]
                table_inf.to_excel('src/excel/vuzs/inf/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx',
                                   encoding='cp1251')
                # listFileExcelInf.append('src/excel/vuzs/inf/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_kadr = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[9]
                table_kadr.to_excel('src/excel/vuzs/kadr/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx',
                                    encoding='cp1251')
                # listFileExcelKadr.append('src/excel/vuzs/kadr/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_forRegion = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[10]
                table_forRegion.to_excel('src/excel/vuzs/forRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx',
                                         encoding='cp1251')
                # listFileExcelForRegion.append('src/excel/vuzs/forRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_allInRegion = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[11]
                table_allInRegion.to_excel('src/excel/vuzs/allInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx',
                                           encoding='cp1251')
                # listFileExcelAllInRegion.append('src/excel/vuzs/allInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_studentInVuz = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[12]
                table_studentInVuz.to_excel(
                    'src/excel/vuzs/studentInVuz/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelStudentInVuz.append('src/excel/vuzs/studentInVuz/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_procentStudent = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[13]
                table_procentStudent.to_excel(
                    'src/excel/vuzs/procentStudent/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelProcentStudent.append('src/excel/vuzs/procentStudent/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_dolyaInRegion = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[14]
                table_dolyaInRegion.to_excel(
                    'src/excel/vuzs/dolyaInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelDolyaInRegion.append('src/excel/vuzs/dolyaInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_poPerechnyam = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[15]
                table_poPerechnyam.to_excel(
                    'src/excel/vuzs/poPerechnyam/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
                # listFileExcelPoPerechnyam.append('src/excel/vuzs/poPerechnyam/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                table_dop = \
                    pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
                                 thousands='.')[16]
                table_dop.to_excel('src/excel/vuzs/dop/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx',
                                   encoding='cp1251')
                # listFileExcelDop.append('src/excel/vuzs/dop/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

                pbar.update(1)
            except Exception as e:
                logging.error(str(datetime.now()) + " " + str(e) + str(" : ") + str(url))
                pbar.update(1)


parse_toExcel()
