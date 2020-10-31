import json

import pandas as pd
url = 'inst.php?id=326'
# table_forRegion = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[10]
# table_forRegion.to_excel('src/excel/vuzs/forRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelForRegion.append('src/excel/vuzs/forRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_allInRegion = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[11]
# table_allInRegion.to_excel('src/excel/vuzs/allInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelAllInRegion.append('src/excel/vuzs/allInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_studentInVuz = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[12]
# table_studentInVuz.to_excel('src/excel/vuzs/studentInVuz/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelStudentInVuz.append('src/excel/vuzs/studentInVuz/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_procentStudent = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[13]
# table_procentStudent.to_excel('src/excel/vuzs/procentStudent/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelProcentStudent.append('src/excel/vuzs/procentStudent/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_dolyaInRegion = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[14]
# table_dolyaInRegion.to_excel('src/excel/vuzs/dolyaInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelDolyaInRegion.append('src/excel/vuzs/dolyaInRegion/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_poPerechnyam = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[15]
# table_poPerechnyam.to_excel('src/excel/vuzs/poPerechnyam/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelPoPerechnyam.append('src/excel/vuzs/poPerechnyam/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')
#
# table_dop = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',', thousands='.')[16]
# table_dop.to_excel('src/excel/vuzs/dop/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')
# listFileExcelDop.append('src/excel/vuzs/dop/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx')

# table_e = \
#     pd.read_html('http://indicators.miccedu.ru/monitoring/2019/_vpo/' + url, decimal=',',
#                  thousands='.')[2]
# table_e.to_excel('src/vuz' + '_' + str(url.split('id=')[1]) + '.xlsx', encoding='cp1251')

with open('src/vuz_merge.json', encoding="utf-8") as json_file:
    dataVuzs = json.load(json_file)
