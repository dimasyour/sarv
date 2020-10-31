URL_MAIN = 'http://indicators.miccedu.ru/monitoring/2019/'
URL_FO = 'http://indicators.miccedu.ru/monitoring/2019/_vpo/'
RRENAME_VUZ = ['Автономная некоммерческая организация высшего образования ',
               'Государственное бюджетное образовательное учреждение высшего образования ',
               'Федеральное государственное бюджетное образовательное учреждение высшего образования ',
               'федеральное государственное бюджетное образовательное учреждение высшего образования',
               'федеральное государственное автономное образовательное учреждение высшего образования',
               'Образовательная автономная некоммерческая организация высшего образования ',
               'Религиозная организация -духовная образовательная организация высшего образования ',
               'Учреждение высшего образования ',
               ' Национальный исследовательский ',
               ]

digitsList = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
# cellAlph = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
cellAlph = ['C']
# def parse_fo(url, teg):
#     r = requests.get(url, headers=HEADERS)
#     r.encoding = 'cp1251'
#     html = r
#     if html.status_code == 200:
#         soup = BeautifulSoup(html.text, 'html.parser')
#         table = soup.find("table", class_=teg)
#         table_rows = table.find_all('tr')
#         listVuz = []
#         listVuz1 = []
#         dictVuz = {}
#
#         for tr in table_rows:
#             td_a = tr.find_all('a')
#             td_1 = tr.find_all(class_='dkont')
#             for i in range(len(td_a)):
#                 for o in range(len(td_1)):
#                     listVuz1.append((str(td_1[o].get_text(strip=True))))
#                 # print(npListVuz1)
#                 listVuz.append(str(td_a[i].get_text(strip=True)))
#                 listVuz.append(str(td_a[i].get('href')))
#             data = np.array(listVuz1)
#             arr = data.reshape(-1, 9)
#             npListVuz1 = np.array(arr)
#             for y in range(len(npListVuz1)):
#                 listVuz1.append(npListVuz1[y])
#         print(listVuz1)
#     else:
#         print('Error parse_main()')
#
#
# def toJSON(filename, var):
#     with open('src/' + filename, 'w', encoding="utf-8") as fp:
#         json.dump(var, fp, ensure_ascii=False)

# def formatCell():
#     listFileExcel_2 = parse_toExcel()
#     print(Fore.GREEN + 'Форматирование Excel...')
#     with tqdm(total=len(listFileExcel_2)) as pbar2:
#         for u in range(len(listFileExcel_2)):
#             wb1 = load_workbook(listFileExcel_2[u])
#             sheet_ranges = wb1['Sheet1']
#             column_c = sheet_ranges['C']
#             column_d = sheet_ranges['D']
#             column_e = sheet_ranges['E']
#             column_f = sheet_ranges['F']
#             column_g = sheet_ranges['G']
#             column_h = sheet_ranges['H']
#             column_i = sheet_ranges['I']
#             column_j = sheet_ranges['J']
#             column_k = sheet_ranges['K']
#             column_l = sheet_ranges['L']
#             for i in range(len(column_c)):
#                 if column_c[i].value == '' and column_c[i + 1].value == '':
#                     pbar2.update(1)
#                 elif column_c[i].value is not None:
#                     s = column_c[i].value
#                     column_c[i].value = s.split('г.')[0]
#             for y in range(len(column_c)):
#                 if column_c[y].value == '' and column_c[y + 1].value == '':
#                     pbar2.update(1)
#                 elif column_c[y].value is not None:
#                     s = column_d[y].value
#                     s1 = s.replace(" ", "")
#                     if s != '' and (s1[0] in digitsList):
#                         intS = int(s) / 100
#                         column_d[y].value = intS
#                     else:
#                         pass
#             pbar2.update(1)
#             wb1.save(listFileExcel_2[u])
#             # rd = sheet_ranges.row_dimensions[3]  # get dimension for row 3
#             # rd.height = 50  # value in points, there is no "auto"