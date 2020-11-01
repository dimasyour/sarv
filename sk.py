import json


def printByCity(city):
    with open('src/by_city.json', encoding="utf-8") as json_file:
        dataCity = json.load(json_file)
    keyDataCity = list(dataCity.keys())
    for i in range(len(keyDataCity)):
        if dataCity[keyDataCity[i]]['region'] == city:
            print(dataCity[keyDataCity[i]])


def printByFo(fo):
    with open('src/by_fo.json', encoding="utf-8") as json_file:
        dataRegion = json.load(json_file)
    keyDataFo = list(dataRegion.keys())
    for i in range(len(keyDataFo)):
        if dataRegion[keyDataFo[i]]['region'] == fo:
            print(dataRegion[keyDataFo[i]])


# слить два JSON, теперь в главном JSON есть значени И региона И ФО
def mergeTwoBy():
    with open('src/by_city.json', encoding="utf-8") as json_file:
        dataCity = json.load(json_file)
    keyDataCity = list(dataCity.keys())
    with open('src/by_fo.json', encoding="utf-8") as json_file:
        dataRegion = json.load(json_file)
    keyDataFo = list(dataRegion.keys())
    for o in range(len(dataCity)):
        for u in range(len(dataRegion)):
            if dataCity[keyDataCity[o]]['vuz_name'] == dataRegion[keyDataFo[u]]['vuz_name']:
                dataCity[keyDataCity[o]].update(eval('{"fo" : "' + dataRegion[keyDataFo[u]]['region'] + '"}'))
    with open('src/json/main.json', 'w',
              encoding="utf-8") as fp:
        json.dump(dataCity, fp, ensure_ascii=False)


# вывод по любому параметру
def printByJson(param, value):
    with open('src/json/main.json', encoding="utf-8") as json_file:
        mainJson = json.load(json_file)
    keyMainJson = list(mainJson.keys())
    for i in range(len(keyMainJson)):
        if mainJson[keyMainJson[i]][param] == value:
            print(mainJson[keyMainJson[i]])

mergeTwoBy()