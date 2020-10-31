import json
import os


def getJsonList(path):
    directory = path
    jsons = os.listdir(directory)
    jsonsList = filter(lambda x: x.endswith('.json'), jsons)
    return jsons


def mergeJsonVuzs():
    listik = getJsonList('src/json/vuzs/')
    fdf = {}
    for i in listik:
        with open('src/json/vuzs/'+i, encoding="utf-8") as json_file:
            dataVuzs = json.load(json_file)
            fdf.update(dataVuzs)
    with open('src/vuz_merge.json', 'w', encoding="utf-8") as fp:
        json.dump(fdf, fp, ensure_ascii=False)

mergeJsonVuzs()