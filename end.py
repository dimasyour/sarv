import json
import os


def getJsonList(path):
    directory = path
    jsons = os.listdir(directory)
    jsonsList = filter(lambda x: x.endswith('.json'), jsons)
    return jsons


def mergeJsonVuzs(json_path, output):
    listik = getJsonList(json_path)
    fdf = {}
    for i in listik:
        with open(json_path + i, encoding="utf-8") as json_file:
            dataVuzs = json.load(json_file)
            fdf.update(dataVuzs)
    with open(output, 'w', encoding="utf-8") as fp:
        json.dump(fdf, fp, ensure_ascii=False)


mergeJsonVuzs('src/json/vuzs/', 'src/vuz_merge.json')
mergeJsonVuzs('src/json/by_fo/', 'src/by_fo.json')
mergeJsonVuzs('src/json/by_city/', 'src/by_city.json')
