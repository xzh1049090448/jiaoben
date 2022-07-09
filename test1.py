'''
xml文件是 product.xml
'''
from xml.etree.ElementTree import parse
import openpyxl

doc = parse('info.xml')  # 开始分析xml文件
infoDic = dict()  # 存储key-value，key为InfoId

fo1 = open("101001.txt", "r")
lines2 = [l.split() for l in fo1.readlines() if l.strip()]


class InfoStruct:
    def __init__(self, setname, infoid, infoname, path):
        self.setName = setname
        self.infoId = infoid
        self.infoName = infoname
        self.path = path


for item in doc.iterfind('SetInfo'):
    setName = item.get("name")
    for it in item.iterfind('InfoItem'):
        infoId = it.get('infoId')
        name = it.findtext('name')
        path1 = it.findtext('path1')
        path2 = it.findtext('path2')
        if path1 is not None and path2 is not None:
            path = path1 + path2
        else:
            path = ''
        infoDic[infoId] = InfoStruct(setName, infoId, name, path)

for i in lines2:
    print("-------")
    print('setName', '=', infoDic[i[0]].setName)
    print('infoId', '=', infoDic[i[0]].infoId)
    print('name', '=', infoDic[i[0]].infoName)
    print('path = ', infoDic[i[0]].path)

# 处理xlsx表
book = openpyxl.load_workbook('biaoge.xlsx')
sheet = book["Sheet1"]