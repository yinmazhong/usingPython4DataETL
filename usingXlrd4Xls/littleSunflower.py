# -*- coding: utf-8 -*-
import xlrd

data = xlrd.open_workbook('/home/xxh/Documents/aprildata/201801jixiao.xls')
chengwuzhang = data.sheets()[0]
gaojichengwuyuan = data.sheets()[1]
chengwuyuan = data.sheets()[2]
chengwuxueyuan = data.sheets()[3]
anquan = data.sheets()[4]
fuwu = data.sheets()[5]
shengchan = data.sheets()[6]
zonghe = data.sheets()[7]
peixun = data.sheets()[8]
chengwuyuanhanghoukaoping = data.sheets()[9]
chengwuzhangfuwumanyidu = data.sheets()[10]
gerenchuqin = data.sheets()[11]

codes = ['A×××49', 'A2×××6', 'A2×××0', 'A1×××8', 'A×××1']
names = ['史××', '张××', '吴××', '××', '××']


def sortKey(key):
    return key[2]


print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++安全+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')
for aqRowIndex in range(anquan.nrows):
    row = anquan.row(aqRowIndex)
    codeCell = anquan.cell_value(rowx=aqRowIndex, colx=3)
    nameCell = anquan.cell_value(rowx=aqRowIndex, colx=2)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        print(
            "-----------------------------------------------------------------------------------------------------------------------------------------------------")
        print(row)
        print(
            "-----------------------------------------------------------------------------------------------------------------------------------------------------")

print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')
print("\n")
print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++服务+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = fuwu.row(0)
print(fuwu.row(0))
rows = []
for aqRowIndex in range(fuwu.nrows):
    row = fuwu.row(aqRowIndex)
    codeCell = fuwu.cell_value(rowx=aqRowIndex, colx=3)
    nameCell = fuwu.cell_value(rowx=aqRowIndex, colx=2)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)
rows.sort(key=sortKey)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

print("\n")
print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++生产+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = shengchan.row(0)
print(shengchan.row(0))
rows = []
for aqRowIndex in range(shengchan.nrows):
    row = shengchan.row(aqRowIndex)
    codeCell = shengchan.cell_value(rowx=aqRowIndex, colx=3)
    nameCell = shengchan.cell_value(rowx=aqRowIndex, colx=2)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        # print(row)
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)
rows.sort(key=sortKey)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')
print("\n")

print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++综合+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = zonghe.row(0)
print(zonghe.row(0))
rows = []
for aqRowIndex in range(zonghe.nrows):
    row = zonghe.row(aqRowIndex)
    codeCell = zonghe.cell_value(rowx=aqRowIndex, colx=3)
    nameCell = zonghe.cell_value(rowx=aqRowIndex, colx=2)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)

rows.sort(key=sortKey)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')
print("\n")
print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++培训+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = peixun.row(0)
print(peixun.row(0))
rows = []
for aqRowIndex in range(peixun.nrows):
    row = peixun.row(aqRowIndex)
    codeCell = peixun.cell_value(rowx=aqRowIndex, colx=3)
    nameCell = peixun.cell_value(rowx=aqRowIndex, colx=2)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)

rows.sort(key=sortKey)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')

print("\n")

print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++乘务员航后考评+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = chengwuyuanhanghoukaoping.row(0)
print(chengwuyuanhanghoukaoping.row(0))
rows = []
for aqRowIndex in range(chengwuyuanhanghoukaoping.nrows):
    row = chengwuyuanhanghoukaoping.row(aqRowIndex)
    codeCell = chengwuyuanhanghoukaoping.cell_value(rowx=aqRowIndex, colx=1)
    nameCell = chengwuyuanhanghoukaoping.cell_value(rowx=aqRowIndex, colx=0)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)

rows.sort(key=sortKey, reverse=True)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print(
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')
print("\n")

print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++乘务长服务满意度+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = chengwuzhangfuwumanyidu.row(0)
print(chengwuzhangfuwumanyidu.row(0))
rows = []
for aqRowIndex in range(chengwuzhangfuwumanyidu.nrows):
    row = chengwuzhangfuwumanyidu.row(aqRowIndex)
    codeCell = chengwuzhangfuwumanyidu.cell_value(rowx=aqRowIndex, colx=1)
    nameCell = chengwuzhangfuwumanyidu.cell_value(rowx=aqRowIndex, colx=0)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)

rows.sort(key=sortKey)
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")

print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')

print("\n")

print(
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++个人出勤+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=')

header = gerenchuqin.row(0)
print(gerenchuqin.row(0))
rows = []
for aqRowIndex in range(gerenchuqin.nrows):
    row = gerenchuqin.row(aqRowIndex)
    codeCell = gerenchuqin.cell_value(rowx=aqRowIndex, colx=1)
    nameCell = gerenchuqin.cell_value(rowx=aqRowIndex, colx=0)
    codeFlag = codeCell in codes
    nameFlag = nameCell in names
    if (codeFlag | nameFlag):
        res = []
        for index in range(len(header)):
            res.append(header[index].value + ':' + str(row[index].value))
        # print(row[0].value,row[1].value,row[2].value,row[3].value)
        rows.append(res)

rows.sort()
for row in rows:
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
    print(row)
    print(
        "-----------------------------------------------------------------------------------------------------------------------------------------------------")
print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++')


# print table0.row_values(1)
