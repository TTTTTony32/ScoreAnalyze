import random
from collections import defaultdict

import xlrd
import xlwt

# 读取全科数据
book = xlrd.open_workbook('E:\chengji.xls')
yuwen = book.sheet_by_index(0)
shuxue = book.sheet_by_index(1)
yingyu = book.sheet_by_index(2)
wuli = book.sheet_by_index(3)
huaxue = book.sheet_by_index(4)
shengwu = book.sheet_by_index(5)
zhengzhi = book.sheet_by_index(6)
dili = book.sheet_by_index(7)

# 初始化字典
source = defaultdict(list)
for i in range(yuwen.nrows):
    source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 0).value)
    source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 2).value)
    source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 3).value)
    for j in range(12):
        source[yuwen.cell(i, 1).value].append(0.0)

# 语文数据读取验证
for i in range(yuwen.nrows):
    if yuwen.cell(i, 1).value in source:
        source[yuwen.cell(i, 1).value][3] = yuwen.cell(i, 4).value
    else:
        print('?')
        source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 0).value)
        source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 2).value)
        source[yuwen.cell(i, 1).value].append(yuwen.cell(i, 3).value)
        source[yuwen.cell(i, 1).value][3] = yuwen.cell(i, 4).value

# 数学数据读取验证
for i in range(shuxue.nrows):
    if shuxue.cell(i, 1).value in source:
        source[shuxue.cell(i, 1).value][4] = shuxue.cell(i, 4).value
    else:
        print('Shuxue Not Found')
        source[shuxue.cell(i, 1).value].append(shuxue.cell(i, 0).value)
        source[shuxue.cell(i, 1).value].append(shuxue.cell(i, 2).value)
        source[shuxue.cell(i, 1).value].append(shuxue.cell(i, 3).value)
        for j in range(12):
            source[shuxue.cell(i, 1).value].append(0.0)
        source[shuxue.cell(i, 1).value][4] = shuxue.cell(i, 4).value
        print(source[shuxue.cell(i, 1).value])

# 英语数据读取验证
for i in range(yingyu.nrows):
    if yingyu.cell(i, 1).value in source:
        source[yingyu.cell(i, 1).value][5] = yingyu.cell(i, 4).value
    else:
        print('Yingyu Not Found')
        source[yingyu.cell(i, 1).value].append(yingyu.cell(i, 0).value)
        source[yingyu.cell(i, 1).value].append(yingyu.cell(i, 2).value)
        source[yingyu.cell(i, 1).value].append(yingyu.cell(i, 3).value)
        for j in range(12):
            source[yingyu.cell(i, 1).value].append(0.0)
        source[yingyu.cell(i, 1).value][5] = yingyu.cell(i, 4).value
        print(source[yingyu.cell(i, 1).value])

# 物理数据读取验证
for i in range(wuli.nrows):
    if wuli.cell(i, 1).value in source:
        source[wuli.cell(i, 1).value][6] = wuli.cell(i, 4).value
    else:
        print('Wuli Not Found')
        source[wuli.cell(i, 1).value].append(wuli.cell(i, 0).value)
        source[wuli.cell(i, 1).value].append(wuli.cell(i, 2).value)
        source[wuli.cell(i, 1).value].append(wuli.cell(i, 3).value)
        for j in range(12):
            source[wuli.cell(i, 1).value].append(0.0)
        source[wuli.cell(i, 1).value][6] = wuli.cell(i, 4).value
        print(source[wuli.cell(i, 1).value])

# 化学数据读取验证
for i in range(huaxue.nrows):
    if huaxue.cell(i, 1).value in source:
        source[huaxue.cell(i, 1).value][7] = huaxue.cell(i, 4).value
        source[huaxue.cell(i, 1).value][8] = huaxue.cell(i, 5).value
    else:
        print('Huaxue Not Found')
        source[huaxue.cell(i, 1).value].append(huaxue.cell(i, 0).value)
        source[huaxue.cell(i, 1).value].append(huaxue.cell(i, 2).value)
        source[huaxue.cell(i, 1).value].append(huaxue.cell(i, 3).value)
        for j in range(12):
            source[huaxue.cell(i, 1).value].append(0.0)
        source[huaxue.cell(i, 1).value][7] = huaxue.cell(i, 4).value
        source[huaxue.cell(i, 1).value][8] = huaxue.cell(i, 5).value
        print(source[huaxue.cell(i, 1).value])

# 生物数据读取验证
for i in range(shengwu.nrows):
    if shengwu.cell(i, 1).value in source:
        source[shengwu.cell(i, 1).value][9] = shengwu.cell(i, 4).value
        source[shengwu.cell(i, 1).value][10] = shengwu.cell(i, 5).value
    else:
        print('Shengwu Not Found')
        source[shengwu.cell(i, 1).value].append(shengwu.cell(i, 0).value)
        source[shengwu.cell(i, 1).value].append(shengwu.cell(i, 2).value)
        source[shengwu.cell(i, 1).value].append(shengwu.cell(i, 3).value)
        for j in range(12):
            source[shengwu.cell(i, 1).value].append(0.0)
        source[shengwu.cell(i, 1).value][9] = shengwu.cell(i, 4).value
        source[shengwu.cell(i, 1).value][10] = shengwu.cell(i, 5).value
        print(source[shengwu.cell(i, 1).value])

# 政治数据读取验证
for i in range(zhengzhi.nrows):
    if zhengzhi.cell(i, 1).value in source:
        source[zhengzhi.cell(i, 1).value][11] = zhengzhi.cell(i, 4).value
        source[zhengzhi.cell(i, 1).value][12] = zhengzhi.cell(i, 5).value
    else:
        print('Zhengzhi Not Found')
        source[zhengzhi.cell(i, 1).value].append(zhengzhi.cell(i, 0).value)
        source[zhengzhi.cell(i, 1).value].append(zhengzhi.cell(i, 2).value)
        source[zhengzhi.cell(i, 1).value].append(zhengzhi.cell(i, 3).value)
        for j in range(12):
            source[zhengzhi.cell(i, 1).value].append(0.0)
        source[zhengzhi.cell(i, 1).value][11] = zhengzhi.cell(i, 4).value
        source[zhengzhi.cell(i, 1).value][12] = zhengzhi.cell(i, 5).value
        print(source[zhengzhi.cell(i, 1).value])

# 地理数据读取验证
for i in range(dili.nrows):
    if dili.cell(i, 1).value in source:
        source[dili.cell(i, 1).value][13] = dili.cell(i, 4).value
        source[dili.cell(i, 1).value][14] = dili.cell(i, 5).value
    else:
        print('dili Not Found')
        source[dili.cell(i, 1).value].append(dili.cell(i, 0).value)
        source[dili.cell(i, 1).value].append(dili.cell(i, 2).value)
        source[dili.cell(i, 1).value].append(dili.cell(i, 3).value)
        for j in range(12):
            source[dili.cell(i, 1).value].append(0.0)
        source[dili.cell(i, 1).value][13] = dili.cell(i, 4).value
        source[dili.cell(i, 1).value][14] = dili.cell(i, 5).value
        print(source[dili.cell(i, 1).value])

# 测试数据
rand = random.randint(0,yuwen.nrows)
print('Test Value：')
print(source[yuwen.cell(rand, 1).value])
print('Student count:')
print(len(source))

#写入文件
file = xlwt.Workbook()
sheet = file.add_sheet('总表')
list_ = []
for item in source.items():
    list_.append(list(item))
for i in range(len(source)):
    sheet.write(i, 0,list_[i][0])
for i in range(len(source)):
    for j in range(15):
        sheet.write(i, j+1,list_[i][1][j])
file.save('out.xls')