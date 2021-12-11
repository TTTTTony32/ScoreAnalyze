import xlrd
import xlwt

# 读取总分数据/初始化列表
book = xlrd.open_workbook('E:\paiming.xls')
sheet = book.sheet_by_index(0)
sourse = []
order = []
for i in range(sheet.nrows):
    sourse.append(sheet.cell(i, 16).value)
for i in range(sheet.nrows):
    order.append(1)
count = 1
i = 0

#计算排名
while i < sheet.nrows - 1:
    if sourse[i] == sourse[i+1]:
        count = count + 1
        order[i] = i+1
        order[i+1] = i+1
        i = i + 1
    else:
        order[i+1] = i+1+count
        count = 1
        i = i + 1
print(len(order))

#写入新表
file = xlwt.Workbook()
nsheet = file.add_sheet('Order')
for i in range(len(order)):
    nsheet.write(i, 0,order[i])
file.save('ordered.xls')