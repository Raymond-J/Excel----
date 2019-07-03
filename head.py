#-*-coding:utf-8-*-
import os
import xlrd
import xlwt
import readline
readline.parse_and_bind("control-v: paste")
#调用库用来读取写入excle表格
print("--------------------Version 1.0 测试版--------------------"
      "\n-这是一个软件内核测试版本，主要功能为合并Excel表格数据。"
      "\n-使用时请将需要合并的文件放入同一目录下，合并后的表格将保存到该目录下。"
      "\n-由于是测试版，没有做UI界面，后期将会发现Bug修复，并进一步更新功能。"
      "\n--------------------------------------------------------\n"
      )
while True:
    path = str(input("请输入文件地址，按回车继续："))
    files = os.listdir(path)
    print("文件夹中包含以下文件:\n")
    for file in files:
        print(file)
    message = input("文件夹中共有" + str(len(files)) + "个文件，是否开始合并？（Y/N）：").upper()
    if message == 'Y':
        break
#读取文件列表

newworkbook = xlwt.Workbook(encoding = 'utf-8')
newsheet = newworkbook.add_sheet('New')
#新建工作簿和工作表

while True:
    sheetsnum = int(input("请输入要合并表单位置（1,2,3……）："))-1
    rowsnum = int(input("请输入合并数据的起始行："))-1
    colsnum = ord((input("请输入合并数据的起始列：")).upper())-65
    message = input("是否开始合并数据？（Y/N）").upper()
    if message == "Y":
        break
    #输入需要统计的行和列定位表格
x = rowsnum

for file in files:
    workbook = xlrd.open_workbook(path+'\\'+file)
    sheet = workbook.sheets()[sheetsnum]
    rowscount = sheet.nrows
    colscount = sheet.ncols
    #计算行列数量

    for i in range(rowsnum, rowscount):
        for j in range(colsnum, colscount):
            newsheet.write(x, j, sheet.cell(i, j).value)
        x = x+1
    #复制数据到新表

newworkbook.save(path+"\\Statistics.xls")
print("统计完成，文件保存为Statistics.xls！")
while True: input()