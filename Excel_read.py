#!/usr/bin/env python
# _*_ UTF-8 _*_
'''
@project:GBDT+LR-Demo
@author:xiaofei
'''

import xlrd

# 用来存储数据
tables = []
newTables = []
def read_excel(path):
    # 打开文件
    workbook = xlrd.open_workbook(path)
    # 获取所有sheet
    sheet_name = workbook.sheet_names()[0]
    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0) # sheet索引从0开始
    # sheet = workbook.sheet_by_name('Sheet1')
    #print (workboot.sheets()[0])
    # sheet的名称，行数，列数
    print (sheet.name,sheet.nrows,sheet.ncols)
    # 获取整行和整列的值（数组）
    rows = sheet.row_values(1) # 获取第2行内容
    # cols = sheet.col_values(2) # 获取第3列内容
    print (rows)
    # print (cols)
    # 将内容读取出来，放到list中。
    for rown in range(sheet.nrows):
       array = {'L1':'','L2':'','L3':'','L4':'','Question':'','Answer':''}
       array['L1'] = sheet.cell_value(rown,0)
       array['L2'] = sheet.cell_value(rown,1)
       array['L3'] = sheet.cell_value(rown,2)
       array['L4'] = sheet.cell_value(rown,3)
       array['Question'] = sheet.cell_value(rown,4)
       array['Answer'] = sheet.cell_value(rown,5)
       tables.append(array)
    print(tables)
    print (len(tables))
if __name__ == '__main__':
    # 读取Excel
    read_excel(r'C:/Users/xiaofei/Desktop/test3.xlsx');
    print ('读取成功')



