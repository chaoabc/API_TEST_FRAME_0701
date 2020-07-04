#!/usr/bin/env python
# encoding: utf-8
# @author: miaoxiaochao
# @file: read_excel.py
# @time: 2020/7/1 9:57 下午
#@desc:

import os
import xlrd

excel_path =os.path.join(os.path.dirname(__file__),'data/test_data.xlsx')
# 题二
wb =xlrd.open_workbook(excel_path)
sheet =wb.sheet_by_name('Sheet1')
def merged_judge(x,y):
    for (srow,erow,scol,ecol)  in  sheet.merged_cells:
        if (x>=srow and x<erow):
            if (y>=scol and y<ecol):
                return '此单元格是合并的'
        return '普通单元格'
print(merged_judge(2,0))


# 题三：获取整列数据，方式一
# def get_sheet_value():
#     lie =[]
#     for i in range(1,sheet.nrows):
#         sheet_value = sheet.cell_value(i, 3)
#         lie.append(sheet_value)
#         lie1 =sorted(lie)
#     return  lie1
# print(get_sheet_value())
#
#题一： 获取整列数据，方式二
# lie =[str(sheet.cell_value(i,3)) for i in range(1,sheet.nrows)]
# print(lie)

# # 获取整行数据
# hang =[str(sheet.cell_value(0,i))  for i in range(1,sheet.ncols)]
# print(hang)

# cell_value =sheet.cell_value(1,0)  # 合并的单元格，值在单元格的0,0  1,0 2,0 3,0……值为空
# print(cell_value)
# #思路：凡是在merged_cells属性范围内的单元格，它的值均等于左上角首个单元格的值
# merged =sheet.merged_cells    #返回一个列表   起始行、结束行、起始列、结束列 [(1, 5, 0, 1)]
# print(sheet.merged_cells)
# def get_merged_cell_value(row_index,col_index):
#     cell_value =None
#     for (srow,erow,scol,ecol) in merged:   #遍历表格中所有合并单元格的位置信息
#         if (srow<=row_index  and erow>row_index):      # 1<=2<5
#             if (scol <=col_index  and ecol >col_index):       # 0<=0<1
#                 #如果满足条件，则把合并单元格首个单元格的值赋值给其他合并单元格
#                 cell_value =sheet.cell_value(srow,scol)
#     return cell_value
# print(get_merged_cell_value(3,0))











