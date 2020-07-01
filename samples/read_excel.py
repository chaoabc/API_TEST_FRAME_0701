#!/usr/bin/env python
# encoding: utf-8
# @author: miaoxiaochao
# @file: read_excel.py
# @time: 2020/7/1 9:57 下午
#@desc:

import os
import xlrd

excel_path =os.path.join(os.path.dirname(__file__),'data/test_data.xlsx')
print(excel_path)

wb =xlrd.open_workbook(excel_path)
sheet =wb.sheet_by_name('Sheet1')
cell_value =sheet.cell_value(1,0)  # 合并的单元格，值在单元格的0,0  1,0 2,0 3,0……值为空
print(cell_value)

#思路：凡是在merged_cells属性范围内的单元格，它的值均等于左上角首个单元格的值
row_index = 2;col_index =0
merged =sheet.merged_cells    #返回一个列表   起始行、结束行、起始列、结束列 [(1, 5, 0, 1)]
print(sheet.merged_cells)
def get_merged_cell_value(row_index,col_index):
    cell_value =None
    for (srow,erow,scol,ecol) in merged:   #遍历表格中所有合并单元格的位置信息
        if (srow<=row_index  and erow>row_index):      # 1<=2<5
            if (scol <=col_index  and ecol >col_index):       # 0<=0<1
                #如果满足条件，则把合并单元格首个单元格的值赋值给其他合并单元格
                cell_value =sheet.cell_value(srow,scol)
    return cell_value
print(get_merged_cell_value(4,0))





