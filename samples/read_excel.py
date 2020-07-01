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
print(sheet.merged_cells)    #返回一个列表   起始行、结束行、起始列、结束列

def get_merged_cell_value(row_index,col_index):
    cell_value =None
    for (srow,erow,scol,ecol) in merged:
        if (srow<=row_index  and erow>row_index):
            if (scol <=col_index  and ecol >col_index):
                cell_value =sheet.cell_value(srow,erow)
    return cell_value
print(get_merged_cell_value(2,0))





