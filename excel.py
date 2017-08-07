#!/usr/bin/env python
# encoding: UTF-8


import xlrd
import xlwt
import getopt
import sys
import copy
import re
from xlutils.copy import copy

# 定义排序函数
def rank_fn(save_file, sheet_idx,add_num):

    #add_num 表示如果表头有一列，则add_num = 0（即针对sheet2），若表头有两列的话则add_num = 1,即针对sheet
    readbook = xlrd.open_workbook(save_file)

    sheet = readbook.sheet_by_index(sheet_idx)

    # 读取xlsx的第一行的单元格内容
    excel_cols = sheet.ncols
    print excel_cols
    if excel_cols < 2+add_num:
        print "save no change"
        return
    node_list = []
    writebook = copy(readbook)
    wsheet = writebook.get_sheet(sheet_idx)

    pattern = re.compile(r'\d+')
    for i in range(1+add_num, excel_cols):
        p = pattern.findall(sheet.cell_value(0, i))
        print int(p[1])
        node_list.append(int(p[1]))
    print node_list
    for i in range(len(node_list) - 1):
        smallest = node_list[i]
        location = i
        for j in range(i, len(node_list)):
            if node_list[j] < smallest:
                print "node_list[j]",node_list[j]
                smallest = node_list[j]
                location = j
        if i != location:
            node_list[i], node_list[location] = node_list[location], node_list[i]
            print "location:", location
            cols_val_sm = sheet.col_values(location+1+add_num)
            clos_val2_al = sheet.col_values(i+1+add_num)
            print cols_val_sm
            print clos_val2_al
            for k in range(len(cols_val_sm)):
                wsheet.write(k, i+1+add_num, cols_val_sm[k])
            for k in range(len(clos_val2_al)):
                wsheet.write(k, location+1+add_num, clos_val2_al[k])

    # 保存排序后的表格
    writebook.save("demo2.xlsx")



if  __name__=='__main__':

    #获取存取数据的txt文本

    save_file = 'demo2.xlsx'
    add_num_list = [1,0]
    for i in range(2):
        sheet_idx = i
        add_num = add_num_list[i]
        rank_fn(save_file, sheet_idx,add_num)
