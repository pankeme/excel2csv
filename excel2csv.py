#!/usr/bin/env python26
# coding=gbk
import sys
import xlrd
import csv
import math
import types
from datetime import date,datetime,time
 
def csv_from_excel(xlsx_filepath, csv_filepath, sheet):
    '''
    excel转换函数
    将excel文件转换为文本文件
    包含旧版的.xls文件和新版的.xlsx文件
    xlsx_filepath: 待转换文件路径
    csv_filepath: 生成的文件路径
    sheet: Excel中工作表索引，第一个工作表的索引为0，以此类推
    '''
    wb = xlrd.open_workbook(xlsx_filepath)
    sh = wb.sheet_by_index(sheet)
    csv_file = open(csv_filepath, 'wb')
    wr = csv.writer(csv_file, quoting=csv.QUOTE_NONE)
    nrows = sh.nrows
    ncols = sh.ncols

    for rownum in xrange(nrows):
        temp = []
        for colnum in xrange(ncols):
            cell = sh.cell(rownum, colnum)
            if cell.ctype is xlrd.XL_CELL_DATE:
                date_value = xlrd.xldate_as_tuple(cell.value, wb.datemode)
                date_list = [str(x) for x in list(date_value)[0:3]]
                if len(date_list[1]) == 1:
                    date_list[1] = '0' + date_list[1]
                if len(date_list[2]) == 1:
                    date_list[2] = '0' + date_list[2]
                temp.append(''.join(date_list))
            elif cell.ctype is xlrd.XL_CELL_NUMBER:
                if math.ceil(cell.value) == math.floor(cell.value) :
                    temp.append(str(int(cell.value)))
                else:
                    temp.append(str(cell.value))
            else:
                temp.append(str(cell.value).encode('gbk'))
        
        #将每行中的数据以Tab分隔
        newrow = '\t'.join(temp)
        wr.writerow([newrow])

    csv_file.close()

if __name__ == "__main__":
    if len(sys.argv) <= 4:
        seet = int(sys.argv[3])
        csv_from_excel(sys.argv[1], sys.argv[2],seet)
    else:
        print '参数不正确'
        sys.exit(1)

