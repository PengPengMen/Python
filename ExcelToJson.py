# -*- coding: utf-8 -*-
# 这段代码主要的功能是把excel表格转换成utf-8格式的json文件
import os
import sys
import codecs
import xlrd
import xdrlib
import json


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        print("Excel表打开成功")
        return data
    except Exception as e:
        print(u'excel 表格读取失败:%s' % e)
        return None


# 根据索引获取Excel表格中的数据,参数:file：Excel文件路径,colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file,colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    records = []
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             record = {}
             for i in range(len(colnames)):
                record[colnames[i]] = row[i]
             records.append(record)
    print("数据读取完成")
    return records


def save_file(file_path, file_name, data):
    output = codecs.open(file_path + "/" + file_name + ".json", "w", "utf-8")
    output.write(data)
    output.close()


if __name__ == '__main__':
    file_path = 'E:\知识库\平台运维\软件安装清单.xlsx'
    out_flie  = 'D:\测试Json'
    file_name = 'test'
    # open_excel(file_path)
    recodes = excel_table_byindex(file_path)
    encodedjson = json.dumps(recodes, ensure_ascii=False, indent=2)
    # encodedjson = json.dumps(recodes)
    print(encodedjson)
    save_file(out_flie, file_name, encodedjson)
    # output = open('data11.json', 'w+')
    # output.write(encodedjson)
