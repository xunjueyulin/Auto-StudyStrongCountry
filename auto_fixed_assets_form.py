#!/usr/bin/env python 
# -*- coding:utf-8 -*-
import pandas as pd
import xlrd,xlwt
from xlutils.copy import copy
from collections import defaultdict

# 读取转固表
# original_form_path = r"E:\auto-excel\SZ-BB-XJ-JRW-201511-002.xls"
# read_original_form = xlrd.open_workbook(original_form_path,formatting_info=True)
# original_form = copy(read_original_form) # 复制excel文件
# original_form_sheet1 = original_form.get_sheet(0)
project_codes = []
value_list = []
project_data = {}


# 读取数据表格，整理成字典,一键多值


def read_data_to_list(path, usecol, col):    # 读取一列作为列表,list的作用在于传入可变实参
    data = pd.read_excel(path, sheet_name="Sheet1", header=0, usecols=[int(usecol)])
    data_to_list = data.values.tolist()
    for data in data_to_list:
        col.append(data[0])
    # print(list)


def read_datas_to_list(path, usecols, col):   # 读取多列作为列表，list作用在于传入可变实参
    data = pd.read_excel(path, sheet_name="Sheet1", header=0, usecols=usecols)
    data_to_lists = data.values.tolist()
    for data in data_to_lists:
        col.append(data)
    # print(list)


# 循环体，每个工程编码对应将字典的值写入到复制的excel中，打印excel，保存并重命名
def execute_excel():
    read_fixed_assets_form = xlrd.open_workbook(path, formatting_info=True)   # 读取转固表，带格式
    fixed_assets_form = copy(read_fixed_assets_form)  # 复制转固表
    fixed_assets_form_sheet1 = fixed_assets_form.get_sheet(0)
    fixed_assets_form_sheet1.write(3, 3, "建设地址")
    fixed_assets_form_sheet1.write(2, 16, "填表日期")
    fixed_assets_form_sheet1.write(3, 16, "转固日期")
    fixed_assets_form_sheet1.write


if __name__ == '__main__':
    read_data_to_list("data.xls",0 ,project_codes)
    read_datas_to_list("data.xls", [1,2,3,4,5,6,7,8,9,10,11,12] ,value_list)
    # data_to_dic(project_data,project_codes,value_list)
    test = dict(zip(project_codes,value_list)) # 将两个列表组成字典





