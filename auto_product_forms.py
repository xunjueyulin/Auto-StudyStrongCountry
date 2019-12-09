#!/usr/bin/env python
# -*- coding:utf-8 -*-
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy

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
def execute_excel_2016_and_before(original_form_path, data):
    read_fixed_assets_form = xlrd.open_workbook(original_form_path, formatting_info=True)   # 读取转固表，带格式
    fixed_assets_form = copy(read_fixed_assets_form)  # 复制转固表
    fixed_assets_form_sheet1 = fixed_assets_form.get_sheet(0)   # 获取操作sheet
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.right = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.top = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.bottom = xlwt.Borders.MEDIUM  # 添加边框-虚线边框
    borders.left_colour = 0x3A  # 边框上色
    borders.right_colour = 0x3A
    borders.top_colour = 0x3A
    borders.bottom_colour = 0x3A
    for key in data:
        dateFormat = xlwt.XFStyle()  # 确定日期格式
        dateFormat.num_format_str='yyyy/mm/dd'
        style = xlwt.XFStyle()  # 边框样式
        style.borders = borders
        # 1.以下填入必填数据
        fixed_assets_form_sheet1.write(1, 11, key)  # 写入项目编号
        fixed_assets_form_sheet1.write(5, 5, data[key][0])  # 写入项目名称
        fixed_assets_form_sheet1.write(2, 2, data[key][1])  # 写入项目建设地址
        fixed_assets_form_sheet1.write(1, 15, data[key][2], dateFormat)  # 写入填表时间
        fixed_assets_form_sheet1.write(2, 15, data[key][3], dateFormat)  # 写入转固日期
        fixed_assets_form_sheet1.write(5, 9, data[key][5])  # 写入人工费
        fixed_assets_form_sheet1.write(6, 9, data[key][6])  # 写入器材费
        fixed_assets_form_sheet1.write(7, 9, data[key][7])  # 写入待摊投资
        fixed_assets_form_sheet1.write(8, 9, data[key][8])  # 写入设计费
        fixed_assets_form_sheet1.write(9, 9, data[key][9])  # 写入监理费
        fixed_assets_form_sheet1.write(10, 9, data[key][10])  # 写入其他费
        # 2.以下写入资产类别
        fixed_assets_form_sheet1.write(5, 2, data[key][4])   # 人工费资产类别
        fixed_assets_form_sheet1.write(6, 2, data[key][4])   # 器材费资产类别
        fixed_assets_form_sheet1.write(7, 2, data[key][4])   # 待摊投资资产类别
        fixed_assets_form_sheet1.write(8, 2, data[key][4])   # 设计费资产类别
        fixed_assets_form_sheet1.write(9, 2, data[key][4])   # 监理费资产类别
        fixed_assets_form_sheet1.write(10, 2, data[key][4])   # 其他费资产类别
        # 3.以下填充带公式单元格
        # 3.1以下类别名称公式
        fixed_assets_form_sheet1.write(5, 3, xlwt.Formula('IF(C6<>"",VLOOKUP(C6,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 人工费类别名称
        fixed_assets_form_sheet1.write(6, 3, xlwt.Formula('IF(C7<>"",VLOOKUP(C7,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 器材费类别名称
        fixed_assets_form_sheet1.write(7, 3, xlwt.Formula('IF(C8<>"",VLOOKUP(C8,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 待摊类别名称
        fixed_assets_form_sheet1.write(8, 3, xlwt.Formula('IF(C9<>"",VLOOKUP(C9,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 设计费类别名称
        fixed_assets_form_sheet1.write(9, 3, xlwt.Formula('IF(C10<>"",VLOOKUP(C10,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 监理费类别名称
        fixed_assets_form_sheet1.write(10, 3, xlwt.Formula('IF(C11<>"",VLOOKUP(C11,固定资产分类表!$A$3:$H$42,8,),"")'))
        # 其他费类别名称
        # 3.2以下不含税公式
        fixed_assets_form_sheet1.write(5, 13, xlwt.Formula('ROUND(J6/(1+M6),2)'))  # 人工费不含税公式
        fixed_assets_form_sheet1.write(6, 13, xlwt.Formula('ROUND(J7/(1+M7),2)'))  # 器材费不含税公式
        fixed_assets_form_sheet1.write(7, 13, xlwt.Formula('ROUND(J8/(1+M8),2)'))  # 待摊投资不含税公式
        fixed_assets_form_sheet1.write(8, 13, xlwt.Formula('ROUND(J9/(1+M9),2)'))  # 设计费不含税公式
        fixed_assets_form_sheet1.write(9, 13, xlwt.Formula('ROUND(J10/(1+M10),2)'))  # 监理费不含税公式
        fixed_assets_form_sheet1.write(10, 13, xlwt.Formula('ROUND(J11/(1+M11),2)'))  # 其他费不含税公式
        fixed_assets_form_sheet1.write(13, 13, xlwt.Formula('SUM(N6:N13)'))  # 不含税合计公式
        # 3.3以下共提月份公式
        fixed_assets_form_sheet1.write(5, 26, xlwt.Formula('IF(C6<>"",VLOOKUP(C6,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 人工费共提月份公式
        fixed_assets_form_sheet1.write(6, 26, xlwt.Formula('IF(C7<>"",VLOOKUP(C7,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 器材费共提月份公式
        fixed_assets_form_sheet1.write(7, 26, xlwt.Formula('IF(C8<>"",VLOOKUP(C8,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 待摊投资共提月份公式
        fixed_assets_form_sheet1.write(8, 26, xlwt.Formula('IF(C9<>"",VLOOKUP(C9,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 设计费共提月份公式
        fixed_assets_form_sheet1.write(9, 26, xlwt.Formula('IF(C10<>"",VLOOKUP(C10,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 监理费共提月份公式
        fixed_assets_form_sheet1.write(10, 26, xlwt.Formula('IF(C11<>"",VLOOKUP(C11,固定资产分类表!$A$3:$H$42,7,),"")'))
        # 其他费共提月份公式
        # 3.4以下补提月份公式
        fixed_assets_form_sheet1.write(5, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 人工费补提月份公式
        fixed_assets_form_sheet1.write(6, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 器材费补提月份公式
        fixed_assets_form_sheet1.write(7, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 待摊投资补提月份公式
        fixed_assets_form_sheet1.write(8, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 设计费补提月份公式
        fixed_assets_form_sheet1.write(9, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 监理费补提月份公式
        fixed_assets_form_sheet1.write(10, 25, xlwt.Formula('(YEAR($P$2)*12+MONTH($P$2))-(YEAR($P$3)*12+MONTH($P$3))'))
        # 其他费补提月份公式
        # 3.5以下补提折旧金额公式
        fixed_assets_form_sheet1.write(5, 24, xlwt.Formula('ROUND(N6*0.95/AA6*Z6,2)'))  # 人工费补提折旧金额公式
        fixed_assets_form_sheet1.write(6, 24, xlwt.Formula('ROUND(N7*0.95/AA7*Z7,2)'))  # 器材费补提折旧金额公式
        fixed_assets_form_sheet1.write(7, 24, xlwt.Formula('ROUND(N8*0.95/AA8*Z8,2)'))  # 待摊补提折旧金额公式
        fixed_assets_form_sheet1.write(8, 24, xlwt.Formula('ROUND(N9*0.95/AA9*Z9,2)'))  # 设计费补提折旧金额公式
        fixed_assets_form_sheet1.write(9, 24, xlwt.Formula('ROUND(N10*0.95/AA10*Z10,2)'))  # 监理费补提折旧金额公式
        fixed_assets_form_sheet1.write(10, 24, xlwt.Formula('ROUND(N11*0.95/AA11*Z11,2)'))  # 其他费补提折旧金额公式
        fixed_assets_form_sheet1.write(13, 24, xlwt.Formula('SUM(Y6:Y13)'))  # 补提折旧合计公式
        # 保存文件
        fixed_assets_form.save(str(key)+'.xls')




if __name__ == '__main__':
    read_data_to_list("data.xls", 0, project_codes)
    read_datas_to_list("data.xls", [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], value_list)
    # data_to_dic(project_data,project_codes,value_list)
    datas_dic = dict(zip(project_codes, value_list)) # 将两个列表组成字典
    execute_excel_2016_and_before('SZ-BB-XJ-JRW-201511-002（2016年以前项目）.xls', datas_dic)







