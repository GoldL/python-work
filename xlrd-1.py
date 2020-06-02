# -*- coding: utf-8 -*-
# @Time    : 2020/6/2 下午9:17
# @Author  : iGolden
# @Software: PyCharm

import xlrd

# 打开工作表
data = xlrd.open_workbook("（最终版） 2016春季数计学院教师辅导答疑安排表.xlsx")
# # 是否加载完成
# data.sheet_loaded(0)
# # 是否卸载完成
# data.unload_sheet(0)
# 获取所有的sheet
data.sheets()
# 根据索引获取工作表
data.sheet_by_index(0)
# 根据sheet名称索引
data.sheet_by_name('Sheet2')
# 获取所有excel工作表的name
data.sheet_names()
sheet = data.sheet_by_index(0)

# 操作行
# 打印有效行数
sheet.nrows
# 获取该行对象组成的列表
print(sheet.row(0))
# 获取类型
print(sheet.row_types(1))
# 获取单元格的value
print(sheet.row(2)[0].value)
# 获取该行所有单元格的value
print(sheet.row_values(2))
# 得到单元格的长度
print(sheet.row_len(2))

# 操作列
# 获取有效列数
print(sheet.ncols)
# 获取该列对象组成的列表
print(sheet.col(1))
# 获取该列所有单元格的value
print(sheet.col_values(0))

# 操作单元格
print(sheet.cell(1, 4))

