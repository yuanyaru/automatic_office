# author:yyr
# createDate:2021-03-23

# pip install xlrd==1.2.0
import xlrd

data = xlrd.open_workbook("data.xlsx")
# print(data.sheet_loaded(0))
# data.unload_sheet(0)
# print(data.sheet_loaded(0))
# print(data.sheet_loaded(1))

# print(data.sheets())  # 获取全部sheet
# print(data.sheets()[0])
# print(data.sheet_by_index(0))  # 根据索引获取工作表
# print(data.sheet_by_name("yx"))  # 根据 sheetname 获取工作表
# print(data.sheet_names())  # 打印所有工作表名
# print(data.nsheets)  # 返回excel中工作表的数量

# 操作excel行
# sheet = data.sheet_by_index(3)  # 获取第一个工作表
# print(sheet.nrows)  # 获取sheet下的有效行数
# print(sheet.row(0))  # 该行单元格对象组成的列表
# print(sheet.row_types(1))  # 获取单元格的数据类型
# print(sheet.row(0)[2])
# print(sheet.row(0)[2].value)  # 获取单元格value
# print(sheet.row_values(0))  # 获取指定行的value
# print(sheet.row_len(0))  # 获取指定行的单元格长度

# 操作excel列
# sheet = data.sheet_by_index(0)
# print(sheet.ncols)
# print(sheet.col(0))  # 返回该列单元格对象组成的列表
# print(sheet.col(0)[2].value)
# print(sheet.col_values(0))  # 返回该列所有单元格的value组成的列表
# print(sheet.col_types(1))

# 操作excel单元格
sheet = data.sheet_by_index(0)
print(sheet.cell(0, 1))
print(sheet.cell_type(0, 1))
print(sheet.cell(0, 1).ctype)
print(sheet.cell(0, 1).value)
print(sheet.cell_value(0, 1))
