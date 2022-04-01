# author:yyr
# createDate:2021-03-25

# excel导入试题到数据库
import xlrd

data = xlrd.open_workbook("data2.xlsx")
sheet = data.sheet_by_index(0)  # 获取工作表
questionList = []  # 构建试题列表


# 试题类
class Question:
    pass


for i in range(sheet.nrows):
    if i > 0:
        obj = Question()  # 构建试题对象
        obj.subject = sheet.cell(i, 1).value  # 题目
        obj.answer = sheet.cell(i, 2).value  # 答案
        obj.optionA = sheet.cell(i, 3).value  # 选项a
        obj.optionB = sheet.cell(i, 4).value  # 选项b
        obj.optionC = sheet.cell(i, 5).value  # 选项c
        obj.optionD = sheet.cell(i, 6).value  # 选项d
        questionList.append(obj)

print(questionList)

# 导入操作 pip install pymysql
from mysqlhelper import *
# 链接到数据库
db = dbhelper("127.0.0.1", 3306, "root", "root", "test")
# 插入语句
sql = "insert into question(subject, answer, optionA, optionB, optionC, optionD) VALUES(%s, %s, %s, %s, %s, %s)"
val = []   # 空列表来存储元组
for item in questionList:
    val.append((item.subject, item.answer, item.optionA, item.optionB, item.optionC, item.optionD))

print(val)
db.executemanydata(sql, val)
