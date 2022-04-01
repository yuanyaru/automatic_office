# author:yyr
# createDate:2021-03-26

import xlrd
import xlsxwriter
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# 1.读取
data = xlrd.open_workbook("info.xlsx")
classinfo = []
for sheet in data.sheets():
    dict = {'name': sheet.name, 'avgsalary': 0}  # 班级信息
    sum = 0  # 存储薪资
    for i in range(sheet.nrows):
        if i > 1:
            sum += float(sheet.cell(i, 5).value)  # 得到薪资数据
    dict['avgsalary'] = sum/(sheet.nrows-2)
    classinfo.append(dict)
print(classinfo)

# 2.写入excel
workbook = xlsxwriter.Workbook("newinfo.xlsx")
sheet = workbook.add_worksheet()  # 创建工作表
# 写入班级数据
nameinfo =[]
salaryinfo =[]
for item in classinfo:
    nameinfo.append(item['name'])
    salaryinfo.append(item['avgsalary'])
sheet.write_column('A1', nameinfo)
sheet.write_column('B1', salaryinfo)
# 3.写入图表
chart = workbook.add_chart({'type': 'column'})
# 标题
chart.set_title({'name': '平均就业薪资'})
# 数据源
chart.add_series({
    'name': '班级',
    'categories': '=Sheet1!$A$1:$A$3',
    'values': '=Sheet1!$B$1:$B$3'
})
sheet.insert_chart('A7', chart)
workbook.close()

# 4.发送邮件
host_server = 'smtp.qq.com'  # 主机地址
# 发件人邮箱
sender = "398916756@qq.com"
# 发件人邮箱授权码
code = "htpuwwkbbklmbiea"
# 收件人邮箱
user1 = "15664889313@163.com"
# 准备邮件数据
# 邮件标题
mail_title = "!!!3月份平均就业薪资"
# 邮件内容
mail_content = "3月份平均就业薪资，请具体查看附件"
# 构建附件
attachment = MIMEApplication(open('newinfo.xlsx', 'rb').read())
attachment.add_header('Content-Disposition', 'attachment', filename='data.xlsx')

# SMTP
smtp = smtplib.SMTP(host_server)
# 登录
smtp.login(sender, code)
# 发送
msg = MIMEMultipart()  # 带附件的实例
msg['Subject'] = mail_title
msg['From'] = sender
msg['To'] = user1
msg.attach(MIMEText(mail_content))
msg.attach(attachment)
smtp.sendmail(sender, user1, msg.as_string())
