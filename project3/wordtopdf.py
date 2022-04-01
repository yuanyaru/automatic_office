# author:yyr
# createDate:2021-03-29

# pip install pywin32
from win32com.client import constants, gencache
import os


def createpdf(wordPath, pdfPath):
    word = gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    # 转换方法
    doc.ExportAsFixedFormat(pdfPath, constants.wdExportFormatPDF)
    # word.Quit()


# 单个文件的转换
# createpdf('D:/git-workspace/automatic_office/project3/info.docx', 'D:/git-workspace/automatic_office/project3/info.pdf')

# 多个文件的转换
print(os.listdir('.'))  # 当前文件夹下的所有文件
wordfiles=[]
for file in os.listdir('.'):
    if file.endswith(('.doc', '.docx')):
        wordfiles.append(file)
print(wordfiles)

for file in wordfiles:
    filepath = os.path.abspath(file)
    index = filepath.rindex('.')
    pdfpath = filepath[:index] + '.pdf'
    # print(pdfpath)
    createpdf(filepath, pdfpath)
