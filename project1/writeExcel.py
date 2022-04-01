# author:yyr
# createDate:2021-03-24

# pip install xlwt==1.2.0
import xlwt

# 字体字号
titlestyle = xlwt.XFStyle()  # 初始化样式
titlefont = xlwt.Font()
titlefont.name = "宋体"
titlefont.bold = True  # 加粗
titlefont.height = 11*20  # 字号
titlefont.colour_index = 0x08  # 字体颜色
titlestyle.font = titlefont

# 单元格对齐方式
cellalign = xlwt.Alignment()
cellalign.horz = 0x02
cellalign.vert = 0x01
titlestyle.alignment = cellalign

# 单元格边框设置
borders = xlwt.Borders()
borders.right = xlwt.Borders.DASHED
borders.bottom = xlwt.Borders.DOTTED
titlestyle.borders = borders

# 背景颜色
datastyle = xlwt.XFStyle()
bgcolor = xlwt.Pattern()
bgcolor.pattern = xlwt.Pattern.SOLID_PATTERN
bgcolor.pattern_fore_colour = 22  # 背景颜色
datastyle.pattern = bgcolor

# 第一步：创建工作簿
wb = xlwt.Workbook()
# 第二步：创建工作表
ws = wb.add_sheet("SOE")
# 第三步：填充数据
ws.write_merge(0, 1, 0, 3, "SVG事件表", titlestyle)
# 写入事件数据
data = (("ID",	"name",	"describe",	"level"),
        (1,	"装置上电",	"station_1_SOE_1",	1),
        (2,	"备用",	"station_1_SOE_2",	1))
for i, item in enumerate(data):
    for j, val in enumerate(item):
        if j == 0:
            ws.write(i + 2, j, val, datastyle)
        else:
            ws.write(i+2, j, val)

# 创建第二个工作表
wsimage = wb.add_sheet("image")
# 写入图片
wsimage.insert_bitmap("soe.bmp", 0, 0)
# 第四步：保存
wb.save("SVG-SOE.xlsx")
