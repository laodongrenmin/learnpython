#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import xlrd
from xlrd.formatting import XFAlignment, XF

from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas




A4_size = [297 * mm, 210 * mm]
top_margin = 50
left_margin = 50


class Pos(object):
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __str__(self):
        return "x: {} y:{}".format(self.x, self.y)


class Size(object):
    def __init__(self, w, h):
        self.w = int(w)
        self.h = int(h)

    def __str__(self):
        return "w:{} h:{}".format(self.w, self.h)


a4_size = Size(w=210 * mm, h=297 * mm)


excel_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/excel.xls'
pdf_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/excel.pdf'


rb = xlrd.open_workbook(excel_file_path, formatting_info=True)
sheet_count = rb.nsheets

print("表格数: {} 名字: {}".format(rb.nsheets, [name for name in rb.sheet_names()]))

for i in range(0, rb.nsheets):
    rs = rb.sheet_by_index(i)
    print("sheet: '{}'  列数: {} 行数: {}".format(rs.name,  rs.ncols, rs.nrows))

rs = rb.sheets()[0]

ncols = min(len(rs.colinfo_map), rs.ncols)
nrows = min(len(rs.rowinfo_map), rs.nrows)

ncols = rs.ncols
nrows = rs.nrows

col_widths = [rs.colinfo_map.get(colx).width/2.54 for colx in range(0, ncols)]
row_heights = [rs.rowinfo_map.get(rowx).height for rowx in range(0, nrows)]

need_size = Size(w=sum(col_widths),
                 h=sum(row_heights))


merged_cells = dict()
for row1, row2, col1, col2 in rs.merged_cells:
    key = '{}_{}'.format(row1, col1)
    size = Size(w=sum(col_widths[col1:col2]), h=sum(row_heights[row1:row2]))
    merged_cells[key] = [size, (row1,row2, col1, col2)]




# excel一页搞到PDF一页上
print("A4 {} need {}".format(a4_size, need_size))
c = canvas.Canvas("c:/Users/zbd/PycharmProjects/mylearn/res/hello.pdf", pagesize=(need_size.w + 200, need_size.h + 100))
from reportlab.lib.colors import yellow, red, black,white, HexColor
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

font = TTFont('STCAIYUN', 'STCAIYUN.TTF')
chcfs = TTFont('彩虹粗仿宋', 'chcfs.ttf')
hwhp = TTFont('华文琥珀', 'STHUPO.TTF')
st = TTFont('宋体', 'simsun.ttc')
kt = TTFont('楷体', 'simkai.ttf')
hwst = TTFont('华文宋体', 'STSONG.TTF')

fs = TTFont('仿宋', 'simfang.ttf')

pdfmetrics.registerFont(chcfs)
pdfmetrics.registerFont(hwhp)
pdfmetrics.registerFont(st)
pdfmetrics.registerFont(kt)
pdfmetrics.registerFont(hwst)
pdfmetrics.registerFont(fs)


start_pos = Pos(x=150, y=50)
line_list = list()
pos = Pos(start_pos.x, start_pos.y)
print(need_size)
for rowx in range(nrows-1, -1, -1):
    for colx in range(0, ncols):
        cell_size = Size(w=col_widths[colx], h=row_heights[rowx])
        cell = rs.cell(rowx, colx)
        xf = rs.book.xf_list[cell.xf_index]
        b = xf.border
        bl, bt, br, bb = b.left_line_style, b.top_line_style, b.right_line_style, b.bottom_line_style
        if bl:
            line_list.append((pos.x, pos.y, pos.x, pos.y+cell_size.h))
        if bt:
            line_list.append((pos.x, pos.y + cell_size.h, pos.x + cell_size.w, pos.y + cell_size.h))
        if br:
            line_list.append((pos.x + cell_size.w, pos.y, pos.x + cell_size.w, pos.y + cell_size.h))
        if bb:
            line_list.append((pos.x, pos.y, pos.x + cell_size.w, pos.y))
        if rowx == 0:
            pass
        if cell.value:
            m = merged_cells.get('{}_{}'.format(rowx, colx))
            if m:
                w, h = m[0].w, m[0].h
            else:
                w, h = cell_size.w, cell_size.h

            font = rs.book.font_list[rs.book.xf_list[cell.xf_index].font_index]

            if font.name in pdfmetrics.getRegisteredFontNames():
                # print("value:{} font name:{} height:{}".format("{}".format(cell.value), font.name, font.height))
                c.setFont(font.name, font.height)
            else:
                c.setFont("彩虹粗仿宋", font.height)
                print("value:{} font name:{} 彩虹粗仿宋 height:{}".format("{}".format(cell.value), font.name, font.height))
                font.name = "彩虹粗仿宋"

            h1 = font.height  # 当为一行时候， 当 自动换行且宽度不够时候以及有多行时
            if cell.ctype == 1:  # string
                v = "{}".format(cell.value)
                tmps = v.split("\n")
                if len(tmps) > 1:
                    h1 = len(tmps) * (1 + 0.2) * h1  # 0.2倍行距
            elif cell.ctype == 2:  # float
                if cell.value % 1 == 0.0:
                    tmps = ['{}'.format(int(cell.value))]
                else:
                    tmps = ['{}'.format(cell.value)]
                pass
            else:
                tmps = ["{}".format(cell.value)]

            xfa: XFAlignment = xf.alignment
            print("row:{} col:{} ctype:{} value:{} hor_align:{} vert_align:{}".format(rowx, colx, cell.ctype,"{}".format(cell.value), xfa.hor_align, xfa.vert_align))



            if xfa.vert_align == 0:  # v_top
                off_h = 0
            elif xfa.vert_align == 1: # v_center
                off_h = (h - h1)/2
            else:   # v_bottom
                off_h = h - h1

            for i in range(0, len(tmps)):
                w1 = pdfmetrics.stringWidth(tmps[i], font.name, font.height)
                if xfa.hor_align == 1:  # left
                    off_w = 0
                elif xfa.hor_align == 2:  # center
                    off_w = (w - w1) / 2
                else:  # rigth
                    off_w = w - w1
                c.drawString(pos.x + off_w + 10, pos.y + cell_size.h - off_h - font.height - i * (1 + 0.2) * font.height, tmps[i])

        pos.x = pos.x + cell_size.w
    pos.x = start_pos.x
    pos.y = pos.y + cell_size.h
c.setLineWidth(4)
c.lines(line_list)
print(line_list)

# c.setFont("彩虹粗仿宋", 40)
# c.drawString(inch, 8 * inch, "(1,8) 彩虹粗仿宋")
# c.setStrokeColor(black)
# c.setLineWidth(10)
# c.rect(10, 10, need_size.w - 20, need_size.h - 20)
# c.setStrokeColor(red)
# c.setLineWidth(10)
# c.rect(100,100, 300,300)


c.showPage()
c.save()



# cell.ctype = ctype
# cell.value = value
# cell.xf_index = xf_index

# rs.books.xf_list[cell.]
pass



# cell = rs.cell(1, 0)
# cell1 = rs.cell(3, 0)
# i = 0
# for font in rs.book.font_list:
#     print('{} : {} {}'.format(i, rs.book.font_list[i].name, rs.book.font_list[i].height))
#     i = i + 1
# print(rs.rich_text_runlist_map)