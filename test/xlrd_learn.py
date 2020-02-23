#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import xlrd
from xlrd.formatting import XFAlignment, XF

from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas


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

rs = rb.sheets()[1]

ncols = rs.ncols
nrows = rs.nrows

# ncols 有空行，导致生成的文件不对称
has_emplty_col = True
while has_emplty_col:
    ncols = ncols - 1
    for rowx in range(0, nrows):
        if rs.cell(rowx, ncols).value:
            has_emplty_col = False
            break


ncols = ncols + 1


col_widths = list([rs.colinfo_map.get(colx).width/2.54 for colx in range(0, ncols)])
row_heights = list([rs.rowinfo_map.get(rowx).height for rowx in range(0, nrows)])

zoom = max(sum(col_widths), sum(row_heights)) / a4_size.h
col_widths = list([width/zoom for width in col_widths])
row_heights = list([height/zoom for height in row_heights])

need_size = Size(w=sum(col_widths) * 1.1,
                 h=sum(row_heights) * 1.1)

print(need_size)

start_pos = Pos(x=need_size.w * 0.05, y=need_size.h * 0.05)

c = canvas.Canvas("c:/Users/zbd/PycharmProjects/mylearn/res/hello.pdf", pagesize=(need_size.w, need_size.h))

merged_cells = dict()  # 保存合并单元格的尺寸，便于把内容输出到合适的位置
leave_lines = dict()  # 合并单元格，内部的线就不需要了，需要删除
for row1, row2, col1, col2 in rs.merged_cells:
    key = '{}_{}'.format(row1, col1)
    size = Size(w=sum(col_widths[col1:col2]), h=sum(row_heights[row1:row2]))
    for rowx in range(row1, row2):
        for colx in range(col1, col2):
            bl, bt, br, bb = 0, 1, 2, 3
            border = list([False, False, False, False])
            if rowx == row1:
                border[bt] = True
            if rowx == (row2 - 1):
                border[bb] = True
            if colx == col1:
                border[bl] = True
            if colx == (col2 - 1):
                border[br] = True
            leave_lines['{}_{}'.format(rowx, colx)] = border

    merged_cells[key] = [size, (row1, row2, col1, col2)]



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
ht = TTFont('黑体', 'simhei.ttf')

pdfmetrics.registerFont(chcfs)
pdfmetrics.registerFont(hwhp)
pdfmetrics.registerFont(st)
pdfmetrics.registerFont(kt)
pdfmetrics.registerFont(hwst)
pdfmetrics.registerFont(fs)
pdfmetrics.registerFont(ht)

line_list = list()
pos = Pos(start_pos.x, start_pos.y)
for rowx in range(nrows-1, -1, -1):
    for colx in range(0, ncols):
        cell_size = Size(w=col_widths[colx], h=row_heights[rowx])
        cell = rs.cell(rowx, colx)
        xf:XF = rs.book.xf_list[cell.xf_index]
        b = xf.border
        bl, bt, br, bb = b.left_line_style, b.top_line_style, b.right_line_style, b.bottom_line_style
        d = leave_lines.get('{}_{}'.format(rowx, colx), [True, True, True, True])
        if d[0] and (bl != 0):   # 实际上是线性， 不同的线，比如虚线
            line_list.append((pos.x, pos.y, pos.x, pos.y+cell_size.h))
        if d[1] and (bt != 0):
            line_list.append((pos.x, pos.y + cell_size.h, pos.x + cell_size.w, pos.y + cell_size.h))
        if d[2] and (br != 0):
            line_list.append((pos.x + cell_size.w, pos.y, pos.x + cell_size.w, pos.y + cell_size.h))
        if d[3] and (bb != 0):
            line_list.append((pos.x, pos.y, pos.x + cell_size.w, pos.y))

        if cell.value:
            m = merged_cells.get('{}_{}'.format(rowx, colx))
            if m:
                w, h = m[0].w, m[0].h
            else:
                w, h = cell_size.w, cell_size.h

            font = rs.book.font_list[rs.book.xf_list[cell.xf_index].font_index]
            font_height = int(font.height / zoom)
            if font.name in pdfmetrics.getRegisteredFontNames():
                # print("value:{} font name:{} height:{}".format("{}".format(cell.value), font.name, font.height))
                c.setFont(font.name, font_height)
            else:
                c.setFont("彩虹粗仿宋", font_height)
                print("value:{} font name:{} 彩虹粗仿宋 height:{}".format("{}".format(cell.value), font.name, font_height))
                font.name = "彩虹粗仿宋"

            h1 = font_height  # 当为一行时候， 当 自动换行且宽度不够时候以及有多行时
            if cell.ctype == 1:  # string
                v = "{}".format(cell.value)
                tmps = v.split("\n")
                if len(tmps) > 1:
                    h1 = len(tmps) * (1 + 0.2) * h1  # 0.2倍行距
            elif cell.ctype == 2:  # float
                if cell.value % 1 == 0.0:
                    tmps = list(['{}'.format(int(cell.value))])
                else:
                    tmps = list(['{}'.format(cell.value)])
            elif cell.ctype == 3:   # date
                d = xlrd.xldate_as_datetime(cell.value, 0)
                try:
                    format_str = rs.book.format_map.get(xf.format_key).format_str
                    print(format_str)
                except:
                    pass
                if format_str is None:
                    format_str = '[$-F800]'   #  默认为yyyy年mm月dd日
                if format_str.find('[$-F800') != -1:
                    d = d.strftime('%Y{}%m{}%d{}').format('年', '月', '日')
                elif format_str.find('[$-F400]') != -1:  # 显示 hh:mm:ss
                    d = d.strftime('%H:%M:%S')
                else:
                    d = d.strftime('%x')
                print("format:{} value:{} to {}".format(format_str, cell.value, d))
                tmps = ["{}".format(d)]
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

            rich_text_list = rs.rich_text_runlist_map.get((rowx, colx))
            if rich_text_list and len(tmps) == 1:  # 某个单元格里面有自定义的格式, 且只有一行，其余的就当单元个的格式
                v = tmps[0]
                rich_text_list = list(rich_text_list)
                rich_text_list.append((len(v), font.font_index))
                w1 = 0
                s = 0
                for pi, fi in rich_text_list:
                    w1 = w1 + pdfmetrics.stringWidth(v[s:pi], font.name, font_height)
                    s = pi
                    print(pi, rs.book.font_list[fi].name)
                    font = rs.book.font_list[fi]
                    font_height = font.height / zoom
                if xfa.hor_align == 1:  # left
                    off_w = 0
                elif xfa.hor_align == 2:  # center
                    off_w = (w - w1) / 2
                else:  # rigth
                    off_w = w - w1
                if off_w < 0:  off_w = 0  # 不能放下就最靠左
                s = 0
                for pi, fi in rich_text_list:
                    w1 = pdfmetrics.stringWidth(v[s:pi], font.name, font_height)
                    c.drawString(pos.x + off_w,
                                 pos.y + cell_size.h - off_h - font_height, v[s:pi])
                    s = pi
                    font = rs.book.font_list[fi]
                    font_height = font.height / zoom
                    off_w = off_w + w1
                    c.setFont(font.name, font_height)
                pass

            else:
                for i in range(0, len(tmps)):
                    w1 = pdfmetrics.stringWidth(tmps[i], font.name, font_height)
                    if xfa.hor_align == 1:  # left
                        off_w = 0
                    elif xfa.hor_align == 2:  # center
                        off_w = (w - w1) / 2
                    else:  # rigth
                        off_w = w - w1
                    if off_w < 0:  off_w = 0   # 不能放下就最靠左
                    c.drawString(pos.x + off_w,
                                 pos.y + cell_size.h - off_h - font_height - i * (1 + 0.2) * font_height, tmps[i])

        pos.x = pos.x + cell_size.w
    pos.x = start_pos.x
    pos.y = pos.y + cell_size.h
c.setLineWidth(1)
c.lines(line_list)
print(line_list)

c.showPage()
c.save()