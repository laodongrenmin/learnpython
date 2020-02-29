#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""=================================================
@Project -> File   ：mylearn -> Excel2PDF
@IDE    ：PyCharm
@Author ：Mr. toiler
@Date   ：2/23/2020 21:56 PM
@Desc   ：利用 xlrd、reportlab 把excel文档转为PDF
=================================================="""
import xlrd
from xlrd.sheet import Sheet, Cell
from xlrd.formatting import *

from reportlab.pdfgen.canvas import Canvas

from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import black, Color
from reportlab.lib.styles import ParagraphStyle

from reportlab.lib.units import inch, mm

from reportlab.platypus import Paragraph, Frame


import os
import re


class PDFCNFontMng(object):
    FONTS = {'彩虹粗仿宋': 'chcfs.ttf', '宋体': 'simsun.ttc', '楷体': 'simkai.ttf', '仿宋': 'simfang.ttf',
             '黑体': 'simhei.ttf', '华文宋体': 'STSONG.TTF'}
    DEFAULT_FONT_NAME = '宋体'

    def __init__(self):
        self.font_names = pdfmetrics.getRegisteredFontNames()
        self.default_font_name = PDFCNFontMng.DEFAULT_FONT_NAME
        self.replace_font_map = dict()
        self.get_font_name(PDFCNFontMng.DEFAULT_FONT_NAME)

    def setFont(self, c, font_name, size=12, leading=None):
        font_name = self.get_font_name(font_name)
        c.setFont(font_name, size, leading if leading else size * 1.2)
        return font_name

    def get_font_name(self, font_name):
        """
        取注册到系统的字体，如果没有取到，再加载字体，如果加载不成功就用默认字体
        :param font_name:
        :return:
        """
        if font_name in self.font_names:
            ret_font_name = font_name
        elif font_name in self.replace_font_map:
            ret_font_name = self.replace_font_map.get(font_name)
        elif font_name in PDFCNFontMng.FONTS.keys():
            try:
                font = TTFont(font_name, PDFCNFontMng.FONTS.get(font_name))
                pdfmetrics.registerFont(font)
                self.font_names = pdfmetrics.getRegisteredFontNames()
                ret_font_name = font_name
            except Exception as e:
                self.replace_font_map[font_name] = self.default_font_name
                print('{} instead of {}'.format(font_name, self.default_font_name))
                font_name = self.default_font_name
                print(e)
        else:
            ret_font_name = self.default_font_name
        if font_name != ret_font_name:
            print('{} instead of {}'.format(font_name, ret_font_name))
        return ret_font_name


class DrawText(object):
    _fontMng = PDFCNFontMng()
    S_FTR = "<font name='{}' size='{}' color='#{}'>{}</font>"

    def __init__(self, s, rect):
        self.s = s
        self.rect = rect
        self.fontMng = DrawText._fontMng
        self.x, self.y = rect[0]
        self.w, self.h = rect[1]
        self.align: XFAlignment = s[0]
        self.wrapped = self.align.text_wrapped
        self.max_height = s[1]
        # report lab pdf left 0 center 1 right 2 adjust 4
        # exl left 1  center 2 right 3 unknow 0
        self.hor_align = s[0].hor_align - 1
        if self.hor_align < 0 or self.hor_align > 2: self.hor_align = 4
        self.vert_align = s[0].vert_align
        self.text = list(s[2])

    def set_font(self, c, font_name, height):
        """ 设置一种可用的中文字体，不存在就试图加载一种存在的并返回 """
        return self.fontMng.setFont(c, font_name, height)

    def draw(self, c: Canvas, pagesize=None):
        story = []
        a_w = pagesize
        need_h = 0
        need_w = self.w
        str_line = ''
        max_height = 0
        for v, font in self.text:  # 一行， 有多行的处理
            font_name, height, color = font
            max_height = max(max_height, height)
            font_name = self.set_font(c, font_name, height)
            v = re.sub(r'\n', '<br/>', v)
            v = re.sub(r'\s', '&nbsp;', v)
            hex_color = hex((color.red << 16) + (color.green << 8) + color.blue).upper()
            str_line = str_line + DrawText.S_FTR.format(font_name, height, hex_color[2:], v)

        style = ParagraphStyle(name='font', fontName=font_name, fontSize=max_height,
                               leading=max_height * 1.2, alignment=self.hor_align)
        print('-'*80)
        print(str_line)
        p = Paragraph(str_line, style)
        w, h = p.wrap(self.w, self.h)
        m_w = p.minWidth()
        if self.wrapped == 0 and m_w > self.w:  # 不自动换行
            need_w = m_w * 1.02
            w, h = need_w, max_height*1.2
        story.append(p)
        need_h = need_h + h
        str_line = ''
            # print('w,h:{},{} m_w:{} need_w h:{},{} v:{} '.format(w, self.h, m_w, need_w, need_h, v))

        # # 垂直对齐
        # align: XFAlignment
        if self.vert_align == 1:  # v_center
            self.y = self.y - (self.h - need_h) / 2
        elif self.vert_align == 2:  # v_bottom
            self.y = self.y - (self.h - need_h)
        else:  # v_top
            pass

        if self.hor_align == 1: # center
            self.x = self.x - (need_w - self.w)/2
        elif self.hor_align == 2: # right
            self.x = self.x - (need_w - self.w) - height * 0.2
        else:
            self.x = self.x + height * 0.2

        f = Frame(self.x, self.y, need_w, self.h, 0, 0, 0, 0, showBoundary=0)
        f.addFromList(story, c)


class Xls2PDF(PDFCNFontMng):
    def __init__(self, filename, is_overwrite=False):
        super(Xls2PDF, self).__init__()
        self.xls_path, self.xls_name = os.path.split(os.path.abspath(filename))
        self.is_overwrite = is_overwrite
        self.rb = xlrd.open_workbook(filename, formatting_info=True)

    def convert(self, pdf_path=None):
        for rs in self.rb.sheets()[2:3]:
            filename = self.get_pdf_filename(rs.name, pdf_path)
            pdf_size, col_widths, row_heights, zoom, start_pos = self.get_size(rs=rs, pagesize=(420*mm,297*mm))
            c = Canvas(filename, pagesize=pdf_size)
            line_list, bg_rect_list, cell_rect_map = self.get_lines(rs, col_widths, row_heights, start_pos)
            # draw background
            for color, pos, size in bg_rect_list:
                c.setFillColor(color)
                c.rect(pos[0], pos[1], size[0], size[1], 0, 1)
            # draw border
            c.setLineWidth(0.2)
            c.setFillColor(black)
            c.lines(line_list)
            # draw text
            self.fill_string(c, zoom, rs, cell_rect_map)

            c.showPage()
            c.save()

    @staticmethod
    def get_cell_font(book, xf_index=None, font_index=None):
        if xf_index:
            font_index = book.xf_list[xf_index].font_index
        font: Font = book.font_list[font_index]
        c_rgb = book.colour_map.get(font.colour_index)
        color = Color(0, 0, 0)
        if c_rgb:
            color = Color(c_rgb[0], c_rgb[1], c_rgb[2])
        return font.name, font.height, color

    def fill_string(self, c: Canvas, zoom, rs: Sheet, cell_rect_map):
        for n_row in range(0, rs.nrows):
            for n_col in range(0, rs.ncols):
                cell: Cell = rs.cell(n_row, n_col)
                v, cell_type, xf_index = cell.value, cell.ctype, cell.xf_index
                if v:
                    xf: XF = rs.book.xf_list[xf_index]
                    font_name, font_height, font_color = self.get_cell_font(rs.book, xf_index=xf.xf_index)
                    format_str = rs.book.format_map.get(xf.format_key).format_str
                    max_height = font_height/zoom
                    if cell_type == 2:   # 数字
                        tmp = self.format_number(format_str, v)
                        draw_text_list = [xf.alignment, max_height,
                                          [(tmp, (font_name, font_height/zoom, font_color))]]
                    elif cell_type == 3:  # date
                        tmp = self.format_date(format_str, xlrd.xldate_as_datetime(v, 0))
                        draw_text_list = [xf.alignment, max_height,
                                          [(tmp, (font_name, font_height/zoom, font_color))]]
                    else:
                        rich_text_list = rs.rich_text_runlist_map.get((n_row, n_col))
                        p_s = list()
                        if rich_text_list:
                            si = 0
                            rich_text_list = list(rich_text_list)
                            rich_text_list.append((len(v), xf.font_index))
                            for pi, fi in rich_text_list:
                                p_s.append((v[si:pi], (font_name, font_height/zoom, font_color)))
                                font_name, font_height, font_color = self.get_cell_font(rs.book, font_index=fi)
                                max_height = max(max_height, font_height / zoom)
                                si = pi
                            print(p_s)

                        else:
                            p_s.append((v, (font_name, font_height/zoom, font_color)))
                            print(p_s)

                        draw_text_list = [xf.alignment, max_height, p_s]
                    dt = DrawText(draw_text_list, cell_rect_map.get((n_row, n_col)))
                    dt.draw(c, None)

    @staticmethod
    def format_date(ft_str, d):
        if ft_str is None:
            ft_str = '[$-F800]'  # 默认为yyyy年mm月dd日
        if ft_str.find('[$-F800') != -1:
            str_d = d.strftime('%Y{}%m{}%d{}').format('年', '月', '日')
        elif ft_str.find('[$-F400]') != -1:  # 显示 hh:mm:ss
            str_d = d.strftime('%H:%M:%S')
        else:
            str_d = d.strftime('%x')
        return str_d

    @staticmethod
    def format_number(ft_str, v):
        # 没有处理前面的货币符号 "¥"* #,##0.00_ ;_ "¥"* \-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_
        if ft_str.find('#,##') != -1:
            thousands = ','
        else:
            thousands = ''
        find_strs = re.findall('0\.0+_', ft_str)
        if len(find_strs) > 0:  # 要统计小数点后面有几个0，就是保留几位小数
            decimal_digits = "{}.{}f".format(thousands, len(find_strs[0]) - 3)
        else:
            decimal_digits = "{}".format(thousands)
            if v % 1 > 0.0:
                pass
            else:
                v = int(v)
        return format(v, decimal_digits)

    def get_lines(self, rs: Sheet,  widths,  heights, start_pos=(0, 0)):
        """得到绘制的表格线、背景框底色和坐标、实际表格的坐标"""
        n_cols = len(widths)
        n_rows = len(heights)
        line_list = list()
        bg_rect = list()
        merged_cell_rect = dict()
        pos = [start_pos[0], start_pos[1]]
        leave_lines, merged_cell_size = self.get_merged_cell_info(rs, widths, heights)
        for n_row in range(n_rows-1, -1, -1):
            for n_col in range(0, n_cols):
                key = (n_row, n_col)
                xf: XF = rs.book.xf_list[rs.cell(n_row, n_col).xf_index]
                cell_size = (widths[n_col], heights[n_row])  # w, h
                leave_line_flag = leave_lines.get(key, [True, True, True, True])
                self.add_cell_lines(line_list, xf.border, leave_line_flag, pos, cell_size)
                bg: XFBackground = xf.background
                if bg.fill_pattern == 1:
                    rgb = rs.book.colour_map.get(bg.pattern_colour_index)
                    if rgb:
                        bg_rect.append((Color(rgb[0], rgb[1], rgb[2]), (pos[0], pos[1]), cell_size))
                # 把实际单元格的大小也一并返回去，合并和非合并的单元格都返回去
                if key in merged_cell_size.keys():
                    s = merged_cell_size[key]
                    merged_cell_rect[key] = ((pos[0], pos[1] + cell_size[1] - s[1]), s)
                else:
                    merged_cell_rect[key] = ((pos[0], pos[1]), cell_size)
                pos[0] = pos[0] + cell_size[0]
            pos[0] = start_pos[0]
            pos[1] = pos[1] + cell_size[1]
        return line_list, bg_rect, merged_cell_rect

    @staticmethod
    def add_cell_lines(ls, b, flag, pos, cell_size):
        def add_cell_line(l_s, line):
            if line not in l_s:
                l_s.append(line)
        bl, bt, br, bb = b.left_line_style, b.top_line_style, b.right_line_style, b.bottom_line_style
        if flag[0] and (bl != 0):  # 实际上是线性， 不同的线，比如虚线, 线的粗细也没有管，有线就画一条一样的
            add_cell_line(ls, (pos[0], pos[1], pos[0], pos[1] + cell_size[1]))
        if flag[1] and (bt != 0):
            add_cell_line(ls, (pos[0],  pos[1] + cell_size[1], pos[0] + cell_size[0],  pos[1] + cell_size[1]))
        if flag[2] and (br != 0):
            add_cell_line(ls, (pos[0] + cell_size[0], pos[1], pos[0] + cell_size[0],  pos[1] + cell_size[1]))
        if flag[3] and (bb != 0):
            add_cell_line(ls, (pos[0],  pos[1], pos[0] + cell_size[0],  pos[1]))

    @staticmethod
    def get_merged_cell_info(rs, widths, heights):
        """得到合并表格真正需要留下的边框， 合并的内部表格，不需要保留边框线"""
        leave_lines = dict()  # 合并单元格，内部的线就不需要了，需要删除
        merged_cell_size = dict()  # 合并单元格，真正的宽度和高度
        for row1, row2, col1, col2 in rs.merged_cells:
            merged_cell_size[(row1, col1)] = (sum(widths[col1:col2]), sum(heights[row1:row2]))
            for n_row in range(row1, row2):
                for n_col in range(col1, col2):
                    border = (n_col == col1, n_row == row1, n_col == (col2 - 1), n_row == (row2 - 1))
                    leave_lines[(n_row, n_col)] = border
        return leave_lines, merged_cell_size

    def get_pdf_filename(self, sheet_name, pdf_path=None):
        if not pdf_path:
            pdf_path = self.xls_path
        if not os.path.exists(pdf_path):
            os.makedirs(pdf_path)
        filename = "{}_{}.pdf".format(os.path.join(pdf_path, self.xls_name), sheet_name)
        if os.path.exists(filename):
            if self.is_overwrite:
                os.remove(filename)
            else:
                raise FileExistsError('{} exists'.format(filename))
        return filename

    @staticmethod
    def get_size(rs: Sheet, pagesize=A4, margin=0.04):
        """
        得到在某种纸张下，需要的合适的大小以及缩放比例, 固定转换适合纸张大小，如果行数过多，只取纸张高能够放下的内容
        :param self:
        :param rs:
        :param pagesize:
        :param margin:
        :return: （pagesize_w, pagesize_h）,(ws, hs), zoom, start_pos  PDF’w&h、SHEET’各个列宽和行高、缩放比例、开始坐标
        """
        n_cols = rs.ncols
        n_rows = rs.nrows

        # # 去除后面的空列, 有时候 xlrd 返回的列数会多一些
        has_empty_col = True
        while has_empty_col:
            n_cols = n_cols - 1
            for n_row in range(0, n_rows):
                if rs.cell(n_row, n_cols).value:
                    has_empty_col = False
                    break
        n_cols = n_cols + 1

        def get_col_width(n_col, standard_width = 2304):
            col_info = rs.colinfo_map.get(n_col)
            return standard_width if col_info is None else col_info.width

        def get_row_height(n_row, default_heigth = 288):
            row_info = rs.rowinfo_map.get(n_row)
            return default_heigth if row_info is None else row_info.height

        # 宽除以2.54是蒙的数据，感觉这个和真实的差不多，没有找到相关资料
        col_widths = list([get_col_width(n_col, rs.standardwidth) / 2.54 for n_col in range(0, n_cols)])
        row_heights = list([get_row_height(n_row, rs.default_row_height) for n_row in range(0, n_rows)])
        p_size = (pagesize[0] * (1-margin), pagesize[1] * (1-margin))  # 准备页边距留 4%， 1000 留 40点
        zoom1 = max(sum(col_widths), sum(row_heights)) / max(p_size)
        zoom2 = min(sum(col_widths), sum(row_heights)) / min(p_size)
        zoom = max(zoom1, zoom2)
        if zoom < 1:
            zoom = 1
        col_widths = [c_w/zoom for c_w in col_widths]
        row_heights = [r_h/zoom for r_h in row_heights]

        # todo: 当一页无法放下时候，没有考虑，无论如何都会压到一页上，哪怕字体很小，以后添加字体很小的判断，保证可以看清楚
        s_w = sum(col_widths)
        s_h = sum(row_heights)
        if s_w > s_h:
            w, h = (max(pagesize), min(pagesize))
        else:
            w, h = (min(pagesize), max(pagesize))
        x = (w - s_w)/2
        y = (h - s_h)/2
        return (w, h), col_widths, row_heights, zoom, (x, y)


if __name__ == '__main__':
    import time
    t = time.time()
    excel_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/orther.xls'
    xls2pdf = Xls2PDF(excel_file_path, True)
    xls2pdf.convert()
    print('time cost: {}'.format(time.time() - t))
