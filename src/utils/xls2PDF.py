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
from xlrd.sheet import Sheet
from xlrd.formatting import *

from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import black, Color

import os


class PDFCNFontMng(object):
    FONTS = {'彩虹粗仿宋':'chcfs.ttf', '宋体':'simsun.ttc', '楷体':'simkai.ttf', '仿宋':'simfang.ttf',
             '黑体':'simhei.ttf'}
    DEFAULT_FONT_NAME = '宋体'

    def __init__(self):
        self.font_names = pdfmetrics.getRegisteredFontNames()
        self.default_font_name = PDFCNFontMng.DEFAULT_FONT_NAME
        self.replace_font_map = dict()
        self.get_font_name(PDFCNFontMng.DEFAULT_FONT_NAME)

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
                self.replace_font_map[font_name, self.default_font_name]
                font_name = self.default_font_name
                print(e)
        else:
            ret_font_name = self.default_font_name
        return ret_font_name


class Xls2PDF(PDFCNFontMng):
    def __init__(self, filename, is_overwrite=False):
        super(Xls2PDF, self).__init__()
        self.xls_path, self.xls_name = os.path.split(os.path.abspath(filename))
        self.is_overwrite = is_overwrite
        self.rb = xlrd.open_workbook(filename, formatting_info=True)

    def convert(self, pdf_path=None):
        for rs in self.rb.sheets():
            filename = self.get_pdf_filename(rs.name, pdf_path)
            pdf_size, col_widths, row_heights, zoom, start_pos = self.get_size(rs=rs)
            c = Canvas(filename, pagesize=pdf_size)
            print(col_widths)
            line_list = self.get_lines(rs, col_widths, row_heights, start_pos)


            print('pdf_size:{} w: {} h:{} zoom:{} start_pos:{}'.format(
                pdf_size, sum(col_widths), sum(row_heights), zoom, start_pos))

            c.setLineWidth(0.2)
            c.setFillColor(black)
            c.lines(line_list)

            c.showPage()
            c.save()

    def get_lines(self, rs: Sheet,  widths,  heights, start_pos=(0, 0)):
        n_cols = len(widths)
        n_rows = len(heights)
        line_list = list()
        pos = [start_pos[0], start_pos[1]]
        leave_lines = self.get_merged_cell_leave_line(rs)
        for n_row in range(n_rows-1, -1, -1):
            for n_col in range(0, n_cols):
                border = rs.book.xf_list[rs.cell(n_row, n_col).xf_index].border
                cell_size = (widths[n_col], heights[n_row])  # w, h
                leave_line_flag = leave_lines.get('{}_{}'.format(n_row, n_col), [True, True, True, True])
                self.add_cell_lines(line_list, border, leave_line_flag, pos, cell_size)
                pos[0] = pos[0] + cell_size[0]
            pos[0] = start_pos[0]
            pos[1] = pos[1] + cell_size[1]
        print(len(line_list))
        return line_list

    def add_cell_lines(self, ls, b, flag, pos, cell_size):
        bl, bt, br, bb = b.left_line_style, b.top_line_style, b.right_line_style, b.bottom_line_style
        lines = list()
        if flag[0] and (bl != 0):  # 实际上是线性， 不同的线，比如虚线, 线的粗细也没有管，有线就画一条一样的
            self.add_cell_line(ls, (pos[0], pos[1], pos[0], pos[1] + cell_size[1]))
        if flag[1] and (bt != 0):
            self.add_cell_line(ls, (pos[0],  pos[1] + cell_size[1], pos[0] + cell_size[0],  pos[1] + cell_size[1]))
        if flag[2] and (br != 0):
            self.add_cell_line(ls, (pos[0] + cell_size[0], pos[1], pos[0] + cell_size[0],  pos[1] + cell_size[1]))
        if flag[3] and (bb != 0):
            self.add_cell_line(ls, (pos[0],  pos[1], pos[0] + cell_size[0],  pos[1]))

    @staticmethod
    def add_cell_line(ls, line):
        if line not in ls:
            ls.append(line)

    @staticmethod
    def get_merged_cell_leave_line(rs):
        """
        得到合并表格真正需要留下的边框， 合并的内部表格，不需要保留边框线
        :param rs:
        :return:
        """
        leave_lines = dict()  # 合并单元格，内部的线就不需要了，需要删除
        for row1, row2, col1, col2 in rs.merged_cells:
            key = '{}_{}'.format(row1, col1)
            for n_row in range(row1, row2):
                for n_col in range(col1, col2):
                    bl, bt, br, bb = 0, 1, 2, 3
                    border = list([False, False, False, False])
                    if n_row == row1:
                        border[bt] = True
                    if n_row == (row2 - 1):
                        border[bb] = True
                    if n_col == col1:
                        border[bl] = True
                    if n_col == (col2 - 1):
                        border[br] = True
                    leave_lines['{}_{}'.format(n_row, n_col)] = border
        return leave_lines

    def test(self, c, start_pos, name):
        font_height = 20
        c.setFont(self.get_font_name('宋体'), font_height)
        c.drawString(start_pos[0] + font_height * 0.2, start_pos[1] + font_height * 0.2, name)
        c.line(start_pos[0], start_pos[1], 800, start_pos[1])

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

        # 宽除以2.54是蒙的数据，感觉这个和真实的差不多，没有找到相关资料
        col_widths = list([rs.colinfo_map.get(n_col).width / 2.54 for n_col in range(0, n_cols)])
        row_heights = list([rs.rowinfo_map.get(n_row).height for n_row in range(0, n_rows)])
        p_size = (pagesize[0] * (1-margin), pagesize[1] * (1-margin))  # 准备页边距留 4%， 1000 留 40点
        zoom1 = max(sum(col_widths), sum(row_heights)) / max(p_size)
        zoom2 = min(sum(col_widths), sum(row_heights)) / min(p_size)
        zoom = max(zoom1, zoom2)
        if zoom < 1:
            zoom = 1
        col_widths = [c_w/zoom for c_w in col_widths]
        row_heights = [r_h/zoom for r_h in row_heights]
        if col_widths > row_heights:
            w, h = (max(pagesize), min(pagesize))
        else:
            w, h = (min(pagesize), max(pagesize))
        # todo: 当一页无法放下时候，没有考虑，无论如何都会压到一页上，哪怕字体很小，以后添加字体很小的判断，保证可以看清楚
        s_w = sum(col_widths)
        s_h = sum(row_heights)
        x = (w - s_w)/2
        y = (h - s_h)/2
        return (w, h), col_widths, row_heights, zoom, (x, y)


if __name__ == '__main__':
    import time
    t = time.time()
    excel_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/excel.xls'
    xls2pdf = Xls2PDF(excel_file_path, True)
    xls2pdf.convert()
    print('time cost: {}'.format(time.time() - t))
