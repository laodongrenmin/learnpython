#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""=================================================
@Project -> File   ：mylearn -> Excel2PDF
@IDE    ：PyCharm
@Author ：Mr. toiler
@Date   ：2/20/2020 17:54 PM
@Desc   ：利用 openpyxl、reportlab 把excel文档转为PDF
=================================================="""


class Excel2PDF(object):
    def __init__(self, excel_file_path, pdf_file_path=None, is_overwrite=False):
        self.excel_file_path = excel_file_path
        self.pdf_file_path = pdf_file_path
        self.is_overwrite = is_overwrite

    def convert(self):

        from xlutils.copy import copy
        import xlrd
        rb = xlrd.open_workbook(self.excel_file_path, formatting_info=True)
        rs = rb.sheet_by_index(0)
        cell = rs.cell(1, 0)
        cell1 =  rs.cell(3, 0)
        i = 0
        for font in rs.book.font_list:
            print('{} : {} {}'.format(i, rs.book.font_list[i].name, rs.book.font_list[i].height))
            i = i + 1
        print(rs.rich_text_runlist_map)
        wb = copy(rb)
        # 通过get_sheet()获取的sheet有write()方法
        ws = wb.get_sheet(0)
        # 获得到sheet了 可以进行 追加写 或者 修改某个单元格数据的操作了 最后不要忘了 save()
        pass






if __name__ == '__main__':
    excel_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/excel.xls'
    pdf_file_path = r'c:/Users/zbd/PycharmProjects/mylearn/res/excel.pdf'
    excel2pdf = Excel2PDF(excel_file_path, pdf_file_path)
    excel2pdf.convert()

    print('ok')