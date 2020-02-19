#!/usr/bin/env python
# -*- coding: UTF-8 -*-
# from https://blog.csdn.net/XnCSD/article/details/85208303
"""=================================================
@Project -> File   ：mylearn -> Office2PDF
@IDE    ：PyCharm
@Author ：Mr. toiler
@Date   ：1/18/2020 11:18 AM
@Desc   ：
=================================================="""
import os, time
from win32com.client import Dispatch, constants, gencache, DispatchEx


class PDFConverter(object):
    def __init__(self):
        self.xlApp = None
        self.wdApp = None
        self.pptApp = None

    def open_excel_app(self):
        if self.xlApp:
            return self.xlApp
        else:
            self.xlApp = DispatchEx("Excel.Application")
            self.xlApp.Visible = False
            self.xlApp.DisplayAlerts = 0

    def open_word_app(self):
        if self.wdApp:
            return self.wdApp
        else:
            gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
            self.wdApp = Dispatch("Word.Application")

    def open_ppt_app(self):
        if self.pptApp:
            return self.pptApp
        else:
            gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
            self.pptApp = Dispatch("PowerPoint.Application")
            # self.pptApp.Visible = 0

    def get_pdf_path(self, office_file_path, pdf_file_path=None, is_overwrite=False):
        if pdf_file_path is None:
            pdf_file_path = office_file_path + '.pdf'
        if os.path.exists(pdf_file_path):
            if is_overwrite:
                os.remove(pdf_file_path)
            else:
                return None
        return pdf_file_path

    def convert_ppt(self, office_file_path, pdf_file_path=None, is_overwrite=False):
        self.open_ppt_app()
        pdf_file_path = self.get_pdf_path(office_file_path, pdf_file_path, is_overwrite)
        ppt = self.pptApp.Presentations.Open(office_file_path, True, False, False)
        ppt.ExportAsFixedFormat(pdf_file_path, FixedFormatType=2, PrintRange=None)
        # ppt.SaveAs(pdf_file_path, 32)

    def convert_word(self, office_file_path, pdf_file_path=None, is_overwrite=False):
        self.open_word_app()
        pdf_file_path = self.get_pdf_path(office_file_path, pdf_file_path, is_overwrite)
        doc = self.wdApp.Documents.Open(office_file_path)
        doc.ExportAsFixedFormat(pdf_file_path, constants.wdExportFormatPDF,
                                Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)

    def convert_excel(self, office_file_path, pdf_file_path=None, is_overwrite=False):
        self.open_excel_app()
        pdf_file_path = self.get_pdf_path(office_file_path, pdf_file_path, is_overwrite)
        books = self.xlApp.Workbooks.Open(office_file_path, False)
        # sheetName = books.Sheets(1).Name 获取名字
        # xlSheet = books.Worksheets(sheetName)
        try:
            i = 1
            while True:
                sheet = books.Sheets(i)  # 实际运行情况是没有表格了会抛出异常
                if sheet:
                    # 设置打印格式为一个Sheet表格打印到适应纸张大小,默认是实际大小，
                    # 纸张是默认的A4，具体更多的参数可以取 excel 里面录制宏，在宏代码里面可以参考
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = 1
                    i = i + 1
                    # break
                else:
                    break
        except Exception:
            pass

        books.ExportAsFixedFormat(0, pdf_file_path)
        books.Close(False)

    def quit(self):
        if self.xlApp:
            self.xlApp.Quit()
            self.xlApp = None
        if self.wdApp:
            self.wdApp.Quit(constants.wdDoNotSaveChanges)
            self.wdApp = None
        if self.pptApp:
            self.pptApp.Quit()
            self.pptApp = None

    def __del__(self):
        self.quit()


if __name__ == '__main__':
    """
    open excel app cost: 2.306511163711548
    convert excel to pdf one time cost: 6.887148380279541
    convert excel to pdf one time cost: 7.744565725326538
    quit excel app cost: 20.519250631332397
    open app and convert and quit app , excel to pdf one time cost: 11.418348789215088
    open word app cost: 5.517384767532349
    convert word to pdf one time cost: 4.319069862365723
    open ppt app cost: 3.7118749618530273
    convert ppt to pdf one time cost: 4.690080642700195
    
    Process finished with exit code 0
    """
    path = os.path.dirname(__file__).strip('src')
    pdf = PDFConverter()
    t = time.time()
    pdf.open_excel_app()
    print('open excel app cost: {}'.format(time.time() - t))
    t = time.time()
    pdf.convert_excel(os.path.join(path, "res/excel.xlsx"), is_overwrite=True)
    print('convert excel to pdf one time cost: {}'.format(time.time() - t))
    t = time.time()
    pdf.convert_excel(os.path.join(path, "res/orther.xlsx"), is_overwrite=True)
    print('convert excel to pdf one time cost: {}'.format(time.time() - t))
    t = time.time()
    pdf.xlApp.Quit()
    pdf.xlApp = None
    print('quit excel app cost: {}'.format(time.time() - t))
    t = time.time()
    pdf.convert_excel(os.path.join(path, "res/orther.xlsx"), os.path.join(path, "res/orther_01.pdf"), is_overwrite=True)
    pdf.xlApp.Quit()
    pdf.xlApp = None
    print('open app and convert and quit app , excel to pdf one time cost: {}'.format(time.time() - t))
    t = time.time()

    pdf.open_word_app()
    print('open word app cost: {}'.format(time.time() - t))
    t = time.time()

    pdf.convert_word(os.path.join(path, "res/word.docx"), is_overwrite=True)

    print('convert word to pdf one time cost: {}'.format(time.time() - t))
    t = time.time()

    pdf.open_ppt_app()
    print('open ppt app cost: {}'.format(time.time() - t))
    t = time.time()

    pdf.convert_ppt(os.path.join(path, r"res\ppt.pptx"), is_overwrite=True)

    print('convert ppt to pdf one time cost: {}'.format(time.time() - t))
    t = time.time()
