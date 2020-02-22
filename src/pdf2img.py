#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import sys, fitz
import os
import time


def pdf_convert2png(pdf_file_path, image_path):
    pdf_doc = fitz.open(pdf_file_path)
    for pg in range(pdf_doc.pageCount):
        page = pdf_doc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。
        # 此处若是不做设置，默认图片大小为：792X612, dpi=96
        zoom_x = 1.3  # 1.33333333  # (1.33333333-->1056x816)   (2-->1584x1224)
        zoom_y = 1.3  # 1.33333333
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)
        if not os.path.exists(image_path):  # 判断存放图片的文件夹是否存在
            os.makedirs(image_path)   # 若图片文件夹不存在就创建
        pix.writePNG(image_path+'/'+'images_%s_%d_%d.png' % (pg, 792 * zoom_x, 612 * zoom_y))  # 将图片写入指定的文件夹内


if __name__ == "__main__":
    start_time = time.time()
    # 开始时间
    path = os.path.dirname(__file__).strip('src')
    pdfPath = path + r'res/hello.pdf'
    imagePath = path + 'res'
    pdf_convert2png(pdfPath, imagePath)

    end_time = time.time()  # 结束时间
    print('cost time: {} ms'.format(1000* (end_time - start_time)))

