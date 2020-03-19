#!/usr/bin/env python
# -*- coding: UTF-8 -*-
import os
import sys
import chardet

from reportlab.pdfgen.canvas import Canvas
from fontTools.ttLib import TTFont
import json

import numpy as np
from PIL import Image, ImageDraw, ImageFont

__all__ = ['get_pdf_filename', 'PDFWriter', 'get_absolute_path', 'guess_encoding',
           'get_problem_hz']


def get_problem_hz(font_file):
    """得到某种字库不支持的汉子显示列表"""
    if not os.path.exists(font_file):
        font_file = os.path.join(os.getenv("windir"), 'fonts/{}'.format(font_file))
    p_f = get_absolute_path(os.path.basename(font_file) + "_p.dat")
    hzs = dict()
    if os.path.exists(p_f):
        with open(p_f, 'rb') as f:
            buf = f.read()
        for s in buf.decode('UTF-8'):
            hzs[s] = s
        return hzs
    array = np.ndarray((24, 24, 3), np.uint8)
    array[:, :, 0] = 255
    array[:, :, 1] = 255
    array[:, :, 2] = 255
    image = Image.fromarray(array)
    # 创建绘制对象
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(font_file, 24)

    for cn in range(19968, 40870):  # 19968 - 40869  '\u4e00', '\u9fa5'
        s = '\\u{}'.format(hex(cn)[2:])
        s = json.loads(f'"{s}"')
    # for s in '李傕':
        draw.text((0, 0), s, font=font, fill=(0, 0, 0, 0))
        count = 0
        for x in range(0, 24):
            for y in range(0, 24):
                r, g, b = image.getpixel((x, y))
                if r < 50:
                    count = count + 1
                if count > 5:
                    break
            else:
                continue
            break;
        if count < 5:
            hzs[s] = s
        draw.text((0, 0), s, font=font, fill=(255, 255, 255, 255))

    tmp = []
    for k in hzs.keys():
        tmp.append(k)
    with open(p_f, 'wb') as f:
        f.write(''.join(tmp).encode('UTF-8'))
    return hzs

def get_absolute_path(filename):
    """得到与执行文件一个路径下的文件的绝对路径"""
    module_path = os.path.dirname(sys._getframe(1).f_code.co_filename)
    return os.path.join(module_path, filename)


def guess_encoding(buf, default_encoding='UTF-8'):
    """探测字符编码，返回字符可能的编码"""
    return chardet.detect(buf).get('encoding', default_encoding)


def get_hz_from_ttf(font_file):
    """返回某个字库里面uni开头的全部UNICODE字"""
    if not os.path.exists(font_file):
        font_file = os.path.join(os.getenv("windir"), 'fonts/{}'.format(font_file))
    font = TTFont(font_file, fontNumber=0)
    glyphNames = font.getGlyphNames()
    all_hz = dict()
    for hz in glyphNames:
        if len(hz) == 7 and hz[:3] == 'uni':
            sz_uni = '\\u{}'.format(hz[3:])
            try:
                hz = json.loads(f'"{sz_uni}"')
                if '\u4e00' <= hz <= '\u9fff':
                    all_hz[hz] = sz_uni
            except:
                pass
    return all_hz


def get_pdf_filename(xls_path, xls_name, sheet_name, pdf_path=None, is_overwrite=False):
    if not pdf_path:
        pdf_path = xls_path
    if not os.path.exists(pdf_path):
        os.makedirs(pdf_path)
    filename = "{}_{}.pdf".format(os.path.join(pdf_path, xls_name), sheet_name)
    if os.path.exists(filename):
        if is_overwrite:
            os.remove(filename)
        else:
            raise FileExistsError('{} exists'.format(filename))
    return filename


class PDFWriter(object):
    def __init__(self, filename, pdf_size):
        self.c = Canvas(filename, pagesize=pdf_size)

