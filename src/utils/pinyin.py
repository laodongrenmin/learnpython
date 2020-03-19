#!/usr/bin/env python
# -*- coding: UTF-8 -*-
__version__ = '0.0.0.1'
__doc__ = """为汉语拼音封装一些工具。

为文本文件提供拼音标注能力， 把不认识的字标注上拼音，认识的字从自定义认识的字
或者加上常用字1500、常用字3500，生成带拼音标注的文本文件、PDF文件。

比如文字 '出差盱眙,吃小龙虾,差钱', 生成带拼音的文件格式如下
 
'出差(chāi)盱(xū)眙(yí),吃小龙虾,差(chà)钱(qián)'

'  chāi xū yí        chà qián'
'出 差  盱 眙,吃小龙虾,差   钱 '
"""
import re
from functools import lru_cache
from pypinyin import pinyin
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import inch, mm
from reportlab.pdfbase.ttfonts import TTFont
from chardet import detect
from utils import get_absolute_path, get_problem_hz


def get_pinyin_text(text, used_char=None, knows_char=None):
    """得到带生字拼音的字符串"""
    return Assistant.get_py(text, used_char, knows_char)


def get_pinyin_file(src_file, dst_file, encoding=None, used_char=None, knows_char=None):
    """从文本文件中得到带生字拼音的字符串并保存到文件中"""
    with open(src_file, 'rb') as f:
        buf = f.read()
    if encoding is None:
        encoding = detect(buf).get('encoding', 'UTF-8')
    text = buf.decode(encoding)
    if dst_file.lower().endswith('.pdf'):
        new_list = Assistant.get_py2pdf(dst_file, text, used_char, knows_char)
    else:
        text, new_list = Assistant.get_py(text, used_char, knows_char)
        with open(dst_file, 'wb') as f:
            f.write(text.encode(encoding))
    return new_list


# 注册微软雅黑
def register_font():
    if 'MSYH' in pdfmetrics.getRegisteredFontNames():
        return
    fonts = ['CHCFS.TTF', 'MSYH.TTC']
    for f in fonts:
        pdfmetrics.registerFont(TTFont(f.split('.')[0], f))


class Constant(object):
    FONT_NAME = 'CHCFS'
    FONT_SIZE = 16
    PY_FONT_NAME = 'MSYH'           # 拼音字体
    PY_FONT_SIZE = 8                # 拼音字体大小
    LINE_SPACE = FONT_SIZE * 1.0    # 行间距
    PAGE_SIZE = 210 * mm, 297 * mm  # 纸张大小
    MARGIN = (10 * mm, 10 * mm, 10 * mm, 10 * mm)  # 边距
    CHAR_SPACE = 1.05  # 字间距


class Assistant(object):
    register_font()

    CHCFS = get_problem_hz('chcfs.ttf')

    def __init__(self):
        self.version = __version__

    @staticmethod
    def get_py2pdf(dst_file, file_content, used_char=None, knows_char=None):
        pagesize = Constant.PAGE_SIZE
        c = canvas.Canvas(dst_file, pagesize=pagesize)
        margin = Constant.MARGIN
        frequently_used_char = get_hz(used_char, knows_char)
        x, y = margin[0], margin[1]
        w, h = pagesize[0] - x - margin[2], pagesize[1] - y - margin[3]
        pos = (x, h)
        y = pos[1] + Constant.LINE_SPACE
        new_word = []

        ps = re.split(r'(\r\n)+', file_content)
        for text in ps:
            if text == '\r\n' or text == '':
                continue
            mp = MyParagraph(text, frequently_used_char)
            while mp.has_next():
                n_w, n_h, hz_list = mp.next(w, h)
                y = y - n_h - Constant.LINE_SPACE
                x = pos[0]
                if y < margin[3]:
                    y = pos[1]
                    c.showPage()

                for hz in hz_list:
                    hz.draw(c, x, y)
                    if hz.py:
                        new_word.append('{}({})'.format(hz.hz, hz.py))
                    x = x + hz.width

        c.showPage()
        c.save()
        return new_word

    @staticmethod
    def get_py(text, used_char=None, knows_char=None):
        frequently_used_char = get_hz(used_char, knows_char)
        p1 = MyParagraph(text, frequently_used_char, first_line_space=0)
        content = []
        new_word = []
        for hz in p1.hz_list:
            content.append(hz.hz)
            if hz.py:
                content.append('({})'.format(hz.py))
                new_word.append('{}({})'.format(hz.hz, hz.py))
        return ''.join(content), new_word


class HZ(object):
    def __init__(self, hz, hz_font_name=Constant.FONT_NAME, hz_font_size=Constant.FONT_SIZE, hz_py_spacing=None,
                 py=None, py_font_name=Constant.PY_FONT_NAME, py_font_size=Constant.PY_FONT_SIZE, is_mark=False):
        self.is_mark = is_mark
        self.hz = hz
        self.hz_font_name = hz_font_name
        self.hz_font_size = hz_font_size
        self.py = py
        self.py_font_name = py_font_name
        self.py_font_size = py_font_size if py else 0
        self.hz_py_spacing = hz_py_spacing if hz_py_spacing else self.py_font_size * 0.2
        self.width = self.hz_font_size
        self.height = self.hz_font_size
        self.hz_width = self.hz_font_size
        self.py_width = self.py_font_size
        self.get_size()

    def draw(self, _canvas: canvas.Canvas, x,  y, align='center'):
        self.get_size()
        if self.py:
            _canvas.setFont(self.py_font_name, self.py_font_size)
            _canvas.drawString(x + (self.width - self.py_width)/2, y + self.hz_font_size + self.hz_py_spacing, self.py)
        _canvas.setFont(self.hz_font_name, self.hz_font_size)
        _canvas.drawString(x + (self.width - self.hz_width)/2, y, self.hz)

    def get_size(self):
        self.width, self.height, self.hz_width, self.py_width, self.hz_font_name = \
            self._compute_size(self.hz, self.hz_font_name, self.hz_font_size, self.py, self.py_font_name,
                               self.py_font_size, self.hz_py_spacing, self.is_mark)
        return self.width, self.height

    @staticmethod
    @lru_cache()
    def _compute_size(hz, hz_font_name, hz_font_size, py, py_font_name, py_font_size, hz_py_spacing, is_mark):
        if hz_font_name == 'CHCFS':
            if (not is_mark) and Assistant.CHCFS.get(hz):
                print(hz)
                hz_font_name = "MSYH"
        hz_width = max(pdfmetrics.stringWidth(hz, hz_font_name, hz_font_size), hz_font_size)
        if py:
            py_width = pdfmetrics.stringWidth(py, py_font_name, py_font_size)
            width = max(py_width, hz_width)
            height = hz_font_size + (py_font_size + hz_py_spacing)
        else:
            py_width = 0
            width = hz_width
            height = hz_font_size

        return width * Constant.CHAR_SPACE, height, hz_width, py_width, hz_font_name

    def __str__(self):
        return self.hz


class MyParagraph(object):
    def __init__(self, content, recognized_words, first_line_space=Constant.FONT_SIZE*2):
        self.recognized_words = recognized_words
        self.content = content
        self.first_line_space = first_line_space
        self.hz_list = self._parse()
        self.length = len(self.hz_list)
        self.cur_pos = 0

    def has_next(self):
        return self.cur_pos < self.length

    def next(self, a_w, a_h):
        need_w = 0
        need_h = Constant.FONT_SIZE
        for tmp_pos in range(self.cur_pos, self.length):
            tmp_hz = self.hz_list[tmp_pos]
            need_w = need_w + tmp_hz.width
            if need_h < tmp_hz.height:
                need_h = tmp_hz.height
            if need_w > a_w:
                break
        if need_w > a_w:
            if tmp_hz.is_mark:
                ret_list = self.hz_list[self.cur_pos:tmp_pos-1]
                self.cur_pos = tmp_pos-1
            else:
                ret_list = self.hz_list[self.cur_pos:tmp_pos]
                self.cur_pos = tmp_pos
        else:
            ret_list = self.hz_list[self.cur_pos:]
            self.cur_pos = self.length
        return need_w, need_h, ret_list

    def _parse(self):
        lst = list()
        if self.first_line_space > 0:
            lst.append(HZ(hz=' ', hz_font_size=self.first_line_space))
        sentences = re.split(r'([^\u4e00-\u9fa5])', self.content)
        for sentence in sentences:
            if sentence == '':
                continue
            if re.search(r'([^\u4e00-\u9fa5])', sentence):
                lst.append(HZ(hz=sentence, is_mark=True))
            else:
                py = pinyin(sentence)
                for i in range(0, len(sentence)):
                    if self.recognized_words.get(sentence[i]):
                        lst.append(HZ(hz=sentence[i]))
                    else:
                        lst.append(HZ(hz=sentence[i], py=py[i][0]))
        return lst

    def __str__(self):
        return ''.join([str(tmp) for tmp in self.hz_list])


@lru_cache(maxsize=5)
def get_hz(_hz_number=1500, knows_char=None):
    if _hz_number not in [None, 1500, 3500]:
        raise Exception('only support 1500 or 3500')
    hz = dict()
    if knows_char:
        for s in knows_char:
            hz[s] = s
    if _hz_number:
        filename = get_absolute_path('hz{}.txt'.format(_hz_number))
        with open(filename, 'rb') as f:
            for s in f.read().decode('UTF-8'):
                if '\u4e00' <= s <= '\u9fff':
                    hz[s] = s
    return hz
