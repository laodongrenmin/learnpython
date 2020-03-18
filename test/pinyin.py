#!/usr/bin/env python
# -*- coding: UTF-8 -*-

# https://www.cnblogs.com/y-sh/p/11911329.html
# https://www.cnblogs.com/zhuminghui/p/10944090.html
# https://blog.csdn.net/yangfengling1023/article/details/89472966  python调用各个分词包
# https://mp.weixin.qq.com/s/-iH8QiAbpyOV-692XC5Nzw 分词那些事儿


# d = '''话说天下大势，分久必合，合久必分。周末七国分争，并入于秦。及秦灭之后，楚、汉分争，又并入于汉。名茶自高祖斩白蛇而起义，一统天下，后来光武中兴，传至献帝，遂(suì)分为三国。推其致乱之由，殆(dài)始于桓(huán)、灵二帝。桓帝禁(jìn)锢(gù)善类，崇信宦官。及桓帝崩(bēng)，灵帝即位，大将军窦(dòu)武(wǔ)、太(tài)傅(fù)陈蕃(fān)共相辅(fǔ)佐。时有宦官曹节等弄权，窦武、陈蕃谋诛之，机事不密，反为所害，中涓(juān)自此愈横。'''
# print(pinyin(d))
#
# import jieba
#
# # jieba.add_word('七国分争',tag='d')
# # d =
# jieba.load_userdict('my_dict.txt')
# result = jieba.cut(d, cut_all=False)
# print(" ".join(result))


import re
import chardet
from pypinyin import pinyin, load_phrases_dict
from functools import lru_cache
import time
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.units import inch, mm
from reportlab.lib.colors import pink, black, red, blue, green
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.lib.styles import ParagraphStyle


load_phrases_dict({'铁骑': [['tiě'], ['qqqí']], '差铁骑': [['chāi'], ['tiě'], ['qí']],})

# 注册彩虹粗仿宋以及微软雅黑
def register_font():
    fonts = ['CHCFS.TTF', 'MSYH.TTC', 'STLITI.TTF']
    for f in fonts:
        pdfmetrics.registerFont(TTFont(f.split('.')[0], f))


fh = ['：', '，', '‘', ',', ';', '。', ',', '“', '”']


@lru_cache(maxsize=1)
def get_hz(_hz_number=1500):
    if _hz_number not in [1500, 3500]:
        raise Exception('only support 1500 or 3500')
    hz = dict()
    with open('hz{}.txt'.format(_hz_number), 'rb') as f:
        for s in f.read().decode('UTF-8'):
            if '\u4e00' <= s <= '\u9fff':
                hz[s] = s
    return hz


class Constant(object):
    FONT_NAME = 'MSYH'
    FONT_SIZE = 20
    PY_FONT_NAME = 'MSYH'
    PY_FONT_SIZE = 8
    LINE_SPACE = FONT_SIZE * 0.6


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
        if self.py:
            self.py_width = pdfmetrics.stringWidth(self.py, self.py_font_name, self.py_font_size)
            self.width = max(self.py_width, self.hz_font_size)
        else:
            self.py_width = 0
            self.width = self.hz_font_size
        self.height = self.hz_font_size + self.py_font_size + self.hz_py_spacing if self.py else 0

    def draw(self, _canvas: canvas.Canvas, x,  y, align='center'):
        _canvas.setFont(self.hz_font_name, self.hz_font_size)
        _canvas.drawString(x + (self.width - self.hz_font_size)/2, y, self.hz)
        if hz.py:
            _canvas.setFont(self.py_font_name, self.py_font_size)
            _canvas.drawString(x + (self.width - self.py_width)/2, y + self.hz_font_size + self.hz_py_spacing, self.py)

    def get_size(self):
        return self.width, self.height

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


def get_story_hz(filename, hz_1500=get_hz()):
    with open(filename, 'rb') as f:
        buf = f.read()
    content = buf.decode(chardet.detect(buf).get('encoding', 'GBK'))
    ps = re.split(r'(\r\n)+', content)
    story = []
    for text in ps:
        if text == '\r\n':
            continue
        story.append(MyParagraph(text, hz_1500))
        # break
        # print(text)
        # print(tmp)
        # break
    return story

from reportlab.platypus import Paragraph, Frame


if __name__ == '__main__':
    register_font()
    pagesize=297*mm, 210*mm
    c = canvas.Canvas('story.pdf', pagesize=pagesize)

    margin = (30*mm, 50*mm, 30*mm, 10*mm)


    story = get_story_hz('文本.txt', get_hz())
    x, y = margin[0], margin[1]
    w, h = pagesize[0] - x - margin[2], pagesize[1] - y - margin[3]
    pos = (x, h)
    y = pos[1]
    c.line(x, h, w, h)

    for s in story:
        while s.has_next():
            n_w, n_h, hz_list = s.next(w, h)    # list [] 为空就表示没有了
            y = y - n_h - Constant.LINE_SPACE
            x = pos[0]
            if y < margin[3]:
                y = pos[1] - n_h - Constant.LINE_SPACE
                c.showPage()
                c.line(x, h, w, h)
            for hz in hz_list:
                hz.draw(c, x, y)
                x = x + hz.width
                # print('{},({},{}),{}'.format(hz.hz, hz.width, hz.height, hz.py))
        # break
    c.showPage()
    c.save()
