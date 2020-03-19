__doc__ = """utils/pinyin.py 单元测试。"""
import os
import unittest
from utils.pinyin import *
from utils import get_absolute_path, guess_encoding, get_hz_from_ttf, get_problem_hz


class MyTestCase(unittest.TestCase):
    text = '出差盱眙,吃小龙虾或者貔貅,差钱'
    test_start = '测试: pinyin.py 开始'
    test_finish = '测试: pinyin.py 结束'
    out_path = 'out_pinyin'

    def test_00000_get_pinyin_text(self):
        title_print(self.test_start)
        file_path = get_absolute_path(self.out_path)
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        for file in os.listdir(file_path):
            os.remove(os.path.join(file_path,file))

    def test_00001_get_pinyin_text(self):
        title_print('测试多音字')
        ret_text, new = get_pinyin_text(text=self.text)
        print('输入：', self.text)
        print('输出：', ret_text)
        print('注音：', ''.join(new))
        self.assertEqual(ret_text,
                         '出(chū)差(chāi)盱(xū)眙(yí),吃(chī)小(xiǎo)龙(lóng)虾(xiā)或(huò)者(zhě)貔(pí)貅(xiū),差(chà)钱(qián)')

    def test_00002_get_pinyin_text(self):
        title_print('测试1500常用字之外的注音')

        ret_text, new = get_pinyin_text(text=self.text, used_char=1500)
        print('输入：', self.text)
        print('输出：', ret_text)
        print('注音：', ''.join(new))
        self.assertEqual(ret_text,
                         '出差盱(xū)眙(yí),吃小龙虾(xiā)或者貔(pí)貅(xiū),差钱')

    def test_00003_get_pinyin_text(self):
        title_print('测试1500常用字以及自定义认识字之外的注音')
        ret_text, new = get_pinyin_text(text=self.text, used_char=1500, knows_char='虾')
        print('输入：', self.text)
        print('输出：', ret_text)
        print('注音：', ''.join(new))
        self.assertEqual(ret_text,
                         '出差盱(xū)眙(yí),吃小龙虾或者貔(pí)貅(xiū),差钱')

    def test_00004_get_pinyin_file(self):
        title_print('测试3500常用字以及自定义认识字之外的注音 生成文本文件')
        filename = get_absolute_path("文本.txt")
        outfile = os.path.join(self.out_path, '文本_py.txt')
        with open(get_absolute_path("三国演义"), 'rb') as f:
            buf = f.read()
        encoding = guess_encoding(buf)
        knows_char = buf.decode(encoding)
        new = get_pinyin_file(filename, outfile, used_char=3500, knows_char=knows_char)
        print(''.join(new))
        self.assertEqual(len(new), 106)

    def test_00005_get_pinyin_file(self):
        title_print('测试3500常用字以及自定义认识字之外的注音 生成PDF')
        filename = get_absolute_path("文本.txt")
        outfile = os.path.join(self.out_path, '三国演义_部分章节.pdf')
        with open(get_absolute_path("三国演义"), 'rb') as f:
            buf = f.read()
        encoding = guess_encoding(buf)
        knows_char = buf.decode(encoding)
        new = get_pinyin_file(filename, outfile, used_char=3500, knows_char=knows_char)
        print(''.join(new))
        # self.assertEqual(len(new), 105)

    def d_test_88888_make_hz(self):
        import json
        all_hz = []
        for c in range(19968, 40870):    # 19968 - 40869  '\u4e00', '\u9fa5'
            s = '\\u{}'.format(hex(c)[2:])
            s = json.loads(f'"{s}"')
            all_hz.append(s)

        s = ''.join(all_hz)
        with open(get_absolute_path('out_pinyin/all_hz.dat'), 'wb') as f:
            f.write(s.encode('UTF-8'))

        get_pinyin_file(get_absolute_path('out_pinyin/all_hz.dat'),
                        get_absolute_path('out_pinyin/all_hz_chcfs.pdf'), used_char=None, knows_char=s)

    def d_test_900000(self):
        chcfs = get_hz_from_ttf(r'c:\windows\fonts\simsun.ttc')
        for key in chcfs.keys():
            print(key)
        print(len(chcfs))    # 20902  20933  20933

    def d_test_900010(self):
        p_hz = get_problem_hz('chcfs.ttf')
        hh = []
        for k in p_hz.keys():
            hh.append(k)
        s = ''.join(hh)
        with open(get_absolute_path('../utils/chcfs.ttf_p.dat'), 'wb') as f:
            f.write(s.encode('UTF-8'))

    def test_999999(self):
        title_print(self.test_finish)


def title_print(text):
    print('{}{}{}'.format('*'*10, text, '*'*(100-len(text))))


if __name__ == '__main__':
    unittest.main()
