__doc__ = """为三国演义、西游记等生成小说PDF"""
from utils.pinyin import get_pinyin_file
from utils import get_absolute_path, get_text_file_content, guess_encoding, get_hz_from_ttf, get_problem_hz
import time
import re
import os


def san_guo_yan_yi():
    if os.path.exists(get_absolute_path("三国演义")):
        print('已经存在，不需要重新生成！')
        return
    file_content = get_text_file_content(get_absolute_path("三国演义.txt"))
    ps = re.split(r'(\r\n)+', file_content)
    new_content = []
    for p in ps:
        if p == '\r\n':
            new_content.append(p)
        else:
            has_shige = False
            sentence_len = 0
            index = 0
            # 寻找诗歌，重新排一下版
            sentences = re.split(r'([^\u4e00-\u9fa5])', p)
            for sentence in sentences:
                sentence = sentence.lstrip('\u3000| ')
                if has_shige and sentence == '”':
                    new_content.append(sentence)
                    new_content.append('\r\n')
                    has_shige = False
                    sentence_len = 0
                    index = 0
                    continue
                if sentence == '':
                    continue
                if re.search(r'([^\u4e00-\u9fa5])', sentence):
                    new_content.append(sentence)
                else:
                    if True or sentence == '更怜一种伤心处':
                        print(sentence)
                    if has_shige:
                        if sentence_len == 0:
                            sentence_len = len(sentence)
                        if sentence_len == len(sentence):
                            if index % 2 == 0:
                                new_content.append('\r\n          ')
                            new_content.append('{}'.format(sentence))
                            index = index + 1
                            continue
                        else:
                            has_shige = False
                            sentence_len = 0
                    new_content.append(sentence)
                    if sentence.find('后人有诗') != -1:
                        has_shige = True
                        sentence_len = 0
                        index = 0
                        continue

    with open(get_absolute_path('三国演义'), 'wb') as f:
        f.write(''.join(new_content).encode('UTF-8'))


def xi_you_ji():
    if os.path.exists(get_absolute_path("西游记")):
        print('已经存在，不需要重新生成！')
        return
    file_content = get_text_file_content(get_absolute_path("西游记.txt"))

    ps = re.split(r'(\r\n)+', file_content)
    new_content = []
    for p in ps:
        if p == '\r\n':
            new_content.append(p)
        else:
            if p.find('吴承恩--西游记') == -1:
                new_content.append(p)

    with open(get_absolute_path('西游记'), 'wb') as f:
        f.write(''.join(new_content).encode('UTF-8'))


if __name__ == '__main__':
    #  三国演义
    print("三国演义")
    t = time.time()
    san_guo_yan_yi()
    t = time.time() - t
    print('处理诗句，cost time: {}'.format(t))
    knows_char = get_text_file_content(get_absolute_path('三国演义认识的字'))
    t = time.time()
    new = get_pinyin_file(get_absolute_path("三国演义"), get_absolute_path("out_pinyin/三国演义_py.text"),
                          used_char=3500, knows_char=knows_char)
    t = time.time() - t
    print('导出文本文件，cost time: {}'.format(t))

    t = time.time()
    new = get_pinyin_file(get_absolute_path("三国演义"), get_absolute_path("out_pinyin/三国演义.pdf"),
                          used_char=3500, knows_char=knows_char)
    t = time.time() - t
    print('cost time: {}'.format(t))

    #  西游记
    print('西游记')
    t = time.time()
    xi_you_ji()
    t = time.time() - t
    print('处理诗句，cost time: {}'.format(t))
    knows_char = get_text_file_content(get_absolute_path('西游记认识的字'))
    t = time.time()
    new = get_pinyin_file(get_absolute_path("西游记"), get_absolute_path("out_pinyin/西游记_py.text"),
                          used_char=3500, knows_char=knows_char)
    t = time.time() - t
    print('导出文本文件，cost time: {}'.format(t))

    t = time.time()
    new = get_pinyin_file(get_absolute_path("西游记"), get_absolute_path("out_pinyin/西游记.pdf"),
                          used_char=3500, knows_char=knows_char)
    t = time.time() - t
    print('cost time: {}'.format(t))
