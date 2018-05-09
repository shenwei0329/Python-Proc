#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   报告模板解析器
#   ==============
#   创建：2018.3.2@成都
#
#   模板格式：
#       以一行为基元，内容为：
#           正文<%宏定义%>正文<%宏定义%>正文...
#

import os.path


class report_parse:

    def __init__(self, fn):
        """
        初始化
        :param fn: 模板文件路径
        """
        self.f = None
        if os.path.isfile(fn):
            self.f = open(fn, 'r')

    def begin(self):
        self.f.seek(0, 0)

    def parse(self):
        if self.f is None:
            print("No file!")
            return None
        while True:
            _line = self.f.readline()
            if len(_line) == 0:
                break
            if _line[0] == '#':
                continue
            _text = []
            for _v in _line.split("<%"):
                _text += _v.split("%>")
            return [item.lstrip().rstrip() for item in _text]
        return None


def main():

    parse = report_parse(u"项目跟踪报告【模板】.txt")
    parse.begin()
    while True:
        _str = parse.parse()
        if _str is None:
            break

        """获取到的解析结果：
           1）满足<context><macro><context><macro>... 排列规则
           2）<context>可以为“空”
        """
        _i = 0
        print("-"*8)
        for _s in _str:
            if _i == 0:
                print("context[%s]" % _s),
            else:
                print("macro[%s]" % _s),
            _i = (_i + 1) % 2
        print("")


if __name__ == '__main__':
    main()
