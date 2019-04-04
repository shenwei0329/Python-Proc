#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：行为分析
#   =========================
#   2019.2.13 @Chengdu
#
#

from __future__ import unicode_literals

try:
    import configparser as configparser
except Exception:
    import ConfigParser as configparser

import os
import sys
from docx.enum.text import WD_ALIGN_PARAGRAPH
import DataHandler.crWord
import DataHandler.doBox
from pylab import mpl

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

import time

mpl.rcParams['font.sans-serif'] = ['SimHei']

doc = None
oui_list = {}
Topic = 1


def _print(_str, title=False, title_lvl=0, color=None, align=None, paragrap=None ):

    global doc, Topic

    _str = u"%s" % _str.replace('\r', '').replace('\n','')

    _paragrap = None

    if title_lvl == 1:
        Topic = 1
    if title:
        if title_lvl == 2:
            _str = "%d、" % Topic + _str
            Topic += 1
        if align is not None:
            _paragrap = doc.addHead(_str, title_lvl, align=align)
        else:
            _paragrap = doc.addHead(_str, title_lvl)
    else:
        if align is not None:
            if paragrap is None:
                _paragrap = doc.addText(_str, color=color, align=align)
            else:
                doc.appendText(paragrap, _str, color=color, align=align)
        else:
            if paragrap is None:
                _paragrap = doc.addText(_str, color=color)
            else:
                doc.appendText(paragrap, _str, color=color)
    print(_str)

    return _paragrap


def get_files_list(path):

    fileList = []

    files = os.listdir(path)
    for f in files:
        _path = path + '\\' + f
        if os.path.isfile(_path):
            fileList.append(f)

    return fileList


def load_oui():
    global oui_list

    f = open('oui.txt')
    while True:
        _l = f.readline()
        if len(_l) > 0:
            if "(hex)" in _l:
                _t = _l.replace('(hex)', '').replace('\t', '').replace('\n', '').replace('\r', '').split('  ')
                oui_list[_t[0]] = _t[1]
        else:
            break


def mac_format(mac):
    global oui_list

    _m = mac.replace(':', '-')
    if _m[0:8] in oui_list:
        return oui_list[_m[0:8]].replace(u'\xa0', u'')
    else:
        return "None"


def main():

    global doc

    if len(sys.argv) != 2:
        print("\tUsage: %s date_of_report\n" % sys.argv[0])
        return

    load_oui()

    _lvl = 1

    _path = "D:\\PythonProc\\ai_home\\wifi_probe_date\\%s" % sys.argv[1]

    _file_list = []
    _err = False
    try:
        _file_list = get_files_list(_path)
    except Exception, e:
        print e
        _err = True
    finally:
        if _err:
            return

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'《WiFi行为分析》报告', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)

    _print(u"%d、关联图" % _lvl, title=True, title_lvl=1)
    _lvl += 1

    _fn = _path + '\\network.png'
    doc.addPic(_fn, sizeof=6)
    doc.addPageBreak()

    for _f in _file_list:

        _mac = _f.split('.')[0]
        _mac = _mac.replace(_mac[2], ':')

        _print(u"%d、%s【%s】" % (_lvl, _mac, mac_format(_mac)), title=True, title_lvl=1)
        _lvl += 1

        _fn = _path + '\\' + _f
        doc.addPic(_fn, sizeof=4.2)

        if (_lvl % 2) == 0:
            doc.addPageBreak()

    doc.saveFile('wifi_analysis_report.docx')


if __name__ == '__main__':
    main()
