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
import mongodb_class
from docx.enum.text import WD_ALIGN_PARAGRAPH
import DataHandler.crWord
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


def build_sql( bg_date, ed_date ):
    _sql = {'$and': [
        {'project_alias': {'$ne': None}},
        {'updated': {'$gte': bg_date}},
        {'updated': {'$lt': ed_date}},
        {'spent_time': {'$gt': 0}}]}
    return _sql


def main():

    global doc

    if len(sys.argv) != 4:
        print("\tUsage: %s DB bg_date ed_date\n" % sys.argv[0])
        return

    _lvl = 1

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'《项目上测试资源投入明细》', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】（%s - %s） <<<' % (time.ctime(), sys.argv[2], sys.argv[3]), align=WD_ALIGN_PARAGRAPH.CENTER)

    db = mongodb_class.mongoDB()
    db.connect_db(sys.argv[1])

    _sql = build_sql(sys.argv[2], sys.argv[3])

    _rec = db.handler("issue", "find", _sql)
    print _rec.count()

    _pj = {}
    for _r in _rec:
        if _r['project_alias'] not in _pj:
            _pj[_r['project_alias']] = []
        _pj[_r['project_alias']].append(_r)

    for _p in sorted(_pj):

        _print(u"%d、项目【%s】" % (_lvl, _p), title=True, title_lvl=1)
        _lvl += 1

        _sn = 1
        doc.addTable(1, 5, col_width=(1, 4, 3, 2, 2))
        _title = (('text', u'序号'),
                  ('text', u'任务内容'),
                  ('text', u'执行日期'),
                  ('text', u'工时'),
                  ('text', u'执行人'),
                  )
        doc.addRow(_title)

        _sum = 0.
        for _r in _pj[_p]:
            _v = float(_r['spent_time'])/3600.
            _text = (
                ('text', "%d" % _sn),
                ('text', _r['summary']),
                ('text', _r['updated'].split('.')[0]),
                ('text', "%0.2f" % _v),
                ('text', _r['users']),
            )
            doc.addRow(_text)
            _sn += 1
            _sum += _v

        _text = (
            ('text', '总计'),
            ('text', ''),
            ('text', ''),
            ('text', "%0.2f" % _sum),
            ('text', ''),
        )
        doc.addRow(_text)

    doc.saveFile('task_detail.docx')


if __name__ == '__main__':
    main()
