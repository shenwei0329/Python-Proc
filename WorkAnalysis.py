#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：工作分析
#   =========================
#   2019.4.12 @Chengdu
#
#

from __future__ import unicode_literals

import handler

try:
    import configparser as configparser
except Exception:
    import ConfigParser as configparser

import numpy as np
from operator import itemgetter
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
from datetime import datetime

mpl.rcParams['font.sans-serif'] = ['SimHei']

doc = None

"""全员"""
Personals = {}
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


def count(src, word):

    _count = 0
    _idx = 0
    _len = len(word)
    _src = src
    while True:
        _idx = _src.find(word)
        if _idx < 0:
            break
        _count += 1
        _idx += _len
        _src = _src[_idx:]

    return _count


def write_title(bgdate,eddate):

    _print("I、简介", title=True, title_lvl=1)

    _print(u"本报告是根据每位员工个人在%s至%s期间的工作日志进行数据分析得出的，有关员工个人工作的行为内容说明。"
           u"基于此分析，可较为准确地了解员工的日常工作行为，进一步了解员工的工作效率。" % (bgdate, eddate))
    _print(u"个人工作的行为包括：")
    _print(u"\t1、任务：该员工日常工作的任务；")
    _print(u"\t2、过程：该员工针对此任务日常的工作内容；")
    _print(u"\t3、工时：该员工针对此任务的工作内容所花费的工时。")


def build_sql(field, bg_date, ed_date):

    _bg = bg_date.replace('/', '-')
    _ed = ed_date.replace('/', '-')

    if len(_bg) == 8 and '-' not in _bg:
        _bg = "%s-%s-%s" % (_bg[0:4], _bg[4:6], _bg[6:])
    if len(_ed) == 8 and '-' not in _ed:
        _ed = "%s-%s-%s" % (_ed[0:4], _ed[4:6], _ed[6:])

    _sql = {'$and': [
        {field: {'$gte': _bg}},
        {field: {'$lte': _ed}}
    ]}
    return _sql, _bg, _ed


def main():
    global Personals, key_object, key_active, key_depth, doc

    if len(sys.argv) != 3:
        print("\tUsage: %s bg_date ed_date\n" % sys.argv[0])
        return

    _sql, _bgdate, _eddate = build_sql("started", sys.argv[1], sys.argv[2])
    print _sql
    _cnt = 0

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'《个人工作行为说明》报告', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)

    write_title(_bgdate, _eddate)

    import mongodb_class
    db = mongodb_class.mongoDB()

    for _job in ['WORK_LOGS']:

        db.connect_db(_job)
        _cur = db.handler("worklog", "find", _sql)

        for _v in _cur:

            if _v.has_key('author') and ('SCGA-' in _v['issue']):

                if _v['author'] not in Personals:
                    Personals[_v['author']] = []
                Personals[_v['author']].append(_v)

    _print(u"I、个人工作行为", title=True, title_lvl=1)
    _print(u"本节提供了全体员工的工作行为情况，包括日常工作任务、内容和耗时。")

    for _p in sorted(Personals):
        _task = {}
        for _t in Personals[_p]:

            _date = _t['started'].split('T')[0]

            if _date not in _task:
                _task[_date] = {}

            if _t['issue'] not in _task[_date]:
                _task[_date][_t['issue']] = []
            _task[_date][_t['issue']].append({"cost": _t['timeSpentSeconds'], "comment": _t['comment']})

        doc.addPageBreak()
        _print(_p, title=True, title_lvl=2)

        doc.addTable(1, 4, col_width=(1, 1, 1, 3))
        _title = (('text', u'日期'),
                  ('text', u'任务'),
                  ('text', u'工时'),
                  ('text', u'内容'),
                  )
        doc.addRow(_title)

        print(u"name: %s" % _p)

        _sum = 0.
        _count = 0
        for _t in sorted(_task):

            _day = 0.
            print(u"date: %s" % _t)
            doc.addRow((
                ('text', _t),
                ('text', ''),
                ('text', ''),
                ('text', ''),
            ))

            for _work in _task[_t]:
                _bg = True
                for _v in _task[_t][_work]:
                    _vv = float(_v['cost'])/3600.
                    _sum += _vv
                    _day += _vv
                    if _bg:
                        _tt = _work
                        _bg = False
                    else:
                        _tt = ""

                    doc.addRow((
                        ('text', ''),
                        ('text', _tt),
                        ('text', "%0.2f" % _vv),
                        ('text', _v['comment'])
                    ))
            _count += 1
            doc.addRow((
                ('text', u'小计'),
                ('text', ''),
                ('text', "%0.2f" % _day),
                ('text', ''),
            ))

        doc.addRow((
            ('text', u'总计'),
            ('text', "%d" % _count),
            ('text', "%0.2f" % _sum),
            ('text', ''),
        ))

    doc.saveFile('work_analysis_report.docx')


if __name__ == '__main__':
    main()
