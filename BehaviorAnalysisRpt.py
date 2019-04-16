#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：行为分析
#   =========================
#   2019.2.13 @Chengdu
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
import getopt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import DataHandler.crWord
import DataHandler.doBox
from pylab import mpl
import WorkLogHandler

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

import time

mpl.rcParams['font.sans-serif'] = ['SimHei']

doc = None

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

    _print(u"本报告是根据每位员工个人在%s至%s期间的工作日志进行数据分析得出的，有关员工个人工作的行为特征。"
           u"基于此分析，可较为准确地了解员工的日常工作行为，进一步了解员工的工作特质。" % (bgdate, eddate))
    _print(u"个人工作的行为特征包括：")
    _print(u"\t1、范围分布：该员工从事的工作范围分布情况，用以了解其工作方向偏向产品和项目；")
    _print(u"\t2、主题分布：该员工日常工作的对象（名词）分布情况，用以了解其日常主要从事哪些工作内容；")
    _print(u"\t3、行为分布：该员工日常工作的行为（动词）分布情况，用以了解其日常主要有哪些工作行为表现。")


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


def main(filter=None, mx=False, mday=22):
    global doc

    _sql, _bgdate, _eddate = build_sql('started', sys.argv[-2], sys.argv[-1])

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'《个人行为特征分析》报告', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)

    write_title(_bgdate, _eddate)

    import mongodb_class
    db = mongodb_class.mongoDB()

    Personals = []
    for _job in ['WORK_LOGS']:

        db.connect_db(_job)
        _cur = db.handler("worklog", "find", _sql)

        for _v in _cur:

            if _v.has_key('author'):

                if (filter is not None) and (filter not in _v['issue']):
                    continue

                _job = _v['issue'].split('-')[0]
                if _v['author'] not in Personals:
                    Personals.append(_v['author'])

    _print(u"I、个人工作行为特征", title=True, title_lvl=1)
    _print(u"本节针对每位员工，给出：")
    _print(u"1）工作日志情况：包含日志记录个数，日志信息量（字数），日志中包含有效关键字的情况，"
           u"并对日志质量进行评价。")
    _print(u"2）工作范围；包含员工在产品研发和项目开发方面的工作投入情况。")
    _print(u"3）行为特征；包含员工在日常工作的重点方向和主要工作行为。")

    for _p in sorted(Personals):

        doc.addPageBreak()
        _print(_p, title=True, title_lvl=2)

        _list, _text, _row = WorkLogHandler.behavior_analysis(_p, _bgdate, _eddate, filter=filter, mday=mday)

        for _v in _text:
            _print(_v)

        _print(u"一、工作范围")
        doc.addTable(1, 3, col_width=(1, 1, 2))
        _title = (('text', u'产品范围'),
                  ('text', u'项目范围'),
                  ('text', u'特征'))
        doc.addRow(_title)
        doc.addRow(_row[0])

        _print(u"二、行为特征")
        doc.addTable(1, 3, col_width=(1.6, 1.6, 2))
        _title = (('text', u'主题分布'),
                  ('text', u'行为分布'),
                  ('text', u'特征'))
        doc.addRow(_title)
        doc.addRow(_row[1])

        if mx:
            _print(u"三、工作明细")
            doc.addTable(1, 4, col_width=(1, 1, 1, 3))
            _title = (('text', u'日期'),
                      ('text', u'任务'),
                      ('text', u'工时'),
                      ('text', u'内容'),
                      )
            doc.addRow(_title)
            for _v in _list:
                doc.addRow(_v)
        else:
            _print(u"三、工作特性")
            _print(u"%s执行%s，%s；%s，%s" % (
                _list[-1][0][1],
                _list[-1][1][1],
                _list[-1][4][1],
                _list[-1][2][1],
                _list[-1][5][1]))

    doc.saveFile('behavior_analysis_report.docx')


if __name__ == '__main__':

    _cmd_line = '\tUsage: %s [-h] [-d] [-m <M>] [-f <Filter>] be-date ed-date\n'

    if len(sys.argv)<3:
        print _cmd_line % sys.argv[0]
        sys.exit(1)

    _has_mx = False
    _filter = None
    _m = 22

    try:
         opts, args = getopt.getopt(sys.argv[1:-2], "hdm:f:", [])
    except getopt.GetoptError:
        print sys.argv[1:-2]
        print "Error: \n", _cmd_line % sys.argv[0]
        sys.exit(2)

    for opt, arg in opts:

        if opt == '-h':
            print _cmd_line % sys.argv[0]
            sys.exit()
        elif opt == '-d':
            _has_mx = True
        elif opt == '-m':
            _m = arg
        elif opt == '-f':
            _filter = arg

    main(filter=_filter, mx=_has_mx, mday=_m)
