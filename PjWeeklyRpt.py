#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：项目周报
#   =========================
#   2019.3.13 @Chengdu
#
#

from __future__ import unicode_literals

try:
    import configparser as configparser
except Exception:
    import ConfigParser as configparser

import sys
import mongodb_class
from docx.enum.text import WD_ALIGN_PARAGRAPH
import DataHandler.crWord
from pylab import mpl
import time

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

mpl.rcParams['font.sans-serif'] = ['SimHei']

doc = None
oui_list = {}
Topic = 1

sp_word = [
    u"公安厅",
    u"公安局",
    u"检察院",
    u"保密局",
    u"委办厅",
    u"办公厅",
    u"党政军",
    u"政法委",
    u"法院",
    u"省检",
    u"高检",
    u"网安",
    u"总队",
    u"院长",
    u"省长",
    u"市长",
    u"州长",
    u"部队",
    u"厅长",
    u"局长",
    u"主任",
    u"司长",
    u"省委",
    u"市委",
    u"装备",
    u"军队",
    u"党政",
    u"军",
    u"政",
    u"委",
    u"厅",
    u"局",
    u"所",
    u"院",
]


def shield_word(s):
    _s = s
    for _w in sp_word:
        _s = _s.replace(_w, 'xx')
    return _s


def _print(_str, title=False, title_lvl=0, color=None, align=None, paragrap=None):

    global doc, Topic

    _str = u"%s" % _str.replace('\r', '').replace('\n', '')
    _str = shield_word(_str)

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
                _paragrap = doc.appendText(paragrap, _str, color=color, align=align)
        else:
            if paragrap is None:
                _paragrap = doc.addText(_str, color=color)
            else:
                _paragrap = doc.appendText(paragrap, _str, color=color)
    print(_str)

    return _paragrap


def build_sql(field, bg_date, ed_date):
    _sql = {'$and': [
        {field: {'$gte': bg_date.replace('-', '').replace('/', '')}},
        {field: {'$lte': ed_date.replace('-', '').replace('/', '')}}
    ]}
    return _sql


def main():

    global doc

    if len(sys.argv) != 3:
        print("\tUsage: %s bg_date ed_date\n" % sys.argv[0])
        return

    _lvl = 1

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'《项目周报》', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)
    _print('周报期间【%s-%s】' % (sys.argv[1], sys.argv[2]), align=WD_ALIGN_PARAGRAPH.CENTER)

    db = mongodb_class.mongoDB()
    db.connect_db("PM_DAILY")

    _sql = build_sql("date", sys.argv[1], sys.argv[2])

    _rec = db.handler("pm_daily", "find", _sql)
    print _rec.count()

    _print(u"%d、本周项目工时统计" % _lvl, title=True, title_lvl=1)
    _lvl += 1
    _print(u"根据本周项目日报统计的各项目资源投入工时如下：")

    _pj = {}
    for _r in _rec:
        if _r['project_id'] not in _pj:
            _pj[_r['project_id']] = {}
            _pj[_r['project_id']]['alias'] = _r['alias']
        if 'date' not in _pj[_r['project_id']]:
            _pj[_r['project_id']]['date'] = []
        _pj[_r['project_id']]['date'].append(_r['date'])

    _rec = db.handler("total_target", "find", {})
    for _r in _rec:
        if _r['project_id'] not in _pj:
            print(">>> Error: total_target[%s] not in pm_daily" % _r['project_id'])
            continue
        if 'id' not in _pj[_r['project_id']]:
            _pj[_r['project_id']]['id'] = {}

        _r['id'] = _r['id']
        if _r['id'] not in _pj[_r['project_id']]['id']:
            _pj[_r['project_id']]['id'][_r['id']] = {
                'summary': _r['summary'],
                'date': _r['date'],
                'percent': []
            }

        if '%' in _r['percent']:
            _pct = _r['percent'].split('%')[0]
        else:
            _pct = '0'
        _pj[_r['project_id']]['id'][_r['id']]['percent'].append(_pct)

    _work_hour = {}
    _pg_list = None
    for _p in sorted(_pj):

        _pg = _print(u"%d、%s【%s】" % (_lvl, _p, _pj[_p]['alias']), title=True, title_lvl=1)
        if _pg_list is None:
            _pg_list = _pg
        _lvl += 1

        _rec = db.handler("stage_target", "find", {'project_id': _p})
        _stage = {}
        for _r in _rec:
            if _r['id'] not in _stage:
                _stage[_r['id']] = {}
            if _r['sub_id'] not in _stage[_r['id']]:
                _stage[_r['id']][_r['sub_id']] = {'percent': []}

            if '%' in _r['percent']:
                _pct = (_r['percent'].split('%')[0]).split('.')[0]
            else:
                _pct = '0'
            _stage[_r['id']][_r['sub_id']]['percent'].append(_pct)
            _stage[_r['id']][_r['sub_id']]['summary'] = _r['summary']
            _stage[_r['id']][_r['sub_id']]['date'] = _r['date']

        _print(u"一、目标完成情况", title=True, title_lvl=3)

        for _id in sorted(_pj[_p]['id'], key=lambda x: int(x.split('.')[1])):
            _pct = sorted(_pj[_p]['id'][_id]['percent'], key=lambda x: int(x))

            _print(u"\t%s）%s" % (_id, _pj[_p]['id'][_id]['summary']), title=True, title_lvl=3)

            if _pct[0] != _pct[-1]:
                _print(u"计划完成时间：%s，本周完成率从 %s%% 变为 %s%% 。" % (
                    _pj[_p]['id'][_id]['date'],
                    _pct[0],
                    _pct[-1]),
                       color=(0, 100, 0)
                )
            else:
                if int(_pct[0]) == 100:
                    _print(u"计划完成时间：%s，目标已完成。" % (
                        _pj[_p]['id'][_id]['date'])
                           )
                else:
                    _print(u"计划完成时间：%s，本周没有进展，完成率仍保持在%s%% 。" % (
                        _pj[_p]['id'][_id]['date'],
                        _pct[0]),
                           color=(255, 0, 0)
                           )

            if _id not in _stage:
                continue

            if len(_stage) == 0:
                continue

            _print(u"包含的阶段目标有：")
            for _sub in sorted(_stage[_id], key=lambda x: int(x.split('.')[1])):
                _pct = sorted(_stage[_id][_sub]['percent'], key=lambda x: int(x))
                if _pct[0] != _pct[-1]:
                    _print(u"\t● 阶段目标：%s，计划完成实现：%s，本周完成率从 %s%% 变为 %s%% 。" % (
                        _stage[_id][_sub]['summary'],
                        _stage[_id][_sub]['date'],
                        _pct[0],
                        _pct[-1]),
                           color=(0, 10, 0)
                           )
                else:
                    if int(_pct[0]) == 100:
                        _print(u"\t● 阶段目标：%s，计划完成实现：%s，目标已完成。" % (
                            _stage[_id][_sub]['summary'],
                            _stage[_id][_sub]['date'])
                        )
                    else:
                        _print(u"\t● 阶段目标：%s，计划完成实现：%s，本周没有进展，完成率仍保持在%s%% 。" % (
                            _stage[_id][_sub]['summary'],
                            _stage[_id][_sub]['date'],
                            _pct[0]),
                            color=(255, 0, 0)
                               )

        _print(u"二、任务完成情况", title=True, title_lvl=3)

        _pg = _print(u"任务明细如下：")
        _sql = build_sql("daily_date", sys.argv[1], sys.argv[2])
        _sql['project_id'] = _p
        _rec = db.handler("today_task", "find", _sql)
        _task = {}
        for _r in _rec:
            if _r['daily_date'] not in _task:
                _task[_r['daily_date']] = {}
            if _r['sub_id'] not in _task[_r['daily_date']]:
                _task[_r['daily_date']][_r['sub_id']] = []
            _task[_r['daily_date']][_r['sub_id']].append({
                'summary': _r['summary'],
                'date': _r['date'],
                'member': _r['member'],
                'percent': _r['percent']
            })

        doc.addTable(1, 5, col_width=(1, 4, 1, 1, 2))
        _title = (
                  ('text', u'日期'),
                  ('text', u'任务内容'),
                  ('text', u'计划'),
                  ('text', u'进度'),
                  ('text', u'执行人'),
                  )
        doc.addRow(_title)

        _cnt = 0

        for _t in sorted(_task):
            _text = (
                ('text', _t),
                ('text', ''),
                ('text', ''),
                ('text', ''),
                ('text', ''),
            )
            doc.addRow(_text)
            _personal = []
            for _sub in sorted(_task[_t], key=lambda x: int(x.split('.')[1])):
                for _data in _task[_t][_sub]:
                    """排除外协类、领导和资源缺失？
                    """
                    if ((u"外协" or u"总" or u"缺失") not in _data['member']) and (_data['member'] not in _personal):
                        _personal.append(_data['member'])
                    _text = (
                        ('text', ''),
                        ('text', _data['summary']),
                        ('text', _data['date']),
                        ('text', _data['percent']),
                        ('text', _data['member']),
                    )
                    doc.addRow(_text)

            """资源投入（人日）就是当天参与任务执行的人员个数"""
            _cnt += len(_personal)

        _print(u"本周工作量：%d（人天）" % _cnt, paragrap=_pg)
        _work_hour[_p] = _cnt

        _print(u"三、本周风险", title=True, title_lvl=3)

        _rec = db.handler("risk", "find", {'project_id': _p})
        if _rec.count() == 0:
            _print(u"无。")
        else:
            _risk = {}
            for _r in _rec:
                if _r["desc"] not in _risk:
                    _risk[_r["desc"]] = _r['way']

            doc.addTable(1, 2, col_width=(3, 4))
            _title = (
                      ('text', u'风险'),
                      ('text', u'解决办法'),
                      )
            doc.addRow(_title)

            for _r in _risk:
                _text = (
                    ('text', _r),
                    ('text', _risk[_r])
                )
                doc.addRow(_text)

        _print(u"四、本周问题", title=True, title_lvl=3)
        _rec = db.handler("problem", "find", {'project_id': _p})
        if _rec.count() == 0:
            _print(u"无。")
        else:
            _problem = {}
            for _r in _rec:
                if _r["desc"] not in _problem:
                    _problem[_r["desc"]] = _r['way']

            doc.addTable(1, 2, col_width=(3, 4))
            _title = (
                      ('text', u'问题'),
                      ('text', u'解决办法'),
                      )
            doc.addRow(_title)

            for _r in _problem:
                _text = (
                    ('text', _r),
                    ('text', _problem[_r])
                )
                doc.addRow(_text)

        doc.addPageBreak()

    for _w in _work_hour:
        _print(u"\t● %-20s：%d 人天" % (_w, _work_hour[_w]), paragrap=_pg_list)

    doc.saveFile('weekly_rpt.docx')


if __name__ == '__main__':
    main()
