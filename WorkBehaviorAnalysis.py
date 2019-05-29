#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：工作行为（个人工作日志）分析
#   =============================================
#   2019.5.22 @Chengdu
#
#   获取每个员工的“工作日志”（含日期、工作内容）
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
from datetime import datetime, date, timedelta
import time

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

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


def cvDate2Chn(date):
    """
    字符日期转换成中文表示
    :param date: "yyyy-mm-dd"
    :return: 中文日期
    """

    _d = date.split(' ')[0].split('-')
    return u"%s年%d月%d日" % (_d[0], int(_d[1]), int(_d[2]))


def build_sql(table, member):
    """
    生成查询条件
    :param table: 表，因各个团队采用了自己的日志收集工具，使得日志字段不同
    :param member: 人员名称
    :return: 查询SQL
    """

    """采用“匹配”查询方式，适合对包含多人字段的查询"""
    if 'worklog' in table:
        _sql = {'author': {'$regex': ".*%s.*" % member}}
    elif 'devops_task' in table:
        _sql = {u"执行者": {'$regex': ".*%s.*" % member}}
    elif 'ops_task' in table:
        _sql = {u"执行人": {'$regex': ".*%s.*" % member}}
    elif 'ops_task_bj' in table:
        _sql = {"$or": [{u"故障处理人": {'$regex': ".*%s.*" % member}}, {u"技术-处理人": {'$regex': ".*%s.*" % member}}]}
    elif 'star_task' in table:
        _sql = {u"责任人": {'$regex': ".*%s.*" % member}}
    else:
        return ""

    return _sql


def inRegion(c_date, bg_date, end_date):
    """
    判断指定的时间是否落入由开始时间和截止时间定义的范围内（注：含这两个时间）。
    :param c_date: 指定的时间
    :param bg_date: 开始时间
    :param end_date: 截止时间
    :return: True，表示落入；False，表示超出
    """

    _date = time.mktime(time.strptime(c_date, '%Y-%m-%d'))
    _bg_date = time.mktime(time.strptime(bg_date, '%Y-%m-%d'))
    _end_date = time.mktime(time.strptime(end_date, '%Y-%m-%d'))

    if (_date >= _bg_date) and (_date <= _end_date):
        return True
    return False


def beforeDate(c_date, o_date):
    """
    判断c_date是否在o_date日期之前。
    :param c_date: 时间1
    :param o_date: 时间2
    :return: True，表示时间1早于时间2；False，表示时间1迟于时间2
    """
    _c_date = time.mktime(time.strptime(c_date, '%Y-%m-%d'))
    _o_date = time.mktime(time.strptime(o_date, '%Y-%m-%d'))
    return _c_date <= _o_date


def afterDate(c_date, o_date):
    """
    判断c_date是否在o_date日期之后。
    :param c_date: 时间1
    :param o_date: 时间2
    :return: True，表示时间1迟于时间2；False，表示时间1早于时间2
    """
    if '-' not in o_date:
        return False
    if '-' not in c_date:
        return False
    _c_date = time.mktime(time.strptime(c_date, '%Y-%m-%d'))
    _o_date = time.mktime(time.strptime(o_date, '%Y-%m-%d'))
    return _c_date > _o_date


def formatDate(c_date):
    """
    规整日期格式，例如：2019-5-30，2019-05-30
    :param c_date:
    :return:
    """
    _ti = time.strptime(c_date, '%Y-%m-%d')
    return time.strftime('%Y-%m-%d', _ti)


def inDateRegion(table, rec, bg_date, end_date):
    """
    判断记录是否在规定的日期内，包含4种情况：1）跨越；2）完成；3）开始；4）开始并完成
    :param table: 表名，根据不同的表取数据
    :param rec: 记录集
    :param bg_date: 起始日期，格式：yyyy-mm-dd HH:MM:SS
    :param end_date: 截止日期，格式：yyyy-mm-dd HH:MM:SS
    :return: 范围内的记录集，重构：人员、日期、项目、工作内容
    """

    _rec = []
    for _r in rec:

        if 'worklog' in table:
            """由【started】字段选择
            """
            _date = _r['started'].split('T')[0]
            if inRegion(_date, bg_date, end_date):
                print _date
                _rec.append(
                    {'date': formatDate(_date),
                     'project': _r['project'],
                     'summary': _r['comment']
                     })

        elif 'devops_task' in table:
            """由【完成时间】字段选择
            """
            _date = _r[u'完成时间'].replace(u"年", '-').replace(u"月", '-').replace(u"日", '')
            if len(_date)>0 and inRegion(_date, bg_date, end_date):
                print _date
                _rec.append(
                    {'date': formatDate(_date),
                     'project': "devops",
                     'summary': _r[u"标题"]
                     })

        elif 'ops_task' in table:
            """由【日期】字段选择
            """
            _date = _r[u'日期'].encode("utf8").replace(u'年', '-').replace(u'月', '-').replace(u'日', '')
            if inRegion(_date, bg_date, end_date):
                print _date
                _rec.append(
                    {'date': formatDate(_date),
                     'project': u"公安运维",
                     'summary': _r['任务'].encode("utf8")
                     })

        elif 'ops_task_bj' in table:
            """由【创建于】字段选择
            """
            _date = _r[u'创建于']
            if inRegion(_date, bg_date, end_date):
                print _date
                if u"故障" in _r[u"跟踪"]:
                    _rec.append(
                        {'date': formatDate(_date),
                         'project': u"北京运维",
                         'summary': _r[u'任务']
                         })
                else:
                    _rec.append(
                        {'date': formatDate(_date),
                         'project': u"北京运维",
                         'summary': _r[u'主题']
                         })

        elif 'star_task' in table:
            """由【开始时间】、【截止时间】、【完成时间】和【状态】字段选择
            """
            if u"已完成" in _r[u"任务状态"]:
                _date = _r[u"完成时间"].split(' ')[0]
                if inRegion(_date, bg_date, end_date):
                    """在指定的时间范围内完成的"""
                    print _date
                    _rec.append(
                        {'date': formatDate(_date),
                         'project': _r[u"分类"],
                         'summary': (_r[u'任务描述'] + _r[u'任务详细描述']).replace('\n', '').replace('\r', '')
                         })
            else:
                _date = _r[u"开始时间"].split(' ')[0]
                _date_ed = _r[u"截止时间"].split(' ')[0]
                # if beforeDate(_date, bg_date) or afterDate(_date_ed, end_date):
                """在指定的时间范围内已开始，且未完成的"""
                if afterDate(_date_ed, end_date):
                    """针对“未完成”的任务，仅考虑当前未到期的"""
                    print _date
                    _rec.append(
                        {'date': formatDate(_date),
                         'project': _r[u"分类"],
                         'summary': (_r[u'任务描述'] + _r[u'任务详细描述']).replace('\n', '').replace('\r', '')
                         })
        else:
            """无效"""
            continue
    return _rec


def usage():
    print """
usage: %(prog)s [options] members

examples:
  %(prog)s -s 2019-04-01 -e 2019-04-02 shenwei
      - to report work-log of shenwei the duty from 2019-04-01 to 2019-04-02.

options:
-s, --started: the date for begin 
-e, --ended: the date for end
-h, --help: this help message

""" % {'prog': os.path.basename(__file__)}


def loadMembers():
    """
    导入花名册
    :return: 人员列表
    """
    _members = []
    _f = open('org_member.txt', "rb")
    while True:
        _s = _f.readline()
        _s = _s.decode("gbk").encode("utf8")
        _s = _s.replace('\n', '').replace('\r', '')
        if _s is None:
            return _members
        if len(_s) == 0:
            return _members
        _members.append(_s)


def main():
    global doc

    """解析命令行
    """
    import getopt
    try:
        opts, args = getopt.getopt(sys.argv[1:], "h:s:e:", ["help", "started=", "ended="])
    except getopt.GetoptError, err:
        print str(err)
        usage()
        sys.exit(2)

    yesterday = date.today() + timedelta(days=-1)
    _bg_time = yesterday.strftime("%Y-%m-%d")
    _now = str(datetime.now()).split(' ')[0]

    for o, a in opts:
        if o in ("-h", "--help"):
            usage()
            sys.exit()
        elif o in ("-s", "--started"):
            _bg_time = a
        elif o in ("-e", "--ended"):
            _now = a

    _lvl = 1

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'个人《工作日志》报告', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print(u'>>> 生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)
    _print(u'时间段：（%s至%s）' % (cvDate2Chn(_bg_time), cvDate2Chn(_now)), align=WD_ALIGN_PARAGRAPH.CENTER)

    data_sources = {"WORK_LOGS": ["worklog"],
                    "ext_system": ["devops_task",
                                   "ops_task",
                                   "ops_task_bj",
                                   "star_task"]
                    }

    _member = {}

    if len(args) == 0:
        _members = loadMembers()
        print _members
    else:
        _members = []
        for _p in args:
            _members.append(_p.decode("gbk").encode("utf8"))

    _lvl = 1
    for _p in _members:

        db = mongodb_class.mongoDB()
        for _db in data_sources:

            db.connect_db(_db)

            for _table in data_sources[_db]:

                print _db, _table

                # _sql = build_sql(_table, db.cvGbk2Utf8(_p))
                _sql = build_sql(_table, _p)
                """查看sql内容
                for _q in _sql:
                    print(u"<%s>:[%s]" % (_q, _sql[_q]))
                """
                if len(_sql) == 0:
                    continue

                _rec = db.handler(_table, "find", _sql)
                # print _rec.count()
                """时间窗口"""
                _rec = inDateRegion(_table, _rec, _bg_time, _now)

                if _p not in _member:
                    _member[_p] = []

                for _r in _rec:
                    _member[_p].append(_r)

            db.close_db()

    _project_alias = {
        u"嘉定": u"嘉定",
        u"嘉兴": u"嘉兴",
        u"甘孜": u"甘孜",
        u"安徽": u"安徽",
        u"湖北": u"湖北",
        u"云南": u"云南",
        u"福田": u"福田",
        "FAST": u"产品研发",
        "HUBBLE": u"产品研发",
        u"公安": u"公安",
        "HLD": u"葫芦岛",
        u"新机场": u"新机场",
        u"警综": u"警综",
        u"国信": u"国信",
        u"指挥项目": u"指挥",
        u"指挥系统": u"指挥",
        "YN": u"云南",
        "GZ": u"甘孜"
    }

    _project = {}
    _no_project = []
    _in_project = []
    for _key in _member:
        for _r in _member[_key]:
            for _pj in _project_alias:
                if (_pj in _r['project'].upper()) or (_pj in _r['summary'].upper()):
                    _pj = _project_alias[_pj]
                    _r['project'] = _pj
                    if _pj not in _project:
                        _project[_pj] = []
                    if _key not in _project[_pj]:
                        _project[_pj].append(_key)

                    if _key not in _in_project:
                        _in_project.append(_key)

    for _key in _member:
        if _key not in _in_project:
            if _key not in _no_project:
                _no_project.append(_key)

    _print(u"资源投入情况", title=True, title_lvl=1)

    _pg = _print(u"相关项目有 %d 个，每个项目资源投入（参与的相关人员）情况如下表：" % len(_project))
    doc.addTable(1, 3, col_width=(1, 1, 3))
    _title = (('text', u'项目'),
              ('text', u"人数"),
              ('text', u'人员投入')
              )
    doc.addRow(_title)
    for _pj in _project:
        _mem = ""
        _cnt = 0
        for _m in _project[_pj]:
            _mem += (_m + '，')
            _cnt += 1
        _text = (
            ('text', u"%s" % _pj),
            ('text', "%d" % _cnt),
            ('text', u"%s" % _mem[:-1])
        )
        doc.addRow(_text)

    if len(_no_project) > 0:
        _str = u"没有参与以上项目（或未提交工作日志的）人员有："
        for _p in _no_project:
            _str += u"%s、" % _p
        _print((_str[:-1]+u"。"))

    doc.addPageBreak()

    _used_cnt = 0
    _n_cnt = 0
    for _key in sorted(_member):

        _print(u"%d、%s" % (_lvl, _key), title=True, title_lvl=1)
        _lvl += 1

        if len(_member[_key]) == 0:
            _print(u"无工作内容。")
            _n_cnt += 1
            print(u"%s" % _key)
        else:
            import WorkLogHandler

            _list, _text, _row = WorkLogHandler.behavior_analysis_by_work_log(_member[_key])

            for _v in _text:
                _print(_v)

            _print(u"一、工作范围")
            doc.addTable(1, 2, col_width=(1, 1))
            _title = (('text', u'工作范围'),
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

            _print(u"三、工作明细")
            doc.addTable(1, 3, col_width=(0.5, 2, 4))
            _title = (('text', u'日期'),
                      ('text', u'项目'),
                      ('text', u'任务内容')
                      )
            doc.addRow(_title)

            for _r in sorted(_member[_key], key=lambda x: x['date']):
                if len(_r['summary']) > 40:
                    _text = (
                        ('text', _r['date']),
                        ('text', u"%s" % _r['project']),
                        ('text', (u"%s" % _r['summary'])[:40] + '...')
                    )
                else:
                    _text = (
                        ('text', _r['date']),
                        ('text', u"%s" % _r['project']),
                        ('text', u"%s" % _r['summary'])
                    )
                doc.addRow(_text)
            _used_cnt += 1
        doc.addPageBreak()

    print(">>> Total: %d, Used: %d, Not: %d" % ((_used_cnt + _n_cnt), _used_cnt, _n_cnt))
    _str = u'本时间段内，应提交个人工作日志的总人数为 %d 人；其中，提交个人工作日志的有 %d 人，未提交的有 %d 人。\n' %\
           ((_used_cnt + _n_cnt), _used_cnt, _n_cnt)
    _print(_str, paragrap=_pg)

    doc.saveFile('WorkLogAnalysisRpt.docx')


if __name__ == '__main__':
    main()
