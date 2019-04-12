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

"""关键词
import os

conf = configparser.ConfigParser()
conf.read(os.path.split(os.path.realpath(__file__))[0] + '/keyword.cnf', encoding="utf-8-sig")

key_object = conf.get('key', 'object')
key_active = conf.get('key', 'active')
key_depth = conf.get('key', 'depth')
"""

object_class = {
    u"云平台": [
        u"FAST",
        u"MIR",
        u"PULSAR",
        u"WH",
        u"WHITE",
        u"APOLLO",
    ],
    u"大数据": [
        u"HUBBLE",
    ],
    u"OneX": [
        u"ONE",
    ],
    u"会议": [
        u"会议",
        u"例会",
    ],
    u"界面": [
        u"门户",
        u"界面",
        u"风格",
        u"大屏",
    ],
    u"需求": [
        u"客户",
        u"需求",
    ],
    u"设计&开发": [
        u"方案",
        u"预案",
        u"效果图",
        u"模型",
        u"模块",
        u"流程",
        u"程序",
        u"代码",
        u"接口",
        u"API",
    ],
    u"缺陷&问题": [
        u"BUG",
        u"缺陷",
        u"问题",
    ],
    u"数据处理": [
        u"数据",
    ],
    u"服务&容器": [
        u"服务",
        u"容器",
    ],
    u"任务": [
        u"任务",
    ],
    u"产品": [
        u"产品",
    ],
    u"项目": [
        u"项目",
    ],
}

key_object = [u"产品",
              u"项目",
              u"FAST",
              u"HUBBLE",
              u"MIR",
              u"PULSAR",
              u"WH",
              u"WHITE",
              u"ONE",
              u"APOLLO",
              u"BUG",
              u"模型",
              u"模块",
              u"流程",
              u"程序",
              u"缺陷",
              u"问题",
              u"服务",
              u"容器",
              u"任务",
              u"接口",
              u"代码",
              u"需求",
              u"例会",
              u"会议",
              u"门户",
              u"界面",
              u"风格",
              u"大屏",
              u"客户",
              u"方案",
              u"预案",
              u"数据",
              u"效果图",
              u"API"]

active_class = {
    u"增删改": [
        u"增",
        u"减",
        u"改",
        u"删",
        u"写",
    ],
    u"协同": [
        u"参加",
        u"讨论",
        u"评审",
        u"了解",
        u"获取",
],
    u"改进": [
        u"梳理",
        u"整理",
        u"处理",
        u"调整",
        u"解决",
        u"优化",
    ],
    u"调测": [
        u"联调",
        u"测试",
        u"验证",
        u"排查",
        u"定位",
    ],
    u"管理": [
        u"安排",
        u"分配",
    ],
    u"设计": [
        u"设计",
        u"迭代",
        u"变更",
    ],
    u"辅助": [
        u"制作",
        u"支撑",
        u"支持",
        u"部署",
        u"提交",
        u"迁移",
    ]
}

key_active = [u"参加",
              u"增",
              u"减",
              u"改",
              u"删",
              u"制作",
              u"修",
              u"处理",
              u"调整",
              u"联调",
              u"解决",
              u"测试",
              u"写",
              u"验证",
              u"排查",
              u"定位",
              u"优化",
              u"了解",
              u"梳理",
              u"整理",
              u"安排",
              u"分配",
              u"部署",
              u"提交",
              u"评审",
              u"设计",
              u"支撑",
              u"支持",
              u"变更",
              u"迭代",
              u"迁移",
              u"讨论",
              u"获取"]
key_depth = [u"已",
             u"完成",
             u"结束",
             u"交付",
             u"正在",
             u"过程中"]

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


def write_title(bgdate, eddate):
    """
    写文档简介
    :param bgdate: 起始日期
    :param eddate: 截止日期
    :return:
    """

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


def main():
    global Personals, key_object, key_active, key_depth, doc

    if len(sys.argv) != 3:
        print("\tUsage: %s bg_date ed_date\n" % sys.argv[0])
        return

    _sql, _bgdate, _eddate = build_sql("created", sys.argv[1], sys.argv[2])
    print _sql
    _cnt = 0

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

    for _job in ['WORK_LOGS']:

        db.connect_db(_job)
        _cur = db.handler("worklog", "find", _sql)

        for _v in _cur:
            if _v.has_key('author'):

                _job = _v['issue'].split('-')[0]
                if _v['author'] not in Personals:
                    Personals[_v['author']] = {'comment': '', 'object': {}, 'active': {}, 'depth': {},
                                               'subject_pd': {}, 'subject_pj': {}, 'issues': [], 'issue': 0}
                Personals[_v['author']]['comment'] += _v['comment'].\
                    replace('"', '').\
                    replace("'", '').\
                    replace(' ', '').\
                    replace('\n', '').\
                    replace('\r', '').upper()

                Personals[_v['author']]['issues'].append(_v['issue'])

                if _job in handler.pd_list:
                    if _job not in Personals[_v['author']]['subject_pd']:
                        Personals[_v['author']]['subject_pd'][_job] = 0
                    Personals[_v['author']]['subject_pd'][_job] += 1
                else:
                    if _job not in Personals[_v['author']]['subject_pj']:
                        Personals[_v['author']]['subject_pj'][_job] = 0
                    Personals[_v['author']]['subject_pj'][_job] += 1
                Personals[_v['author']]['issue'] += 1

    """添加issue的summary信息"""
    for _p in Personals:
        for _i in Personals[_p]['issues']:
            __sql = _sql
            __sql['issue'] = _i
            db.connect_db(_i.split('-')[0])
            _cur = db.handler("issue", "find_one", __sql)
            if _cur is not None:
                Personals[_p]['comment'] += _cur['summary']

    doc.addPageBreak()

    _print(u"II、总体工作行为特征", title=True, title_lvl=1)
    _print(u"本节说明全体员工的总体工作行为情况，包括工作日志情况、重点工作投向、工作的聚焦点和工作的行为特点。")

    _log_sum = 0
    _text_sum = 0
    _obj_sum = 0
    _act_sum = 0
    for _p in sorted(Personals):

        _log_sum += Personals[_p]['issue']
        _text_sum += len(Personals[_p]['comment'])

        for _obj in sorted(key_object):
            if _obj not in Personals[_p]['object']:
                Personals[_p]['object'][_obj] = 0
            Personals[_p]['object'][_obj] += count(Personals[_p]['comment'], _obj)
            _obj_sum += Personals[_p]['object'][_obj]
        for _act in sorted(key_active):
            if _act not in Personals[_p]['active']:
                Personals[_p]['active'][_act] = 0
            Personals[_p]['active'][_act] += count(Personals[_p]['comment'], _act)
            _act_sum += Personals[_p]['active'][_act]
        for _dep in sorted(key_depth):
            if _dep not in Personals[_p]['depth']:
                Personals[_p]['depth'][_dep] = 0
            Personals[_p]['depth'][_dep] += count(Personals[_p]['comment'], _dep)

    _print(u"1）参与人员总数：%d 名" % len(Personals))
    _print(u"2）全员工作日志记录总数：%d 个，信息量：%d 字" % (_log_sum, _text_sum))
    _print(u"3）全员有效主题和行为各 %d、%d 个" % (_obj_sum, _act_sum))

    _q = float(np.mean([_obj_sum, _act_sum]))/float(_log_sum)
    _q_str = u"极差"
    if _q >= 1.0:
        _q_str = u"优"
    elif _q >= 0.8:
        _q_str = u"良"
    elif _q >= 0.5:
        _q_str = u"中"
    elif _q >= 0.3:
        _q_str = u"差"

    _print(u"4）日志记录的总体质量：%s（%0.2f）" % (_q_str, _q))

    _print(u"一、全员工作范围")

    doc.addTable(1, 3, col_width=(1, 1, 2))
    _title = (('text', u'产品范围'),
              ('text', u'项目范围'),
              ('text', u'特征'))
    doc.addRow(_title)

    """产品研发投向"""
    _issue_sum = 0
    _labels = []
    _data = []
    _pd_items = []
    _items = {}
    for _p in Personals:
        for _pd in sorted(handler.pd_list):
            if _pd not in _items:
                _items[_pd] = 0
            if _pd in Personals[_p]['subject_pd']:
                _issue_sum += Personals[_p]['subject_pd'][_pd]
                _items[_pd] += Personals[_p]['subject_pd'][_pd]

    for _i in _items:
        _data.append(_items[_i])
        _labels.append(_i)
        _pd_items.append((_i, _items[_i]))

    _fn_pd = DataHandler.doBox.radar_chart(u"【产品】工作范围", _labels, _data)

    """项目开发投向"""
    _labels = []
    _data = []
    _pj_items = []
    _items = {}
    for _p in Personals:
        for _pj in sorted(handler.pj_list):
            if _pj not in _items:
                _items[_pj] = 0
            if _pj in Personals[_p]['subject_pj']:
                _issue_sum += Personals[_p]['subject_pj'][_pj]
                _items[_pj] += Personals[_p]['subject_pj'][_pj]

    for _i in _items:
        _data.append(_items[_i])
        _labels.append(_i)
        _pj_items.append((_i, _items[_i]))

    _fn_pj = DataHandler.doBox.radar_chart(u"【项目】工作范围", _labels, _data)

    _v = sorted(_pd_items, key=itemgetter(1), reverse=True)
    _str = ""
    if len(_v) > 0:
        if _v[0][1] > 0:
            _str = u"• 产品研发的投向以%s为主" % _v[0][0]
            if len(_v) > 1 and _v[1][1] > 0:
                _str += u"，其次分别为"
                _lvl = 0
                for _vv in _v[1:]:
                    if _vv[1] > 0:
                        _str += u"%s，" % _vv[0]
                        _lvl += 1
                        if _lvl > 3:
                            break
                _str = _str[:-1]
            _str += u"。"

    _v = sorted(_pj_items, key=itemgetter(1), reverse=True)
    __str = ""
    if len(_v) > 0:
        if _v[0][1] > 0:
            __str = u"• 项目开发的投向以%s为主" % _v[0][0]
            if len(_v) > 1 and _v[1][1] > 0:
                __str += u"，其次分别为"
                _lvl = 0
                for _vv in _v[1:]:
                    if _vv[1] > 0:
                        __str += u"%s，" % _vv[0]
                        _lvl += 1
                        if _lvl > 3:
                            break
                __str = __str[:-1]
            __str += u"。"

    doc.addRow((('pic', _fn_pd, 1.6),
                ('pic', _fn_pj, 1.6),
                ('text', u'%s\n%s' % (_str, __str))))

    _print(u"二、全员行为特征")

    doc.addTable(1, 3, col_width=(1.6, 1.6, 2))
    _title = (('text', u'主题分布'),
              ('text', u'行为分布'),
              ('text', u'特征'))
    doc.addRow(_title)

    _labels = []
    _data = []
    _sum = {}
    _obj_items = []
    for _obj in sorted(object_class):
        _labels.append(_obj)
        if _obj not in _sum:
            _sum[_obj] = 0
        for _o in object_class[_obj]:
            for _p in Personals:
                _sum[_obj] += Personals[_p]['object'][_o]

        _data.append(_sum[_obj])
        _obj_items.append((_obj, _sum[_obj]))

    _fn_object = DataHandler.doBox.radar_chart(u"工作主题分布", _labels, _data)

    _labels = []
    _data = []
    _act_items = []
    for _obj in sorted(active_class):
        _labels.append(_obj)
        if _obj not in _sum:
            _sum[_obj] = 0
        for _o in active_class[_obj]:
            for _p in Personals:
                _sum[_obj] += Personals[_p]['active'][_o]
        _data.append(_sum[_obj])
        _act_items.append((_obj, _sum[_obj]))

    _fn_active = DataHandler.doBox.radar_chart(u"工作行为分布", _labels, _data)

    _v = sorted(_obj_items, key=itemgetter(1), reverse=True)
    _str = ""
    if len(_v) > 0:
        if _v[0][1] > 0:
            _str = u"• 工作以%s为主" % _v[0][0]
            if len(_v) > 1 and _v[1][1] > 0:
                _str += u"，其次分别为"
                _lvl = 0
                for _vv in _v[1:]:
                    if _vv[1] > 0:
                        _str += u"%s，" % _vv[0]
                        _lvl += 1
                        if _lvl > 3:
                            break
                _str = _str[:-1]
            _str += u"。"

    _v = sorted(_act_items, key=itemgetter(1), reverse=True)
    __str = ""
    if len(_v) > 0:
        if _v[0][1] > 0:
            __str = u"• 工作行为集中在%s" % _v[0][0]
            if len(_v) > 1 and _v[1][1] > 0:
                __str += u"，其次分别为"
                _lvl = 0
                for _vv in _v[1:]:
                    if _vv[1] > 0:
                        __str += u"%s，" % _vv[0]
                        _lvl += 1
                        if _lvl > 3:
                            break
                __str = __str[:-1]
            __str += u"。"

    doc.addRow((('pic', _fn_object, 1.6),
                ('pic', _fn_active, 1.6),
                ('text', u'%s\n%s' % (_str, __str))))

    doc.addPageBreak()

    _print(u"III、个人工作行为特征", title=True, title_lvl=1)
    _print(u"本节针对每位员工，给出：")
    _print(u"1）工作日志情况：包含日志记录个数，日志信息量（字数），日志中包含有效关键字的情况，"
           u"并对日志质量进行评价。")
    _print(u"2）工作范围；包含员工在产品研发和项目开发方面的工作投入情况。")
    _print(u"3）行为特征；包含员工在日常工作的重点方向和主要工作行为。")

    for _p in sorted(Personals):

        doc.addPageBreak()
        _print(_p, title=True, title_lvl=2)

        _obj_sum = 0
        _act_sum = 0
        for _obj in sorted(key_object):
            _obj_sum += Personals[_p]['object'][_obj]
        for _act in sorted(key_active):
            _act_sum += Personals[_p]['active'][_act]

        _print(u"1）工作日志记录总数：%d 个，信息量：%d 字" %
               (Personals[_p]['issue'],
                len(Personals[_p]['comment'])))

        _print(u"2）有效主题和行为各 %d、%d 个" %
               (_obj_sum,
                _act_sum))

        _q = float(np.mean([_obj_sum, _act_sum]))/float(Personals[_p]['issue'])
        _q_str = u"极差"
        if _q >= 1.0:
            _q_str = u"优"
        elif _q >= 0.8:
            _q_str = u"良"
        elif _q >= 0.5:
            _q_str = u"中"
        elif _q >= 0.3:
            _q_str = u"差"

        _print(u"3）日志记录的质量：%s（%0.2f）" % (_q_str, _q))

        _print(u"一、工作范围")

        doc.addTable(1, 3, col_width=(1, 1, 2))
        _title = (('text', u'产品范围'),
                  ('text', u'项目范围'),
                  ('text', u'特征'))
        doc.addRow(_title)

        _issue_sum = 0
        _labels = []
        _data = []
        _pd_items = []
        for _pd in sorted(handler.pd_list):
            _labels.append(_pd)
            if _pd in Personals[_p]['subject_pd']:
                # print _pd, Personals[_p]['subject_pd'][_pd], ";",
                _issue_sum += Personals[_p]['subject_pd'][_pd]
                _data.append(Personals[_p]['subject_pd'][_pd])
                _pd_items.append((_pd, Personals[_p]['subject_pd'][_pd]))
            else:
                _data.append(0)
                _pd_items.append((_pd, 0))

        _fn_pd = DataHandler.doBox.radar_chart(u"【产品】工作范围", _labels, _data)

        _labels = []
        _data = []
        _pj_items = []
        for _pj in sorted(handler.pj_list):
            _labels.append(_pj)
            if _pj in Personals[_p]['subject_pj']:
                # print _pj, Personals[_p]['subject_pj'][_pj], ";",
                _issue_sum += Personals[_p]['subject_pj'][_pj]
                _data.append(Personals[_p]['subject_pj'][_pj])
                _pj_items.append((_pj, Personals[_p]['subject_pj'][_pj]))
            else:
                _data.append(0)
                _pj_items.append((_pj, 0))

        _fn_pj = DataHandler.doBox.radar_chart(u"【项目】工作范围", _labels, _data)

        _v = sorted(_pd_items, key=itemgetter(1), reverse=True)
        _str = ""
        if len(_v) > 0:
            if _v[0][1] > 0:
                _str = u"• 产品研发以%s为主" % _v[0][0]
                if len(_v) > 1 and _v[1][1] > 0:
                    _str += u"，其次分别为"
                    _lvl = 0
                    for _vv in _v[1:]:
                        if _vv[1] > 0:
                            _str += u"%s，" % _vv[0]
                            _lvl += 1
                            if _lvl > 3:
                                break
                    _str = _str[:-1]
                _str += u"。"

        _v = sorted(_pj_items, key=itemgetter(1), reverse=True)
        __str = ""
        if len(_v) > 0:
            if _v[0][1] > 0:
                __str = u"• 项目开发以%s为主" % _v[0][0]
                if len(_v) > 1 and _v[1][1] > 0:
                    __str += u"，其次分别为"
                    _lvl = 0
                    for _vv in _v[1:]:
                        if _vv[1] > 0:
                            __str += u"%s，" % _vv[0]
                            _lvl += 1
                            if _lvl > 3:
                                break
                    __str = __str[:-1]
                __str += u"。"

        doc.addRow((('pic', _fn_pd, 1.6),
                    ('pic', _fn_pj, 1.6),
                    ('text', u'%s\n%s' % (_str, __str))))

        _print(u"二、行为特征")

        doc.addTable(1, 3, col_width=(1.6, 1.6, 2))
        _title = (('text', u'主题分布'),
                  ('text', u'行为分布'),
                  ('text', u'特征'))
        doc.addRow(_title)

        _labels = []
        _data = []
        _sum = {}
        _obj_items = []
        for _obj in sorted(object_class):
            _labels.append(_obj)
            if _obj not in _sum:
                _sum[_obj] = 0
            for _o in object_class[_obj]:
                _sum[_obj] += Personals[_p]['object'][_o]

            _data.append(_sum[_obj])
            _obj_items.append((_obj, _sum[_obj]))

        _fn_object = DataHandler.doBox.radar_chart(u"工作主题分布", _labels, _data)

        _labels = []
        _data = []
        _act_items = []
        for _obj in sorted(active_class):
            _labels.append(_obj)
            if _obj not in _sum:
                _sum[_obj] = 0
            for _o in active_class[_obj]:
                _sum[_obj] += Personals[_p]['active'][_o]
            _data.append(_sum[_obj])
            _act_items.append((_obj, _sum[_obj]))

        _fn_active = DataHandler.doBox.radar_chart(u"工作行为分布", _labels, _data)

        _v = sorted(_obj_items, key=itemgetter(1), reverse=True)
        _str = ""
        if len(_v) > 0:
            if _v[0][1] > 0:
                _str = u"• 工作以%s为主" % _v[0][0]
                if len(_v) > 1 and _v[1][1] > 0:
                    _str += u"，其次分别为"
                    _lvl = 0
                    for _vv in _v[1:]:
                        if _vv[1] > 0:
                            _str += u"%s，" % _vv[0]
                            _lvl += 1
                            if _lvl > 3:
                                break
                    _str = _str[:-1]
                _str += u"。"

        _v = sorted(_act_items, key=itemgetter(1), reverse=True)
        __str = ""
        if len(_v) > 0:
            if _v[0][1] > 0:
                __str = u"• 工作行为集中在%s" % _v[0][0]
                if len(_v) > 1 and _v[1][1] > 0:
                    __str += u"，其次分别为"
                    _lvl = 0
                    for _vv in _v[1:]:
                        if _vv[1] > 0:
                            __str += u"%s，" % _vv[0]
                            _lvl += 1
                            if _lvl > 3:
                                break
                    __str = __str[:-1]
                __str += u"。"

        doc.addRow((('pic', _fn_object, 1.6),
                    ('pic', _fn_active, 1.6),
                    ('text', u'%s\n%s' % (_str, __str))))

    doc.saveFile('behavior_analysis_report.docx')


if __name__ == '__main__':
    main()
