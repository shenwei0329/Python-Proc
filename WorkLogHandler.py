#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：工作日志处理
#   =============================
#   2019.4.12 @Chengdu
#
#

from __future__ import unicode_literals

import DataHandler.doBox
import mongodb_class
import sys
import handler
import numpy as np
from operator import itemgetter

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

"""关键词"""
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

db = mongodb_class.mongoDB()


def build_sql(field, bg_date, ed_date):

    _bg = bg_date.replace('/', '-')
    _ed = ed_date.replace('/', '-')

    if len(_bg) == 8 and '-' not in _bg:
        _bg = "%s-%s-%s" % (_bg[0:4], _bg[4:6], _bg[6:])
    if len(_ed) == 8 and '-' not in _ed:
        _ed = "%s-%s-%s" % (_ed[0:4], _ed[4:6], _ed[6:])

    _sql = {'$and': [
        {field: {'$gte': _bg}},
        {field: {'$lte': _ed}},
    ]}
    return _sql


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


def behavior_analysis(personal, bg_date, ed_date, filter=None, mday=22):
    """
    个人工作行为分析
    :param personal: 个人
    :param bg_date: 起始日期
    :param ed_date: 截止日期
    :param filter: 项目过滤器，如"SCGA-"
    :param mday：法定工作天数
    :return: _list，任务内容；_text，行为数据说明； _row，行为特征
    """

    _sql = build_sql("started", bg_date, ed_date)
    _sql['author'] = personal

    db.connect_db("WORK_LOGS")
    _cur = db.handler("worklog", "find", _sql)

    _log = []
    """构建行为结构
    - comment：内容
    - object：工作对象
    - active：
    - depth：
    - issues：任务数组
    - issue：任务
    """
    _behavior = {'comment': '', 'object': {}, 'active': {}, 'depth': {},
                 'subject_pd': {}, 'subject_pj': {}, 'issues': [], 'issue': 0}
    _obj_sum = 0
    _act_sum = 0

    for _v in _cur:

        _job = _v['issue'].split('-')[0]
        _behavior['comment'] += _v['comment']. \
            replace('"', ''). \
            replace("'", ''). \
            replace(' ', ''). \
            replace('\n', ''). \
            replace('\r', '').upper()

        _behavior['issues'].append(_v['issue'])

        if _job in handler.pd_list:
            if _job not in _behavior['subject_pd']:
                _behavior['subject_pd'][_job] = 0
            _behavior['subject_pd'][_job] += 1
        else:
            if _job not in _behavior['subject_pj']:
                _behavior['subject_pj'][_job] = 0
            _behavior['subject_pj'][_job] += 1
        _behavior['issue'] += 1

        if (filter is not None) and (filter not in _v['issue']):
            continue
        _log.append(_v)

    _list = []
    _task = {}
    for _t in sorted(_log):

        _date = _t['started'].split('T')[0]

        if _date not in _task:
            _task[_date] = {}

        if _t['issue'] not in _task[_date]:
            _task[_date][_t['issue']] = []
        _task[_date][_t['issue']].append({"cost": _t['timeSpentSeconds'],
                                          "comment": _t['comment'].replace('\n', "").replace('\r', "")})

    _sum = 0.
    _count = 0
    for _t in sorted(_task):

        _day = 0.
        _list.append((
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

                if len(_v['comment']) > 40:
                    _list.append((
                        ('text', ''),
                        ('text', _tt),
                        ('text', "%0.2f" % _vv),
                        ('text', _v['comment'][:40]+"...")
                    ))
                else:
                    _list.append((
                        ('text', ''),
                        ('text', _tt),
                        ('text', "%0.2f" % _vv),
                        ('text', _v['comment'])
                    ))

        _count += 1
        _list.append((
            ('text', ''),
            ('text', u'小计'),
            ('text', "%0.2f" % _day),
            ('text', ''),
        ))

    _list.append((
        ('text', u'总计'),
        ('text', u"%d个工作日" % _count),
        ('text', "投入%0.2f个工时" % _sum),
        ('text', ''),
        ('text', u'法定工作日占比为%0.2f%%' % (float(_count*100)/float(mday))),
        ('text', u'日均工作 %0.2f 小时' % (_sum/_count)),
    ))

    for _i in _behavior['issues']:
        __sql = {'issue': _i}
        db.connect_db(_i.split('-')[0])
        _cur = db.handler("issue", "find_one", __sql)
        if _cur is not None:
            _behavior['comment'] += _cur['summary']

    _obj_sum = 0
    _act_sum = 0
    for _obj in sorted(key_object):
        if _obj not in _behavior['object']:
            _behavior['object'][_obj] = 0
        _behavior['object'][_obj] += count(_behavior['comment'], _obj)
        _obj_sum += _behavior['object'][_obj]
    for _act in sorted(key_active):
        if _act not in _behavior['active']:
            _behavior['active'][_act] = 0
        _behavior['active'][_act] += count(_behavior['comment'], _act)
        _act_sum += _behavior['active'][_act]
    for _dep in sorted(key_depth):
        if _dep not in _behavior['depth']:
            _behavior['depth'][_dep] = 0
        _behavior['depth'][_dep] += count(_behavior['comment'], _dep)

    _text = [u"1）工作日志记录总数：%d 个，信息量：%d 字" % (_behavior['issue'], len(_behavior['comment'])),
             u"2）有效主题和行为各 %d、%d 个" % (_obj_sum, _act_sum)]

    _q = float(np.mean([_obj_sum, _act_sum]))/float(_behavior['issue'])
    _q_str = u"弱"
    if _q >= 1.0:
        _q_str = u"有效"

    _text.append(u"3）日志记录的质量：%s（指标为%0.2f）" % (_q_str, _q))

    _issue_sum = 0
    _labels = []
    _data = []
    _pd_items = []
    for _pd in sorted(handler.pd_list):
        _labels.append(_pd)
        if _pd in _behavior['subject_pd']:
            _issue_sum += _behavior['subject_pd'][_pd]
            _data.append(_behavior['subject_pd'][_pd])
            _pd_items.append((_pd, _behavior['subject_pd'][_pd]))
        else:
            _data.append(0)
            _pd_items.append((_pd, 0))

    _fn_pd = DataHandler.doBox.radar_chart(u"【产品】工作范围", _labels, _data)

    _labels = []
    _data = []
    _pj_items = []
    for _pj in sorted(handler.pj_list):
        _labels.append(_pj)
        if _pj in _behavior['subject_pj']:
            _issue_sum += _behavior['subject_pj'][_pj]
            _data.append(_behavior['subject_pj'][_pj])
            _pj_items.append((_pj, _behavior['subject_pj'][_pj]))
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

    _row = [(('pic', _fn_pd, 1.6), ('pic', _fn_pj, 1.6), ('text', u'%s\n%s' % (_str, __str)))]

    _labels = []
    _data = []
    _sum = {}
    _obj_items = []
    for _obj in sorted(object_class):
        _labels.append(_obj)
        if _obj not in _sum:
            _sum[_obj] = 0
        for _o in object_class[_obj]:
            _sum[_obj] += _behavior['object'][_o]

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
            _sum[_obj] += _behavior['active'][_o]
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

    _row.append((('pic', _fn_object, 1.6), ('pic', _fn_active, 1.6), ('text', u'%s\n%s' % (_str, __str))))

    return _list, _text, _row


pj_list = [u"嘉定",
            u"嘉兴",
            u"甘孜",
            u"安徽",
            u"湖北",
            u"云南",
            u"福田",
            u"产品研发",
            u"公安",
            u"葫芦岛",
            u"新机场",
            u"警综",
            u"国信",
            u"指挥"]


def behavior_analysis_by_work_log(log):
    """
    根据“工作日志”内容进行个人行为分析
    :param log: 日志，{'date': 日期, 'project': 项目, 'summary': 内容}
    :return: _list，任务内容；_text，行为数据说明； _row，行为特征
    """

    _list = []

    _behavior = {'comment': '', 'object': {}, 'active': {}, 'depth': {},
                 'project': {}, 'issues': [], 'issue': 0}

    for _t in sorted(log, key=lambda x: x['date']):
        print _t['project']
        if _t['project'] not in _behavior['project']:
            _behavior['project'][_t['project']] = 0

        _behavior['project'][_t['project']] += 1
        _behavior['issue'] += 1
        _behavior['comment'] += _t['summary']

    _obj_sum = 0
    _act_sum = 0
    for _obj in sorted(key_object):
        if _obj not in _behavior['object']:
            _behavior['object'][_obj] = 0
        _behavior['object'][_obj] += count(_behavior['comment'], _obj)
        _obj_sum += _behavior['object'][_obj]
    for _act in sorted(key_active):
        if _act not in _behavior['active']:
            _behavior['active'][_act] = 0
        _behavior['active'][_act] += count(_behavior['comment'], _act)
        _act_sum += _behavior['active'][_act]
    for _dep in sorted(key_depth):
        if _dep not in _behavior['depth']:
            _behavior['depth'][_dep] = 0
        _behavior['depth'][_dep] += count(_behavior['comment'], _dep)

    _text = [u"1）工作日志记录总数：%d 个，信息量：%d 字" % (_behavior['issue'], len(_behavior['comment'])),
             u"2）有效主题和行为各 %d、%d 个" % (_obj_sum, _act_sum)]

    _q = float(np.mean([_obj_sum, _act_sum]))/float(_behavior['issue'])
    _q_str = u"弱"
    if _q >= 1.0:
        _q_str = u"有效"

    _text.append(u"3）日志记录的质量：%s（指标为%0.2f）" % (_q_str, _q))

    _labels = []
    _data = []
    _pj_items = []
    for _pj in sorted(pj_list):
        _labels.append(_pj)
        if _pj in _behavior['project']:
            _data.append(_behavior['project'][_pj])
            _pj_items.append((_pj, _behavior['project'][_pj]))
        else:
            _data.append(0)
            _pj_items.append((_pj, 0))

    _fn_pj = DataHandler.doBox.radar_chart(u"【项目】工作范围", _labels, _data)

    _v = sorted(_pj_items, key=itemgetter(1), reverse=True)
    __str = ""
    if len(_v) > 0:
        if _v[0][1] > 0:
            __str = u"• 工作以%s为主" % _v[0][0]
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

    _row = [(('pic', _fn_pj, 1.6), ('text', u'%s' % __str))]

    _labels = []
    _data = []
    _sum = {}
    _obj_items = []
    for _obj in sorted(object_class):
        _labels.append(_obj)
        if _obj not in _sum:
            _sum[_obj] = 0
        for _o in object_class[_obj]:
            _sum[_obj] += _behavior['object'][_o]

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
            _sum[_obj] += _behavior['active'][_o]
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

    _row.append((('pic', _fn_object, 1.6), ('pic', _fn_active, 1.6), ('text', u'%s\n%s' % (_str, __str))))

    return _list, _text, _row

