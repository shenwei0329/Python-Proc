#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   研发管理MIS系统：根据“工时”计算考核分
#   =======================================
#   2020.1.2 @Chengdu
#
#

import datetime
import numpy as np
import mongodb_class

"""全员"""
Personals = {}

Level0 = (1, 0.85)
Level1 = (0.84, 0.6)
Level2 = (0.59, 0.4)
Level3 = (0.39, 0.2)


def calw(group, per_v, lvl):

    _w = []
    # 合并同分项
    _vg = []
    for _g in group:
        _v = int(_g[1]/100)*100
        if _v not in _vg:
            _vg.append(_v)
    # print(_vg)
    _d = lvl[0] - lvl[1]
    if len(_vg) > 1:
        _dv = _d / (len(_vg)-1)
    else:
        _dv = 0
    _vv = []
    _w_v = lvl[0]
    for __ in _vg:
        _vv.append(int((_w_v + 0.005) * 100.)/100.)
        _w_v -= _dv
    # print(_vv)
    for __vg, __w in zip(_vg, _vv):
        # print(">>>{} {}".format(per_v, __vg))
        if per_v >= __vg:
            return __w
    return 0


def calDay(bg, ed, fmt):
    # print("> {} {} {}".format(bg, ed, fmt))
    _bg = datetime.datetime.strptime(bg, fmt)
    _ed = datetime.datetime.strptime(ed, fmt)
    if str(_bg).split(" ")[0] == str(_ed).split(" ")[0]:
        return 1
    # print("{}-{}={}".format(_ed, _bg, _ed-_bg))
    return int(str(_ed-_bg).split(' ')[0])


def main():
    global Personals

    with open('per.txt', 'r', encoding='utf-8') as f:
        for _line in f.readlines():
            _line = _line.replace("\n", "").replace("\r", "")
            print("[{}]".format(_line))
            Personals[_line] = 0.

    db = mongodb_class.mongoDB()

    # WORK_LOGS
    db.connect_db("WORK_LOGS")
    _sql = {'$and': [
        {'updated': {'$gte': "2019-01-01"}},
        {'updated': {'$lte': "2019-12-31"}},
    ]}
    _cur = db.handler("worklog", "find", _sql)
    for _rec in _cur:
        if _rec["author"] in Personals:
            Personals[_rec["author"]] += float(_rec["timeSpentSeconds"])

    # DevOpesTask
    db.connect_db("ext_system")
    _sql = {}
    _cur = db.handler("devops_task", "find", _sql)
    for _rec in _cur:
        _per = _rec["执行者"].split("(")[0]
        if "已完成" not in _rec['任务状态']:
            continue
        if "2019" not in _rec['开始时间']:
            continue
        if _per in Personals:
            # print(">>> DevOpesTask: {}".format(_per))
            bg = _rec['开始时间'].split(" ")[0]
            ed = _rec['完成时间'].split(" ")[0]
            if "年" in bg:
                _n = calDay(bg, ed, "%Y年%m月%d日")
            elif "-" in bg:
                _n = calDay(bg, ed, "%Y-%m-%d")
            else:
                _n = calDay(bg, ed, "%Y/%m/%d")
            Personals[_per] += float(8*3600*_n)

    # OpsTask
    db.connect_db("ext_system")
    _sql = {}
    _cur = db.handler("ops_task", "find", _sql)
    for _rec in _cur:
        if "执行人" not in _rec:
            continue
        _per = _rec["执行人"]
        if _per in Personals:
            Personals[_per] += float(8*3600)

    # TowerTask
    db.connect_db("ext_system")
    _sql = {}
    _pp = {}
    _cur = db.handler("tower", "find", _sql)
    for _rec in _cur:
        if "负责人" not in _rec:
            continue
        _per = _rec["负责人"]
        if _per not in _pp:
            _pp[_per] = []
        if _rec["完成时间"] not in _pp[_per]:
            _pp[_per].append(_rec["完成时间"])
    for _p in _pp:
        if _p in Personals:
            Personals[_per] += float(8*3600*len(_pp[_p]))

    # StarTask
    db.connect_db("ext_system")
    _sql = {}
    _cur = db.handler("star_task", "find", _sql)
    for _rec in _cur:
        if "责任人" not in _rec:
            continue
        _per = _rec["责任人"]
        if "完成" not in _rec['任务状态']:
            continue
        if "2019" not in _rec['完成时间']:
            continue
        if _per in Personals:
            bg = _rec['开始时间'].split(" ")[0]
            ed = _rec['完成时间'].split(" ")[0]
            if "年" in bg:
                _n = calDay(bg, ed, "%Y年%m月%d日")
            elif "-" in bg:
                _n = calDay(bg, ed, "%Y-%m-%d")
            else:
                _n = calDay(bg, ed, "%Y/%m/%d")
            Personals[_per] += float(8*3600*_n)

    _max = 0
    _val = []
    for _p in sorted(Personals, key=lambda x: Personals[x], reverse=True):
        if Personals[_p] > _max:
            _max = Personals[_p]
            Personals[_p] = Personals[_p]/2
        print("> [{}]\t{}\t".format(_p, Personals[_p]), end=" ")
        print("=" * int((80 * Personals[_p])/_max))
        _val.append(Personals[_p])

    x = np.array(_val)
    _mean = int(x.mean())
    _std = int(x.std())
    print(">>>\t Mean: {}\tStd: {}".format(_mean, int(_std)))
    _high = _mean + _std
    _low = _mean - _std

    l0 = []
    l1 = []
    l2 = []
    l3 = []
    for _p in sorted(Personals, key=lambda x: Personals[x], reverse=True):
        if Personals[_p] >= _high:
            l0.append((_p, Personals[_p]))
        elif Personals[_p] >= _mean:
            l1.append((_p, Personals[_p]))
        elif Personals[_p] >= _low:
            l2.append((_p, Personals[_p]))
        else:
            l3.append((_p, Personals[_p]))

    # L0
    for _p in l0:
        _w = calw(l0, _p[1], Level0)
        print("{}\t优\t{}".format(_p[0], _w))

    # L1
    for _p in l1:
        _w = calw(l1, _p[1], Level1)
        print("{}\t良\t{}".format(_p[0], _w))

    # L2
    for _p in l2:
        _w = calw(l2, _p[1], Level2)
        print("{}\t中\t{}".format(_p[0], _w))

    # L3
    for _p in l3:
        _w = calw(l3, _p[1], Level3)
        print("{}\t差\t{}".format(_p[0], _w))


if __name__ == '__main__':
    main()
