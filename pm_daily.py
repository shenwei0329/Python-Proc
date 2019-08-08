# -*- coding: utf-8 -*-

import sys
import docx
import mongodb_class
import uuid

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


def getParam(text):

    _text = text.replace('-', "").replace(' ', "").replace('/', "").split(u"】")
    if len(_text[-1]) == 0:
        _text = _text[: -1]

    _params = []
    for _p in _text:
        if _p.find(u"【") == 0:
            __p = _p.split(u"【")[1:]
        else:
            __p = _p.split(u"【")
        for _i in __p:
            _params.append(_i)

    return _params


def build_id(params):

    _str = "build_id"

    for _v in params:
        _str += str(_v)

    _h = hash(_str)
    return uuid.uuid3(uuid.NAMESPACE_DNS, str(_h))


def file_handler(_file):

    _short_file = _file.split("\\")[-1]
    # print "fileHandler: ", _short_file

    if ('.docx' not in _short_file) and ('.doc' in _short_file):
        print "Invalid file name: ", _short_file
        return
    else:

        mongo_db = mongodb_class.mongoDB('PM_DAILY')

        _heading_lvl = 0
        _step = -1
        _desc = ""
        _way = ""

        _doc = docx.Document(_file)

        _total_target_lvl = 1.1
        _daily = {}
        for para in _doc.paragraphs:

            _params = getParam(para.text)

            if "Title" in para.style.name:
                _daily['title'] = {"alias": _params[0], "project_id": _params[2]}
            else:
                if para.style.name in ["Normal", "List Paragraph"]:

                    if _heading_lvl == 0 and u"日报" in para.text:
                        _text_lvl = 1
                        _daily['title']['date'] = _params[0]
                    elif _heading_lvl == 1:

                        if 'total_target' not in _daily:
                            _daily['total_target'] = []

                        if para.style.name in "List Paragraph":
                            # print ">>> %0.1f <<< %d" % (_total_target_lvl, len(_params))
                            if len(_params) == 3:
                                _daily['total_target'].append({'id': "%0.1f" % _total_target_lvl,
                                                               'summary': _params[0],
                                                               'date': _params[1],
                                                               'percent': _params[2],
                                                               'daily_date': _daily['title']['date']
                                                               })
                            elif len(_params) == 5:
                                _daily['total_target'].append({'id': "%0.1f" % _total_target_lvl,
                                                               'summary': _params[0],
                                                               'requirement': _params[1],
                                                               'method': _params[2],
                                                               'date': _params[3],
                                                               'percent': _params[4],
                                                               'daily_date': _daily['title']['date']
                                                               })
                            elif len(_params) == 4:
                                # print _params
                                _daily['total_target'].append({'id': "%0.1f" % _total_target_lvl,
                                                               'summary': _params[0],
                                                               'requirement': _params[1],
                                                               'method': "",
                                                               'date': _params[2],
                                                               'percent': _params[3],
                                                               'daily_date': _daily['title']['date']
                                                               })
                            else:
                                print(u"文档正文【目标】格式错误！")
                                return
                            _total_target_lvl += 0.1
                        else:
                            if len(_params) == 4:
                                _daily['total_target'].append({'id': _params[0],
                                                               'summary': _params[1],
                                                               'date': _params[2],
                                                               'percent': _params[3],
                                                               'daily_date': _daily['title']['date']
                                                               })
                            elif len(_params) == 6:
                                _daily['total_target'].append({'id': _params[0],
                                                               'summary': _params[1],
                                                               'requirement': _params[2],
                                                               'method': _params[3],
                                                               'date': _params[4],
                                                               'percent': _params[5],
                                                               'daily_date': _daily['title']['date']
                                                               })
                            else:
                                print(u"文档正文【目标】格式错误！")
                                return

                    elif _heading_lvl == 2:

                        if para.style.name in "List Paragraph":
                            print(u"文档正文【阶段目标】格式错误！")
                            return

                        if 'stage_target' not in _daily:
                            _daily['stage_target'] = []
                        if len(_params) == 5:
                            _daily['stage_target'].append({'sub_id': _params[0],
                                                           'id': _params[1],
                                                           'summary': _params[2],
                                                           'date': _params[3],
                                                           'percent': _params[4],
                                                           'daily_date': _daily['title']['date']
                                                           })
                        elif len(_params) == 7:
                            _daily['stage_target'].append({'sub_id': _params[0],
                                                           'id': _params[1],
                                                           'summary': _params[2],
                                                           'requirement': _params[3],
                                                           'method': _params[4],
                                                           'date': _params[5],
                                                           'percent': _params[6],
                                                           'daily_date': _daily['title']['date']
                                                           })
                        elif len(_params) >=6:
                            _daily['stage_target'].append({'sub_id': _params[0],
                                                           'id': _params[1],
                                                           'summary': _params[2],
                                                           'requirement': _params[3],
                                                           'method': _params[4],
                                                           'date': _params[5],
                                                           'daily_date': _daily['title']['date']
                                                           })
                        else:
                            # show_message(hwnd, u"文档正文【阶段目标】格式错误！")
                            print(u"文档正文【阶段目标】参数个数错误！<%d>" % len(_params))

                    elif _heading_lvl == 3:

                        if para.style.name in "List Paragraph":
                            print(u"文档正文【今日工作汇报】格式错误！")
                            return

                        if 'today' not in _daily:
                            _daily['today'] = []
                        if len(_params) >= 5:
                            _daily['today'].append({'sub_id': _params[0],
                                                    'summary': _params[1],
                                                    'date': _params[2],
                                                    'percent': _params[3],
                                                    'member': _params[4],
                                                    'daily_date': _daily['title']['date']
                                                    })
                    elif _heading_lvl == 4:

                        if para.style.name in "List Paragraph":
                            print(u"文档正文【明日工作计划】格式错误！")
                            return

                        if 'tomorrow' not in _daily:
                            _daily['tomorrow'] = []
                        if len(_params) >= 4:
                            _daily['tomorrow'].append({'sub_id': _params[0],
                                                       'summary': _params[1],
                                                       'date': _params[2],
                                                       'member': _params[3],
                                                       'daily_date': _daily['title']['date']
                                                       })
                    elif _heading_lvl == 5:
                        if 'risk' not in _daily:
                            _daily['risk'] = []
                        if len(_params) > 1:
                            if u"描述" in _params[0]:
                                _desc = _params[1].replace(":", "").replace("：", "")
                            elif u"应对" in _params[0]:
                                _way = _params[1].replace(":", "").replace("：", "")
                                if len(_desc) > 0 or len(_way) > 0:
                                    _daily['risk'].append({"index": _step,
                                                           "desc": _desc,
                                                           "way": _way,
                                                           'daily_date': _daily['title']['date']
                                                           })
                    elif _heading_lvl == 6:
                        if 'problem' not in _daily:
                            _daily['problem'] = []
                        if len(_params) > 1:
                            if u"描述" in _params[0]:
                                _desc = _params[1].replace(":", "").replace("：", "")
                            elif u"应对" in _params[0]:
                                _way = _params[1].replace(":", "").replace("：", "")
                                if len(_desc) > 0 or len(_way) > 0:
                                    _daily['problem'].append({"index": _step,
                                                              "desc": _desc,
                                                              "way": _way,
                                                              'daily_date': _daily['title']['date']
                                                              })
                    elif _heading_lvl == 7:
                        if 'other' not in _daily:
                            _daily['other'] = []
                            _step = 0
                        _daily['other'].append(para.text)

                if "Heading 1" in para.style.name:
                    if u"总体目标" in para.text:
                        _heading_lvl = 1
                        """总体目标完成百分比"""
                        _daily['title']['total_percent'] = _params[0]
                    elif u"阶段目标" in para.text:
                        _heading_lvl = 2
                    elif u"今日工作" in para.text:
                        _heading_lvl = 3
                    elif u"明日工作" in para.text:
                        _heading_lvl = 4
                    elif u"风险" in para.text:
                        _heading_lvl = 5
                        _step = -1
                    elif u"问题" in para.text:
                        _heading_lvl = 6
                        _step = -1
                    else:
                        _heading_lvl = 7
                elif "Heading 2" in para.style.name:
                    if _heading_lvl in [5, 6]:
                        _step += 1

        """去重：是否已录入"""
        _t = mongo_db.handler("pm_daily", "find_one", _daily['title'])
        if _t is None:
            """记录项目标题"""
            mongo_db.handler("pm_daily", "insert", _daily['title'])

            """记录总体目标情况"""
            _idx = 1
            for _v in _daily['total_target']:
                _daily['title']['_id'] = build_id([str(_idx),
                                                   "total_target",
                                                   _daily['title']['date'],
                                                   _daily['title']['project_id']])
                try:
                    mongo_db.handler("total_target", "insert", dict(_daily['title'].items() + _v.items()))
                except Exception, e:
                    print e
                _idx += 1

            """记录阶段目标情况"""
            _idx = 1
            for _v in _daily['stage_target']:
                _daily['title']['_id'] = build_id([str(_idx),
                                                   "stage_target",
                                                   _daily['title']['date'],
                                                   _daily['title']['project_id']])
                try:
                    mongo_db.handler("stage_target", "insert", dict(_daily['title'].items() + _v.items()))
                except Exception, e:
                    print e
                _idx += 1

            """记录当天任务执行情况"""
            _idx = 1
            for _v in _daily['today']:
                _daily['title']['_id'] = build_id([str(_idx),
                                                   "today",
                                                   _daily['title']['date'],
                                                   _daily['title']['project_id']])
                try:
                    mongo_db.handler("today_task", "insert", dict(_daily['title'].items() + _v.items()))
                except Exception, e:
                    print e
                _idx += 1

            """记录明天计划"""
            _idx = 1
            if 'tomorrow' in _daily:
                for _v in _daily['tomorrow']:
                    _daily['title']['_id'] = build_id([str(_idx),
                                                       "tomorrow",
                                                       _daily['title']['date'],
                                                       _daily['title']['project_id']])
                    try:
                        mongo_db.handler("tomorrow_plan", "insert", dict(_daily['title'].items() + _v.items()))
                    except Exception, e:
                        print e
                    _idx += 1

            """记录风险信息"""
            if 'risk' in _daily:
                for _v in _daily['risk']:
                    _daily['title']['_id'] = build_id([str(_v['index']),
                                                       _daily['title']['date'],
                                                       _daily['title']['project_id']])
                    try:
                        mongo_db.handler("risk", "insert", dict(_daily['title'].items() + _v.items()))
                    except Exception, e:
                        print e

            """记录问题信息"""
            if 'problem' in _daily:
                for _v in _daily['problem']:
                    _daily['title']['_id'] = build_id([str(_v['index']),
                                                       _daily['title']['date'],
                                                       _daily['title']['project_id']])
                    try:
                        mongo_db.handler("problem", "insert", dict(_daily['title'].items() + _v.items()))
                    except Exception, e:
                        print e
        else:
            print(u"该文档【%s】已导入系统！" % _short_file)


