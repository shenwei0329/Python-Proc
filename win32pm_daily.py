# -*- coding: utf-8 -*-

import win32gui
import win32con
import win32api
import sys
from win32con import *
import docx
from DataHandler import mongodb_class
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


def fileHandler(_file):

    _short_file = _file.split("\\")[-1]
    print "fileHandler: ", _short_file

    if (('.docx' not in _short_file) and ('.doc' in _short_file)):
        print "Invalid file name: ", _short_file
        return
    else:

        mongo_db = mongodb_class.mongoDB('PM_DAILY')

        _heading_lvl = 0
        _step = -1
        _desc = ""
        _way = ""

        _doc = docx.Document(_file)

        _daily = {}
        for para in _doc.paragraphs:

            _params = getParam(para.text)
            # print _params

            if "Title" in para.style.name:
                _daily['title'] = {"alias": _params[0], "project_id": _params[2]}
            else:
                if "Normal" in para.style.name:
                    if _heading_lvl == 0 and u"日报" in para.text:
                        _text_lvl = 1
                        _daily['title']['date'] = _params[0]
                    elif _heading_lvl == 1:
                        if 'total_target' not in _daily:
                            _daily['total_target'] = []
                        if len(_params) >= 4:
                            _daily['total_target'].append({'id': _params[0],
                                                           'summary': _params[1],
                                                           'date': _params[2],
                                                           'percent': _params[3],
                                                           'daily_date': _daily['title']['date']
                                                           })
                    elif _heading_lvl == 2:
                        if 'stage_target' not in _daily:
                            _daily['stage_target'] = []
                        if len(_params) >= 5:
                            _daily['stage_target'].append({'sub_id': _params[0],
                                                           'id': _params[1],
                                                           'summary': _params[2],
                                                           'date': _params[3],
                                                           'percent': _params[4],
                                                           'daily_date': _daily['title']['date']
                                                           })
                    elif _heading_lvl == 3:
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
                                    _daily['risk'].append({"index": _step, "desc": _desc, "way": _way})
                    elif _heading_lvl == 6:
                        if 'problem' not in _daily:
                            _daily['problem'] = []
                        if len(_params) > 1:
                            if u"描述" in _params[0]:
                                _desc = _params[1].replace(":", "").replace("：", "")
                            elif u"应对" in _params[0]:
                                _way = _params[1].replace(":", "").replace("：", "")
                                if len(_desc) > 0 or len(_way) > 0:
                                    _daily['problem'].append({"index": _step, "desc": _desc, "way": _way})
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


def WndProc(hwnd, msg, wParam, lParam):
    """
    视窗处理
    :param hwnd: 句柄
    :param msg: 消息
    :param wParam: 参数1
    :param lParam: 参数2
    :return: 下一个消息处理器
    """

    if msg == WM_PAINT:
        hdc,ps = win32gui.BeginPaint(hwnd)
        rect = win32gui.GetClientRect(hwnd)
        _str = u'请把【项目经理：日报】文件拖拽到这！'
        win32gui.DrawText(hdc, _str, len(_str)*2-4, rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER)
        win32gui.EndPaint(hwnd, ps)
        win32gui.DragAcceptFiles(hwnd, 1)
    elif msg == WM_DESTROY:
        win32gui.PostQuitMessage(0)
    elif msg == WM_DROPFILES:
        """文件拖拉消息
        """
        hDropInfo = wParam
        filescount = win32api.DragQueryFile(hDropInfo)
        for i in range(filescount):
            """获取文件名"""
            filename = win32api.DragQueryFile(hDropInfo, i)
            """处理文件"""
            fileHandler(filename)
        win32api.DragFinish(hDropInfo)

    return win32gui.DefWindowProc(hwnd,msg,wParam,lParam)  


wc = win32gui.WNDCLASS()
wc.hbrBackground = COLOR_BTNFACE + 1
wc.hCursor = win32gui.LoadCursor(0,IDI_APPLICATION)
wc.lpszClassName = "RDMinfoSystem no Windows"
wc.lpfnWndProc = WndProc
reg = win32gui.RegisterClass(wc)
hwnd = win32gui.CreateWindow(reg,
                             u'研发管理信息系统：外部数据采集器',
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
                             CW_USEDEFAULT,
                             CW_USEDEFAULT,
                             CW_USEDEFAULT,
                             CW_USEDEFAULT,
                             0,
                             0,
                             0,
                             None)
win32gui.SetWindowPos(hwnd,
                      win32con.HWND_TOPMOST,
                      0, 0, 320, 120,
                      win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE | win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
win32gui.ShowWindow(hwnd, SW_SHOWNORMAL)
win32gui.UpdateWindow(hwnd)  
win32gui.PumpMessages()
