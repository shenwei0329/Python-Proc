# -*- coding: utf-8 -*-

import win32gui
import win32con
import win32api
import sys
from win32con import *
import doXLSX4tables

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

def fileHandler(_file):

    _short_file = _file.split("\\")[-1]
    print "fileHandler: ", _short_file

    if (('.xlsx' not in _short_file) and ('.xls' in _short_file)) and ('~$' not in _short_file):
        print "Invalid file name: ", _short_file
        return
    else:
        if doXLSX4tables.main(_file):
            print(u"任务情况：")
            doXLSX4tables.printInd(doXLSX4tables.task)
            print(u"任务来源情况：")
            doXLSX4tables.printInd(doXLSX4tables.source)
            print(u"计划情况：")
            doXLSX4tables.printInd(doXLSX4tables.plan)
            print(u"成果类型情况：")
            doXLSX4tables.printInd(doXLSX4tables.t_type)
            print(u"成果归属情况：")
            doXLSX4tables.printInd(doXLSX4tables.t_base)


def WndProc(hwnd, msg, wParam, lParam):

    if msg == WM_PAINT:
        hdc,ps = win32gui.BeginPaint(hwnd)
        rect = win32gui.GetClientRect(hwnd)
        _str = u'请把【xlsx】文件拖拽到这！'
        win32gui.DrawText(hdc, _str, len(_str)*2-4, rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER)
        win32gui.EndPaint(hwnd, ps)
        win32gui.DragAcceptFiles(hwnd, 1)
    elif msg == WM_DESTROY:
        win32gui.PostQuitMessage(0)
    elif msg == WM_DROPFILES:
        hDropInfo = wParam
        filescount = win32api.DragQueryFile(hDropInfo)
        for i in range(filescount):
            filename = win32api.DragQueryFile(hDropInfo, i)
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
