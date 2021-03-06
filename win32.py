# -*- coding: utf-8 -*-

import win32gui
import win32con
import win32api
import sys
from win32con import *
import doXLSX4ext

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


def fileHandler(_file):

    f = open('ext_file.txt', 'r')
    file_list = f.read()
    f.close()

    _short_file = _file.split("\\")[-1]
    print "fileHandler: ", _short_file

    if _short_file in file_list:
        print "Nothing to do!( in the ext_file.txt )"
        return

    if (('.xlsx' not in _short_file) and ('.xls' in _short_file)) and ('~$' not in _short_file):
        doXLSX4ext.main(_file)
        file_list += _short_file
    else:
        print "Invalid file name: ", _short_file
        return

    f = open('ext_file.txt', 'w')
    f.write(file_list)
    f.close()


def WndProc(hwnd, msg, wParam, lParam):

    if msg == WM_PAINT:
        hdc,ps = win32gui.BeginPaint(hwnd)
        rect = win32gui.GetClientRect(hwnd)
        _str = u'请把【钉钉】文件拖拽到这！'
        win32gui.DrawText(hdc, _str, len(_str)*2, rect, DT_SINGLELINE | DT_CENTER | DT_VCENTER)
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
