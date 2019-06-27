# -*- coding: utf-8 -*-

import win32gui
import win32con
import win32api
import sys
import os
from win32con import *
import doXLSX4ext_org

# from __future__ import unicode_literals

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


def fileHandler(_file):

    if (('.xlsx' in _file) or ('.xls' in _file)) and ('~$' not in _file):
        doXLSX4ext_org.main(_file)
    else:
        print "Invalid file: ", _file
        return


def scan_files(directory):
    files_list = []

    for root, sub_dirs, files in os.walk(directory):
        for special_file in files:
            files_list.append(os.path.join(root, special_file))
    return files_list


def WndProc(hwnd, msg, wParam, lParam):

    if msg == WM_PAINT:
        hdc,ps = win32gui.BeginPaint(hwnd)
        rect = win32gui.GetClientRect(hwnd)
        _str = u'请把【星任务】文件拖拽到这！'
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
            if os.path.isdir(filename):
                _files = scan_files(filename)
            else:
                _files = [filename]

            if len(_files) > 0:
                for _f in _files:
                    fileHandler(_f)
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
