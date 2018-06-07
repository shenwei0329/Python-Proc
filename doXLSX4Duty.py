#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
# XLSX 文件解析器
# ===============
# 2018年6月7日@成都
#
#   收集“钉钉”考勤数据入库
#

from __future__ import unicode_literals

import sys
from DataHandler import mongodb_class
import MySQLdb

import logging
logging.basicConfig(format='%(asctime)s --%(lineno)s -- %(levelname)s:%(message)s',
                    filename='doXLSX4Duty.log', level=logging.WARN)

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
import time
from datetime import datetime
from DataHandler import mysql_hdr

db = MySQLdb.connect(host="172.16.60.2", user="tk", passwd="53ZkAuoDVc8nsrVG", db="nebula", charset='utf8')
# db = MySQLdb.connect(host="47.93.192.232",user="root",passwd="sw64419",db="nebula",charset='utf8')
MySQLhdr = mysql_hdr.SqlService(db)

RowName = ['KQ_NAME',
           'KQ_PART',
           'KQ_USERID',
           'KQ_DATE',
           'KQ_WORKTIME',
           'KQ_AM',
           'KQ_AM_STATE',
           'KQ_PM',
           'KQ_PM_STATE',
           'KQ_REF',
           'KQ_GROUP',
]
ColList = [0, 1, 2, 5, 6, 7, 8, 9, 10, 15, 3]

import locale


class XlsxHandler:

    def __init__(self, pathname):

        self.data = xlrd.open_workbook(pathname)
        self.tables = self.data.sheets()
        self.table = self.tables[0]
        self.nrows = self.table.nrows

    def getSheetNumber(self):
        return len(self.tables)

    def setSheet(self, n):
        if n < len(self.tables):
            self.table = self.tables[n]
            self.nrows = self.table.nrows

    def getNrows(self):
        return self.nrows

    def getData(self, row, col):
        """
        获取xlsx记录单元的数据
        :param table: 数据源
        :param row: 行号
        :param col: 列号
        :return:
        """
        try:
            return self.table.row_values(row)[col]
        except:
            return None

    def getXlsxColStr(self):

        # print(">>> getXlsxColStr: %d" % self.table.ncols)
        _col = self.getXlsxColName(self.table.ncols)
        _str = ""
        for _s in _col:
            _str += (_s + ',')

        # _str = _str.decode("utf-8", "replace")
        return _str, self.table.ncols

    def getXlsxColName(self, nCol):

        _col = []
        for i in range(14, nCol):
            _colv = self.table.row_values(0)[i]
            _col.append(_colv)
            # print(">>> Col[%s]" % _colv)
        return _col

    def getXlsxAllColName(self, nCol):

        _col = []
        for i in range(nCol):
            _colv = self.table.row_values(0)[i]
            _col.append(_colv)
            # print(">>> Col[%s]" % _colv)
        return _col

    def getXlsxRow(self, i, nCol, lastRow):
        """
        获取某行数据
        :param i: 行号
        :param nCol: 字段数
        :param lastRow: 上一行数据
        :return: 指定行的数据
        """

        # print("%s- getXlsxRow[%d,%d]" % (time.ctime(), i, nCol))

        """ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
        """
        _row = []
        for _i in range(nCol):
            __row = self.table.row_values(i)[_i]
            __ctype = self.table.cell(i, _i).ctype
            # print __row,__ctype
            if __ctype == 3:
                _date = datetime(*xlrd.xldate_as_tuple(__row, 0))
                __row = _date.strftime('%Y/%m/%d')
                __row = __row.split('/')
                __row = u"%d年%d月%d日" % (int(__row[0]), int(__row[1]), int(__row[2]))
            elif __ctype == 2:
                __row = str(__row).replace('.0','')
            elif __ctype == 5:
                __row = ''
            _row.append(__row)

        # print _row

        row = []
        _i = 0
        for _r in _row:

            # print(">>>[%s]" % _r)

            if _r is None or len(str(_r)) == 0:
                """用第一行的内容填充合并字段的其它单元"""
                if lastRow is not None:
                    _r = lastRow[_i]
            # print _r
            row.append(_r)
            _i = _i + 1
        return row

    def getXlsxRowByList(self, i, ColList, lastRow):
        """
        获取某行数据
        :param i: 行号
        :param ColList: 字段列表
        :param lastRow: 上一行数据
        :return: 指定行的数据
        """

        # print("%s- getXlsxRow[%d,%d]" % (time.ctime(), i, nCol))

        """ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
        """
        _row = []
        for _i in ColList:
            __row = self.table.row_values(i)[_i]
            __ctype = self.table.cell(i, _i).ctype
            # print __row,__ctype
            if __ctype == 3:
                _date = datetime(*xlrd.xldate_as_tuple(__row, 0))
                __row = _date.strftime('%Y/%m/%d')
                __row = __row.split('/')
                __row = u"%d年%d月%d日" % (int(__row[0]), int(__row[1]), int(__row[2]))
            elif __ctype == 2:
                __row = str(__row).replace('.0','')
            elif __ctype == 5:
                __row = ''
            _row.append(__row)

        # print _row

        row = []
        _i = 0
        for _r in _row:

            # print(">>>[%s]" % _r)

            if _r is None or len(str(_r)) == 0:
                """用第一行的内容填充合并字段的其它单元"""
                if lastRow is not None:
                    _r = lastRow[_i]
            # print _r
            row.append(_r)
            _i = _i + 1
        return row


def doList(xlsx_handler):

    _count = 0
    xlsx_handler.setSheet(1)
    for _idx in range(4, xlsx_handler.getNrows()):

        _row = xlsx_handler.getXlsxRowByList(_idx, ColList, None)

        _sql = 'INSERT INTO checkon_t('
        for _c in RowName:
            _sql = _sql + _c + ','
        _sql = _sql + "created_at,updated_at) VALUES("

        for _r in _row:
            if len(_r) == 0:
                _s = "'#'"
            else:
                _r = str(_r).replace('\'', '')
                _s = "'%s'" % _r
                '''去掉内容中的空格、回车键等字符
                '''
                _s = _s.replace(' ', '^')
                _s = _s.replace('\n', ' ')
                _s = _s.replace('\r', ' ')
                '''勉强的超长限制
                '''
                if len(_s) > 2048:
                    _s = _s[:2048]
            _sql = _sql + _s + ','

        _sql = _sql + "now(),now())"
        # print(">>>SQL:[%s]" % _sql)
        MySQLhdr.insert(_sql)
        _count += 1
        print ".",

    print "\n[", _count, "]", "(",xlsx_handler.getNrows(),")"
    logging.log(logging.WARN, "doList: number of record be insect: %d" % _count)


def main():

    filename = sys.argv[1]
    print filename

    xlsx_handler = XlsxHandler(filename)

    try:
        doList(xlsx_handler)
        print("%s- Done" % time.ctime())
        logging.log(logging.WARN, "main: Done!")
        return True

    except Exception, e:
        print e
        # print("%s- Done[Nothing to do]" % time.ctime())
        logging.log(logging.WARN, "main: Nothing to do!")
        return False


if __name__ == '__main__':
    main()

#
# Eof
