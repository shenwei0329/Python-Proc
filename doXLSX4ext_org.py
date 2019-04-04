#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
# XLSX 文件解析器
# ===============
#
#

from __future__ import unicode_literals

import sys
from DataHandler import mongodb_class

import logging
logging.basicConfig(format='%(asctime)s --%(lineno)s -- %(levelname)s:%(message)s',
                    filename='doXLSX4ext.log', level=logging.WARN)

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
import time
from datetime import datetime


class XlsxHandler:

    def __init__(self, pathname):

        # logging.log(logging.WARN, u"XlsxHandler __init__(%s)" % pathname)

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
        for i in range(1, nCol):
            _colv = self.table.row_values(0)[i]
            _col.append(_colv)
            # print(">>> Col[%s]" % _colv)
        return _col

    def getXlsxAllColName(self, nCol):

        _col = []
        for i in range(nCol):
            _colv = self.table.row_values(1)[i]
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
                __row = str(__row)
                # __row = str(__row).replace('.0', '')
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


def doList(xlsx_handler, mongodb, _type, _op, _ncol, keys):

    _keys = keys

    # print("%s- doList ing <%d:%d>" % (time.ctime(), _ncol, xlsx_handler.getNrows()))

    _rows = []
    for i in range(1, xlsx_handler.getNrows()):
        _row = xlsx_handler.getXlsxRow(i, _ncol, None)
        if not _row[0].isdigit():
            continue
        _rows.append(_row)

    _col = xlsx_handler.getXlsxAllColName(_ncol)

    # print("...5")
    # print _col

    _count = 0
    _key = []
    if len(_rows) > 0:

        if 'APPEND' in _op:
            '''追加方式，如日志记录
            '''
            for _row in _rows:

                _value = {}
                _i = 0

                _search = {}
                for _v in _keys:
                    _search[_col[_v]] = _row[_v]

                # print _search

                # _v = mongodb.handler(_type, 'find_one', _search)
                # print _v

                for _c in _col:
                    _value[_c] = _row[_i]
                    _i += 1

                print ">>> update table: ", _type
                try:
                    if _search not in _key:
                        # mongodb.handler(_type, 'remove', _search)
                        mongodb.handler(_type, 'update', _search, _value)
                        _key.append(_search)
                        _count += 1
                except Exception, e:
                    print "error: ", e
                finally:
                    print '.',
        print "[", _count, "]"
        logging.log(logging.WARN, "doList: number of record be inputted: %d" % _count)


def main(filename):

    mongo_db = mongodb_class.mongoDB('ext_system')

    if filename is None:
        filename = sys.argv[1]
    print filename

    xlsx_handler = XlsxHandler(filename)

    try:

        logging.log(logging.WARN, ">>> 2.")

        _str, _ncols = xlsx_handler.getXlsxColStr()

        _table = "star_task"

        logging.log(logging.WARN, ">>> 3.")

        doList(xlsx_handler, mongo_db, _table, "APPEND", _ncols, range(3))
        print("%s- Done" % time.ctime())
        logging.log(logging.WARN, "main: Done!")
        return True

    except Exception, e:
        print e
        # print("%s- Done[Nothing to do]" % time.ctime())
        logging.log(logging.WARN, "main: Nothing to do!")
        return False


if __name__ == '__main__':
    main(None)

#
# Eof
