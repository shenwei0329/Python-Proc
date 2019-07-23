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

import xlrd
import time
from datetime import datetime

import logging
logging.basicConfig(format='%(asctime)s --%(lineno)s -- %(levelname)s:%(message)s',
                    filename='doXLSX4ext.log', level=logging.WARN)

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')

special_file_name = {u'部署实施': 'devops_task',
                     u'devops': 'devops_task',
                     u'【公安】运维': 'ops_task',
                     u'【运维】北区': 'ops_task_bj',
                     u'【北区】运维': 'ops_task_bj',
                     u"outfile_": 'git_log',
                     }


class XlsxHandler:

    def __init__(self, pathname):

        # logging.log(logging.WARN, u"XlsxHandler __init__(%s)" % pathname)

        import chardet

        _result = chardet.detect(pathname)
        self.filen = pathname.decode(_result['encoding']).encode("utf8")
        self.table_name = "star_task"
        self.isSpecial()

        self.data = xlrd.open_workbook(pathname)
        self.tables = self.data.sheets()
        self.table = self.tables[0]
        self.nrows = self.table.nrows

    def isSpecial(self):
        for _f in special_file_name:
            print _f, self.filen
            if _f in self.filen:
                self.table_name = special_file_name[_f]
                return True
        return False

    def getTableName(self):
        return self.table_name

    def getSheetNumber(self):
        return len(self.tables)

    def setSheet(self, n):
        if n < len(self.tables):
            self.table = self.tables[n]
            self.nrows = self.table.nrows

    def getNrows(self):
        if self.isSpecial():
            return self.nrows-1
        else:
            return self.nrows-3

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

        print(">>> getXlsxColStr: %d" % self.table.ncols)
        _col = self.getXlsxColName(self.table.ncols)

        for _c in _col:
            print(u"%s" % _c)

        _str = ""
        for _s in _col:
            _str += (_s + ',')

        # _str = _str.decode("utf-8", "replace")
        return _str, self.table.ncols

    def getFirstRow(self):
        if self.isSpecial():
            return 1
        else:
            return 2

    def getXlsxColName(self, nCol):

        _col = []
        if u'工单记录及发包情况' in self.filen:
            return [u'任务', u'执行人', u'日期']

        if self.isSpecial():
            for i in range(0, nCol):
                _colv = self.table.row_values(0)[i]
                _col.append(_colv)
                # print(u">>> Col[%s]" % _colv)
        else:
            for i in range(0, nCol):
                _colv = self.table.row_values(1)[i]
                _col.append(_colv)
                # print(">>> Col[%s]" % _colv)
        return _col

    def getXlsxAllColName(self, nCol):

        if u'工单记录及发包情况' in self.filen:
            return [u'任务', u'执行人', u'日期']

        _col = []
        for i in range(nCol):
            if self.isSpecial():
                _colv = self.table.row_values(0)[i]
            else:
                _colv = self.table.row_values(1)[i]
            _col.append(_colv)
            # print(u">>> Col[%s]" % _colv)
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

            try:
                print "getXlsxRow: ", __row, __ctype
            except Exception, e:
                print e
                __row = "*"
            finally:
                __row = __row

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
            try:
                _row.append(__row)
            except Exception, e:
                print e
                _row.append("*")
            finally:
                __row = ""

        print ">>> ", _row

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


def doList(xlsx_handler, mongodb, _type, _op, _ncol):

    # print("%s- doList ing <%d:%d>" % (time.ctime(), _ncol, xlsx_handler.getNrows()))

    _rows = []
    _first = xlsx_handler.getFirstRow()
    for i in range(_first, _first + xlsx_handler.getNrows()):
        _row = xlsx_handler.getXlsxRow(i, _ncol, None)
        _rows.append(_row)

    _col = xlsx_handler.getXlsxAllColName(_ncol)

    # print("...5")
    # print _col

    _count = 0
    if len(_rows) > 0:

        if 'APPEND' in _op:
            '''追加方式，如日志记录
            '''
            for _row in _rows:

                _value = {}
                _i = 0

                for _c in _col:
                    _value[_c] = _row[_i]
                    # print(u">>> [%s] = <%s>" % (_c, _row[_i]))
                    _i += 1

                print ">>> update table: ", _type
                try:
                    mongodb.handler(_type, 'update', _value, _value)
                    # print _value
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
        print _str
        print _ncols
        # _table = "star_task"
        _table = xlsx_handler.getTableName()

        logging.log(logging.WARN, ">>> 3.")

        doList(xlsx_handler, mongo_db, _table, "APPEND", _ncols)
        print("%s- Done<%s>" % (time.ctime(), _table))
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
