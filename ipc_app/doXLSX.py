#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#

import sysv_ipc as ipc
import xlrd,time,os,MySQLdb

TableDic = {
    'MEMBER': 'member_t',
    'DATAELEMENT':'data_element_t',
}

try:
    q = ipc.MessageQueue(19640419001,ipc.IPC_CREAT | ipc.IPC_EXCL)
except:
    q = ipc.MessageQueue(19640419001)

def getData(table,col,row):

    try:
        return table.row_values(col)[row]
    except:
        return None

def getXlsxHead(table):

    _type = getData(table,0,0)
    _op = getData(table,0,1)
    _nrec = getData(table,0,2)
    if _nrec is not None:
        _nrec = int(_nrec)
    _rec = getData(table,0,3)
    print _type,_op,_nrec,_rec
    return _type,_op,_nrec,_rec

def getXlsxColName(table,nCol):

    _col = []
    for i in range(nCol):
        _colv = table.row_values(1)[i]
        _col.append(_colv)
        print(">>>Col[%s]" % _colv)
    return _col

def getXlsxRow(table,i,nCol):

    _row = table.row_values(i)[:nCol]
    row = []
    for _r in _row:
        row.append(_r)
    return row

def doSQL(_sql):

    db = MySQLdb.connect(host="localhost",user="root",passwd="mysqlroot",db="nebula",charset='utf8')
    cursor = db.cursor()
    try:
        cursor.execute(_sql)
        db.commit()
    except:
        db.rollback()

    db.close()

def doRecord(tab,_type,_op):
    print(">>>doRecord")

def main():
    try:
        m = q.receive(block=False)
        if m:
            print m[0]
            try:
                data = xlrd.open_workbook(m[0])
                table = data.sheets()[0]
                nrows = table.nrows
                print(">>>nrows: %d" % nrows)
                
                _type,_op,_ncol,_rectype = getXlsxHead(table)
                if not TableDic.has_key(_type):
                    print(">>>Type invalid![%s]" % _type)
                else:
                    if (_rectype is not None) and ('RECORD' in _rectype):
                        print(">>>Type not be LIST")
                    else:
                        _col = getXlsxColName(table,_ncol)

                        _rows = []
                        for i in range(2,nrows):
                            _row = getXlsxRow(table,i,_ncol)
                            if _row[0][0]==":":
                                continue
                            _rows.append(_row)

                        print _rows

                        if len(_rows)>0:

                            for _row in _rows:
                                _sql = 'INSERT INTO '+ TableDic[_type] + '('
                                for _c in _col:
                                    _sql = _sql + _c + ','
                                _sql = _sql + "created_at,updated_at) VALUES("
                                print _row
                                for _r in _row:
                                    print _r
                                    _sql = _sql + "'%s'"%_r + ','
                                _sql = _sql + "now(),now())"
                                print _sql
                                print(">>>SQL:[%s]" % _sql)
                                doSQL(_sql)
                        print(">>>Done")
            except:
                print("Xlsx file open Error!")
            os.remove(m[0])
    except:
        print(">>>Done[Nothing to do]")


if __name__ == '__main__':
    main()

#
# Eof
