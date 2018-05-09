#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#


from jira import JIRA
import types,json,sys,MySQLdb

reload(sys)
sys.setdefaultencoding('utf-8')

def doSQLinsert(db,cur,_sql):
    try:
        cur.execute(_sql)
        db.commit()
    except:
        db.rollback()

def doSQLcount(cur,_sql):
    try:
        cur.execute(_sql)
        _result = cur.fetchone()
        _n = _result[0]
        if _n is None:
            _n = 0
    except:
        _n = 0
    return _n

def doSQL(cur,_sql):
    cur.execute(_sql)
    return cur.fetchall()

if __name__ == '__main__':

    """连接数据库"""
    db = MySQLdb.connect(host="47.93.192.232",user="root",passwd="sw64419",db="nebula",charset='utf8')
    cur = db.cursor()

    _sql = 'select task_name,task_level from project_task_t where task_resources<>"#"'
    _res = doSQL(cur, _sql)

    _n = 0
    _kv = {}
    for _name in _res:
        _level = _name[1][:-2]
        _sql = 'select task_name from project_task_t where task_level="%s"' % _level
        _topics = doSQL(cur, _sql)
        for _topic in _topics:
            _like_str = "%s-%s" % (_topic[0], _name[0])
            #print _like_str
            _sql = u'select issue_id,summary,users,state from jira_task_t where ' \
                   u'summary like "%%%s%%"' % _like_str
            _tasks = doSQL(cur, _sql)
            for _task in _tasks:
                if not _kv.has_key(_task[0]):
                    _kv[_task[0]] = []
                (_kv[_task[0]]).append([_like_str,_task[1],_task[2],_task[3]])
                _n += 1

    for _key in _kv.keys():
        print _key
        for _v in _kv[_key]:
            _str = "\t- "
            for __v in _v:
                _str += u"【%s】 " % __v
            print(_str)
        print("-"*8)
    print _n
