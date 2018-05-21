#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
# 产品资源在项目中的投入汇总
# ==========================
# 2018年5月21日@成都
#
#   Usage: python doDay.py 开始日期 结束日期
#
#
#

import os
import MySQLdb
import sys
from DataHandler import crWord
from docx.enum.text import WD_ALIGN_PARAGRAPH
import time
import types
from DataHandler import mongodb_class

reload(sys)
sys.setdefaultencoding('utf-8')

"""汇总项目的别名
"""
pj_alias = u"福田基础支撑平台"

ProjectAlias = {u'产品设计组': 'CPSJ', u'云平台研发组': 'FAST',
                u'大数据研发组': 'HUBBLE', u'系统组': 'ROOOT', u'测试组': 'TESTCENTER',
                u'嘉兴项目': 'JX', u'甘孜州项目': 'GZ', u'四川公安': 'SCGA'}

"""定义时间区间
"""
days_month = 22
numb_days = 5
workhours = 40

doc = None
Topic_lvl_number = 0
Topic = [u'一、',
         u'二、',
         u'三、',
         u'四、',
         u'五、',
         u'六、',
         u'七、',
         u'八、',
         u'九、',
         u'十、',
         u'十一、',
         u'十二、',
         ]


def _print(_str, title=False, title_lvl=0, color=None, align=None, paragrap=None ):

    global doc, Topic_lvl_number, Topic

    _str = u"%s" % _str.replace('\r', '').replace('\n','')

    _paragrap = None

    if title_lvl == 1:
        Topic_lvl_number = 0
    if title:
        if title_lvl==2:
            _str = Topic[Topic_lvl_number] + _str
            Topic_lvl_number += 1
        if align is not None:
            _paragrap = doc.addHead(_str, title_lvl, align=align)
        else:
            _paragrap = doc.addHead(_str, title_lvl)
    else:
        if align is not None:
            if paragrap is None:
                _paragrap = doc.addText(_str, color=color, align=align)
            else:
                doc.appendText(paragrap, _str, color=color, align=align)
        else:
            if paragrap is None:
                _paragrap = doc.addText(_str, color=color)
            else:
                doc.appendText(paragrap, _str, color=color)
    print(_str)

    return _paragrap


def getPjTaskListByGroup(pg, pj_alias):
    """
    按组列出 项目入侵 任务。
    :param pg: 插入点
    :param pj_alias：项目别名
    :return:
    """

    _print(u'任务明细如下：')
    doc.addTable(1, 5, col_width=(2, 4, 2, 2, 2))
    _title = (('text', u'模块'),
              ('text', u'任务'),
              ('text', u'耗时'),
              ('text', u'状态'),
              ('text', u'执行人'))
    doc.addRow(_title)

    _spent_time = 0
    _count = 0

    _total_cost = 0.

    for _grp in ProjectAlias:

        """mongoDB数据库
        """
        mongodb = mongodb_class.mongoDB(ProjectAlias[_grp])

        _search = {"issue_type": u"任务",
                   "project_alias": pj_alias,
        }
        _cur = mongodb.handler('issue', 'find', _search)

        if _cur.count() == 0:
            continue

        for _issue in _cur:
            _text = ()
            for _it in ['components', 'summary', 'spent_time', 'status', 'users']:
                if type(_issue[_it]) is not types.NoneType:
                    if type(_issue[_it]) is not types.IntType:
                        _text += (('text', _issue[_it]),)
                    else:
                        _text += (('text', "%0.2f" % (float(_issue[_it])/3600.)),)
                        _spent_time += _issue[_it]
                        _total_cost += (float(_issue['spent_time']) / 3600.)
                else:
                    _text += (('text', '-'),)
            doc.addRow(_text)
            _count += 1

    _text = (('text', u'合计'),
             ('text', ""),
             ('text', "%0.2f" % (float(_spent_time)/3600.)),
             ('text', ""),
             ('text', "")
             )
    doc.addRow(_text)

    print u"Total cost：", _total_cost
    doc.setTableFont(8)
    _print("")

    _print(u"目前，产品研发资源在【%s】项目上共执行%d个工程项目任务，投入%0.2f工时。" %
           (pj_alias, _count, float(_spent_time)/3600.),
           paragrap=pg)

    """插入分页"""
    # doc.addPageBreak()


def main():
    """
    人员任务执行情况汇总报告
    :return: 报告
    """

    global st_date, ed_date, numb_days, doc, workhours, pj_alias

    """创建word文档实例
    """
    doc = crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'产品资源在项目上的投入日报', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)

    _print("工程项目的支撑情况", title=True, title_lvl=1)
    pg = _print("任务明细", title=True, title_lvl=2)
    getPjTaskListByGroup(pg, pj_alias)

    doc.saveFile('pjday.docx')
    _cmd = 'python doc2pdf.py pjday.docx pj-daily-%s.pdf' % time.strftime('%Y%m%d', time.localtime(time.time()))
    os.system(_cmd)

    """删除过程文件"""
    _cmd = 'del /Q pic\\*'
    os.system(_cmd)


if __name__ == '__main__':

    main()

#
# Eof

