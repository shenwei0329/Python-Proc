#!/usr/local/bin/python2.7
# coding=utf-8
#
#   通过epic:项目入侵，统计某项目的资源投入量
#

import sys
import time
import os
import DataHandler.doPie
import DataHandler.doHour
import DataHandler.doBarOnTable
from docx.enum.text import WD_ALIGN_PARAGRAPH
import DataHandler.crWord
import types
import DataHandler.mongodb_class


reload(sys)
sys.setdefaultencoding('utf-8')

from pylab import mpl
mpl.rcParams['font.sans-serif'] = ['SimHei']

sp_name = [u'杨飞', u'吴昱珉', u'王学凯', u'许文宝',
           u'饶定远', u'金日海', u'沈伟', u'谭颖卿',
           u'吴丹阳', u'查明', u'柏银', u'崔昊之']
GroupName = [u'产品设计组', u'云平台研发组', u'大数据研发组', u'系统组', u'测试组']
ProjectAlias = {u'产品设计组': 'CPSJ', u'云平台研发组': 'FAST',
                u'大数据研发组': 'HUBBLE', u'系统组': 'ROOOT', u'测试组': 'TESTCENTER'}

"""定义时间区间
"""
st_date = '2017-11-4'
ed_date = '2017-11-6'
numb_days = 5
workhours = 40

"""公司定义的 人力资源（预算）直接成本 1000元/人天，22天/月，128元/人时
"""
CostDay = 128
Tables = ['count_record_t',]
TotalMember = 0
costProject = ()
ProductList = []
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


def is_pj(pj_info, summary):

    for _pj in pj_info:

        if _pj in (u'%s' % summary.replace(u'\xa0', u' ').replace(' ', '')).upper():
            return True

    return False


def getPjTaskListByGroup(pg):
    """
    按组列出 项目入侵 任务。
    :param pg: 插入点
    :param pj_info：项目信息
    :return:
    """

    pj_info = {}

    mongodb = DataHandler.mongodb_class.mongoDB('ext_system')
    projects = mongodb.handler('project_t', 'find', {})
    for _pj in projects:
        if _pj[u'别名'] not in pj_info:
            pj_info[_pj[u'别名']] = [_pj[u'名称'], 0.]

    issues = []
    for _grp in GroupName:

        """mongoDB数据库
        """
        mongodb.get_datebase(ProjectAlias[_grp])
        _search = {'issue_type': 'epic', 'summary': u'项目入侵'}
        _epic = mongodb.handler('issue', 'find_one', _search)
        if _epic is None:
            continue

        _search = {'epic_link': _epic['issue'], 'status': u'完成'}
        _cur = mongodb.handler('issue', 'find', _search)
        for _issue in _cur:
            issues.append(_issue)

    _count = 0
    _total_cost = 0.

    _print(u"1、项目（含立项和跟踪）的投入情况", title=True, title_lvl=3)

    _print(u'任务明细如下：')
    doc.addTable(1, 5, col_width=(4, 4, 1, 1, 1))
    _title = (('text', u'项目'),
              ('text', u'任务'),
              ('text', u'耗时'),
              ('text', u'状态'),
              ('text', u'执行人'))
    doc.addRow(_title)

    for _pj in pj_info:

        _cost = 0.
        _first = True

        for _issue in issues:

            _text = (('text', ""), )

            if _pj not in (u'%s' % _issue['summary'].replace(u'\xa0', u' ').replace(' ', '')).upper():
                continue

            for _it in ['summary', 'spent_time', 'status', 'users']:

                if type(_issue[_it]) is not types.NoneType:
                    if type(_issue[_it]) is not types.IntType:
                        _text += (('text', _issue[_it].replace(u'【项目入侵】', '')),)
                    else:
                        _text += (('text', "%0.2f" % (float(_issue[_it])/3600.)),)
                        _cost += (float(_issue['spent_time']) / 3600.)
                        _total_cost += (float(_issue['spent_time']) / 3600.)
                else:
                    _text += (('text', '-'),)

            if len(_text) > 0:
                if _first:
                    __text = (('text', u'%s' % pj_info[_pj][0]),
                             ('text', ""),
                             ('text', ""),
                             ('text', ""),
                             ('text', "")
                             )
                    doc.addRow(__text)
                    _first = False
            doc.addRow(_text)

            _count += 1

        if not _first:
            _text = (('text', ""),
                     ('text', u"小计"),
                     ('text', ""),
                     ('text', ""),
                     ('text', "%0.2f" % _cost)
                     )
            doc.addRow(_text)

    doc.setTableFont(8)
    _print("")

    _print(u"2、非项目的投入情况", title=True, title_lvl=3)

    _print(u'任务明细如下：')
    doc.addTable(1, 5, col_width=(4, 4, 1, 1, 1))
    _title = (('text', ""),
              ('text', u'任务'),
              ('text', u'耗时'),
              ('text', u'状态'),
              ('text', u'执行人'))
    doc.addRow(_title)

    _cost = 0.

    for _issue in issues:

        if is_pj(pj_info, _issue['summary']):
            continue

        _text = (('text', ""), )
        for _it in ['summary', 'spent_time', 'status', 'users']:

            if type(_issue[_it]) is not types.NoneType:
                if type(_issue[_it]) is not types.IntType:
                    _text += (('text', _issue[_it].replace(u'【项目入侵】', '')),)
                else:
                    _text += (('text', "%0.2f" % (float(_issue[_it])/3600.)),)
                    _cost += (float(_issue['spent_time']) / 3600.)
                    _total_cost += (float(_issue['spent_time']) / 3600.)
            else:
                _text += (('text', '-'),)

        doc.addRow(_text)
        _count += 1

    if len(_text) > 0:
        _text = (('text', ""),
                 ('text', u"小计"),
                 ('text', ""),
                 ('text', ""),
                 ('text', "%0.2f" % _cost)
                 )
        doc.addRow(_text)

        _text = (('text', ""),
                 ('text', u"小计"),
                 ('text', ""),
                 ('text', ""),
                 ('text', "%0.2f" % _cost)
                 )
        doc.addRow(_text)
    doc.setTableFont(8)
    _print("")

    _print(u"目前，产品研发资源共执行%d个工程项目任务，投入%0.2f工时（成本%0.2f万元）。" %
           (_count, _total_cost, (_total_cost*2.5)/(26*8)),
           paragrap=pg)

    """插入分页"""
    # doc.addPageBreak()


def main():
    """
    周报2018版本主体
    :return: 周报
    """

    global TotalMember, orgWT, costProject, fd, st_date, ed_date, numb_days, doc, workhours

    if len(sys.argv) != 3:
        print("\n\tUsage: python %s start_date end_date\n" % sys.argv[0])
        return

    st_date = sys.argv[1]
    ed_date = sys.argv[2]

    """创建word文档实例
    """
    doc = DataHandler.crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'产品部投入非产品类工作的统计报告', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)

    _print("工程项目的支撑情况", title=True, title_lvl=1)
    pg = _print("任务明细", title=True, title_lvl=2)
    getPjTaskListByGroup(pg)

    doc.saveFile('pj_support.docx')
    _cmd = 'python doc2pdf.py pj_support.docx pj_support-%s.pdf' % time.strftime('%Y%m%d',time.localtime(time.time()))
    os.system(_cmd)

    """删除过程文件"""
    _cmd = 'del /Q pic\\*'
    os.system(_cmd)


if __name__ == '__main__':

    main()

