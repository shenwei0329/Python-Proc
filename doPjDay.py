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

import sys
from DataHandler import crWord
from docx.enum.text import WD_ALIGN_PARAGRAPH
import time
from DataHandler import PersonalStat

reload(sys)
sys.setdefaultencoding('utf-8')

"""汇总项目的别名
"""
pj_alias = u"福田基础支撑平台"

PdAlias = {u'产品设计组': 'CPSJ',
           u'云平台研发组': 'FAST',
           u'大数据研发组': 'HUBBLE',
           u'系统组': 'ROOOT',
           }

RdmAlias = {
           u'测试组': 'TESTCENTER',
           u'研发管理': 'RDM',
            }

PjAlias = {u'嘉兴项目': 'JX',
           u'甘孜州项目': 'GZ',
           u'四川公安': 'SCGA',
           }

"""定义时间区间
"""
st_date = '2018-05-01'
ed_date = '2018-05-31'
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


def main():
    """
    人员任务执行情况汇总报告
    :return: 报告
    """

    global doc, st_date, ed_date

    if len(sys.argv) == 3:
        st_date = sys.argv[1]
        ed_date = sys.argv[2]

    Persional = PersonalStat.Persional(st_date, ed_date)

    """创建word文档实例
    """
    doc = crWord.createWord()
    """写入"主题"
    """
    doc.addHead(u'人员工作情况汇总', 0, align=WD_ALIGN_PARAGRAPH.CENTER)

    _print('>>> 报告生成日期【%s】 <<<' % time.ctime(), align=WD_ALIGN_PARAGRAPH.CENTER)
    _print('')

    _print(u'统计时间段：%s 至 %s' % (st_date, ed_date))

    _print(u"产品研发组情况", title=True, title_lvl=1)

    _pg = _print(u'明细数据：')

    _member_count = 0
    for _grp in PdAlias:

        Persional.clearPersional()

        _print(u"【%s】组的任务执行情况" % _grp, title=True, title_lvl=2)

        doc.addTable(1, 6, col_width=(2, 2, 2, 2, 2, 2))
        _title = (('text', u'名称'),
                  ('text', u'任务总数'),
                  ('text', u'完成数'),
                  ('text', u'完成率'),
                  ('text', u'任务工时数'),
                  ('text', u'工作日志\n记录的工时数'),
                  )
        doc.addRow(_title)

        Persional.scanProject(PdAlias[_grp])

        _member_count += Persional.getNumbOfMember()

        _workhour = 0
        _wl_workhour = 0

        for _persion in Persional.getNameList():

            _done, _ratio = Persional.getNumberDone(_persion)
            _wh = Persional.getSpentTime(_persion)/3600.
            _workhour += _wh

            _wl_wh = Persional.getWorklogSpentTime(_persion)/3600.
            _wl_workhour += _wl_wh

            _text = (('text', u'%s' % _persion),
                     ('text', '%d' % Persional.getNumbOfTask(_persion)),
                     ('text', '%d' % _done),
                     ('text', '%0.2f%%' % _ratio),
                     ('text', '%0.2f' % _wh),
                     ('text', '%0.2f' % _wl_wh),
                     )
            doc.addRow(_text)

        _text = (('text', u"总计"),
                 ('text', ""),
                 ('text', ""),
                 ('text', ""),
                 ('text', "%0.2f" % _workhour),
                 ('text', "%0.2f" % _wl_workhour),
                 )
        doc.addRow(_text)
        doc.setTableFont(8)
        _print("")

    _print(u'产品中心总人数：%d' % _member_count, paragrap=_pg)

    _print(u"项目开发组情况", title=True, title_lvl=1)

    _pg = _print(u'明细数据：')

    _member_count = 0
    for _grp in PjAlias:

        Persional.clearPersional()

        _print(u"【%s】组的任务执行情况" % _grp, title=True, title_lvl=2)

        doc.addTable(1, 6, col_width=(2, 2, 2, 2, 2, 2))
        _title = (('text', u'名称'),
                  ('text', u'任务总数'),
                  ('text', u'完成数'),
                  ('text', u'完成率'),
                  ('text', u'任务工时数'),
                  ('text', u'工作日志\n记录的工时数'),
                  )
        doc.addRow(_title)

        Persional.scanProject(PjAlias[_grp])

        _member_count += Persional.getNumbOfMember()

        _workhour = 0
        _wl_workhour = 0

        for _persion in Persional.getNameList():

            _done, _ratio = Persional.getNumberDone(_persion)
            _wh = Persional.getSpentTime(_persion)/3600.
            _workhour += _wh

            _wl_wh = Persional.getWorklogSpentTime(_persion)/3600.
            _wl_workhour += _wl_wh

            _text = (('text', u'%s' % _persion),
                     ('text', '%d' % Persional.getNumbOfTask(_persion)),
                     ('text', '%d' % _done),
                     ('text', '%0.2f%%' % _ratio),
                     ('text', '%0.2f' % _wh),
                     ('text', '%0.2f' % _wl_wh),
                     )
            doc.addRow(_text)

        _text = (('text', u"总计"),
                 ('text', ""),
                 ('text', ""),
                 ('text', ""),
                 ('text', "%0.2f" % _workhour),
                 ('text', "%0.2f" % _wl_workhour),
                 )
        doc.addRow(_text)
        doc.setTableFont(8)
        _print("")

    _print(u'注册的项目开发总人数：%d' % _member_count, paragrap=_pg)

    doc.saveFile('persion.docx')

    import os
    _cmd = 'python doc2pdf.py persion.docx persion-%s.pdf' % time.strftime('%Y%m%d',time.localtime(time.time()))
    os.system(_cmd)


if __name__ == '__main__':

    main()

#
# Eof

