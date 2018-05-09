# coding=utf8
"""
本示例使用 networkx 库构建复杂、更人性化的的关系图
安装 pip install networkx
参考 https://networkx.github.io/
"""

from __future__ import unicode_literals

import networkx as nx
from networkx.readwrite import json_graph
from pyecharts import Graph

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

from DataHandler import mongodb_class
from pyecharts import Geo

mongo_db = mongodb_class.mongoDB('ext_system')

data = []
addr_data = {}

_rec = mongo_db.handler('trip_req', 'find', {u'外出类型': u'出差'})
for _r in _rec:
    _addr = _r['起止地点'].split(u'到')
    if len(_addr) == 1:
        _addr = _r['起止地点'].split(' ')
    if len(_addr) == 1:
        _addr = _r['起止地点'].split('-')
    if len(_addr) == 1:
        _addr = _r['起止地点'].split('_')
    if len(_addr) == 1:
        _addr = _r['起止地点'].split('～')

    for __addr in _addr:
        __addr = __addr.replace(' ', '')
        print __addr
        if __addr == u'上海嘉兴':
            __addr = u'嘉兴'
        if __addr not in addr_data:
            addr_data[__addr] = 1
        else:
            addr_data[__addr] += 1

for _data in addr_data:
    if u'四川省厅' in _data:
        continue
    data.append((_data, addr_data[_data]),)

print data

geo = Geo("出差情况", "数据来自出差申请",
         title_color="#fff",
         title_pos="center",
         width=1200,
         height=600,
         background_color='#404a59')

attr, value = geo.cast(data)

geo.add("", attr, value, visual_range=[0, 10], visual_text_color="#fff", symbol_size=15, is_visualmap=True)
geo.show_config()
geo.render()
