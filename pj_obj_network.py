# -*- coding: UTF-8 -*-
#

import mongodb_class
import matplotlib
import networkx as nx
import sys

matplotlib.use('Agg')

from pylab import mpl
mpl.rcParams['font.sans-serif'] = ['SimHei']


def pj_obj_graph(macs):

    _G = nx.MultiDiGraph()
    edges = {}
    center = []

    for _mac in sorted(macs, key=lambda x: x[0]):
        if _mac[0] not in center:
            center.append(_mac[0])
        if _mac[0]+_mac[1] not in edges:
            edges[_mac[0] + _mac[1]] = 1
            _G.add_edge(_mac[0], _mac[1], **({'weight': _mac[2]}))

    label = {}
    for _n in _G.nodes():

        try:
            label[_n] = _n
        except Exception, e:
            print e
            continue

    return _G, label, center


def do_main(_bg_date, _now):

    db = mongodb_class.mongoDB()
    db.connect_db("WORK_LOGS")

    print _bg_time, " --> ", _now

    _sql = {
            "$and": [{"created": {"$gt": _bg_time}}, {"created": {"$lte": _now}}]
            }

    _rec = db.handler("worklog", "find", _sql)

    _macs = []

    print _rec.count()

    if _rec.count() == 0:
        return

    for _r in _rec:
        if "TESTCENTER" in _r['issue'].split('-')[0]:
            continue
        _mac = (_r['issue'].split('-')[0], _r['author'], float(_r['timeSpentSeconds'])/3600.)
        # print _mac
        _macs.append(_mac)

    G, _label, _center = pj_obj_graph(_macs)

    # print _center

    H = nx.Graph(G)
    edgewidth = []
    wins = dict.fromkeys(G.nodes(), 0.01)

    for (u, v, d) in H.edges(data=True):
        edgewidth.append(d['weight'])
        if u in _center:
            wins[u] += d['weight']
        if v in _center:
            wins[v] += d['weight']

    # print wins

    try:
        import matplotlib.pyplot as plt

        plt.rcParams['text.usetex'] = False
        plt.figure(figsize=(18, 24), dpi=120)
        plt.rcParams['text.usetex'] = False
        plt.axis('off')
        pos = nx.spring_layout(H, iterations=16)
        nx.draw_networkx_edges(H, pos, alpha=0.3, width=edgewidth, edge_color='y')
        _node_size = [wins[v]*60 for v in H]
        nx.draw_networkx_nodes(H, pos, node_size=_node_size, node_color='m', alpha=0.2)
        nx.draw_networkx_edges(H, pos, alpha=0.4, node_size=0, width=0.2, edge_color='k')
        nx.draw_networkx_labels(H, pos, fontsize=8)

        font = {'fontname': 'SimHei',
                'color': 'k',
                'fontweight': 'bold',
                'fontsize': 16}
        plt.title(u"资源分布【%s - %s】" % (_bg_time, _now), font)

        font = {'fontname': 'SimHei',
                'color': 'k',
                'fontweight': 'bold',
                'fontsize': 12}
        plt.text(0.5, 0.98, u"项目节点大小 = 总工作量", font,
                 horizontalalignment='center',
                 transform=plt.gca().transAxes)
        plt.text(0.5, 0.96, u"边大小 = 个人工作量", font,
                 horizontalalignment='center',
                 transform=plt.gca().transAxes)

        plt.savefig("network%s-%s.png" % (_bg_time.split('T')[0].replace('-', ''), _now.split('T')[0].replace('-', '')))

    except Exception, e:
        print e
        return


if __name__ == '__main__':

    """
    yesterday = date.today() + timedelta(days=-1)
    _bg_time = yesterday.strftime("%Y-%m-%d") + " 00:00:00"
    _now = str(datetime.now()).split(' ')[0] + " 00:00:00"
    """

    if len(sys.argv) == 3:
        _bg_time = sys.argv[1] + "T00:00:00"
        _now = sys.argv[2] + "T00:00:00"
        do_main(_bg_time, _now)
