#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
from jira import JIRA
from jira.client import GreenHopper

import types,json

options = {
    'server': 'http://172.16.60.13:8080',
}

jira = JIRA('http://172.16.60.13:8080', basic_auth=('shenwei','sw64419'))
gh = GreenHopper(options, basic_auth=('shenwei','sw64419'))

"""
projects = jira.projects()
for _proj in projects:
    print _proj

"""
"""
print dir(issue)
['JIRA_BASE_URL', '_IssueFields', '_READABLE_IDS', '__class__', '__delattr__', '__dict__', '__doc__', '__eq__', 
'__format__', '__getattr__', '__getattribute__', '__hash__', '__init__', '__module__', '__new__', '__reduce__', 
'__reduce_ex__', '__repr__', '__setattr__', '__sizeof__', '__str__', '__subclasshook__', '__weakref__', '_base_url', 
'_default_headers', '_get_url', '_load', '_options', '_parse_raw', '_resource', '_session', 'add_field_value', 
'delete', 'expand', 'fields', 'find', 'id', 'key', 'permalink', 'raw', 'self', 'update']

print dir(issue.fields)
['__class__', '__delattr__', '__dict__', '__doc__', '__format__', '__getattribute__', '__hash__', '__init__', 
'__module__', '__new__', '__reduce__', '__reduce_ex__', '__repr__', '__setattr__', '__sizeof__', '__str__', 
'__subclasshook__', '__weakref__', u'aggregateprogress', u'aggregatetimeestimate', u'aggregatetimeoriginalestimate', 
u'aggregatetimespent', u'assignee', u'components', u'created', u'creator', u'customfield_10200', u'customfield_10300', 
u'customfield_10301', u'customfield_10302', u'customfield_10303', u'customfield_10400', u'customfield_10500', 
u'customfield_10501', u'customfield_10502', u'customfield_10503', u'customfield_10504', u'customfield_10505', 
u'customfield_10506', u'customfield_10507', u'customfield_10508', u'customfield_10509', u'customfield_10510', 
u'customfield_10511', u'customfield_10800', u'customfield_11001', u'customfield_11002', u'customfield_11100', 
u'customfield_11200', u'customfield_11300', u'customfield_11304', u'customfield_11400', u'customfield_11402', 
u'customfield_11403', u'customfield_11404', u'customfield_11405', u'customfield_11406', u'customfield_11407', 
u'description', u'duedate', u'environment', u'fixVersions', u'issuelinks', u'issuetype', u'labels', u'lastViewed', 
u'priority', u'progress', u'project', u'reporter', u'resolution', u'resolutiondate', u'status', u'subtasks', 
u'summary', u'timeestimate', u'timeoriginalestimate', u'timespent', u'updated', u'versions', u'votes', u'watches', 
u'workratio']

_f = dir(issue.fields)

for __f in _f:
    if type(__f) is not types.NoneType:

        print __f

        _cmd = "if type(issue.fields.%s) is types.ListType: print issue.fields.%s" % (__f,__f)
        exec(_cmd)

        _cmd = "if type(issue.fields.%s) is types.IntType: print issue.fields.%s" % (__f, __f)
        exec (_cmd)
"""
"""
    ['JIRA_BASE_URL', '_READABLE_IDS', '__class__', '__delattr__', '__dict__', '__doc__', 
    '__eq__', '__format__', '__getattr__', '__getattribute__', '__hash__', '__init__', 
    '__module__', '__new__', '__reduce__', '__reduce_ex__', '__repr__', '__setattr__', 
    '__sizeof__', '__str__', '__subclasshook__', '__weakref__', '_base_url', '_default_headers', 
    '_get_url', '_load', '_options', '_parse_raw', '_resource', '_session', 
    'active', 'avatarUrls', 'delete', 'displayName', 'emailAddress', 'find', 'key', 'name', 
    'raw', 'self', 'timeZone', 'update']

    print(u"\tName：%s，User：%s，e-mail：%s" % (watcher.name,watcher.displayName,watcher.emailAddress))
    print(u"\ttimeZone %s，update %s" % (watcher.timeZone,watcher.update))
    if watcher.active:
        print(u"\t【有效】")
    else:
        print(u"\t【无效】")
    #print watcher._parse_raw

votes = jira.votes(issue)
print votes
logs = jira.worklogs(issue)
print logs
"""


def dispGreenHopper():
    global gh

    _f = gh.fields()
    for __f in _f:
        if __f['name'] != "Story Points":
            continue
        __cns = __f['clauseNames']
        print('-' * 8)
        for _n in __cns:
            print u"name: %s" % _n
        print "id: ", u"%s" % __f['id']
        print "name: ", u"%s" % __f['name']


def get_story_point(issue):
    if type(issue.fields.customfield_10304) is not types.NoneType:
        return issue.fields.customfield_10304
    return -1.


def get_story_epic_link(issue):
    if type(issue.fields.customfield_10300) is not types.NoneType:
        return issue.fields.customfield_10300
    return ""

def get_task_time(issue):
    return {"agg_time": issue.fields.aggregatetimeestimate,
            "org_time": issue.fields.timeoriginalestimate}


def get_version(issue):
    if len(issue.fields.fixVersions) > 0:
        return issue.fields.fixVersions[0]
    return None


def get_desc(issue):
    return issue.fields.summary


def show_name(issue):
    return str(issue)


def get_type(issue):
    return  issue.fields.issuetype


def get_status(issue):
    return issue.fields.status


def show_issue(issue):

    print("[%s]-%s" % (show_name(issue), get_desc(issue))),
    print u"类型：%s" % get_type(issue),
    print(u'状态：%s' % get_status(issue)),

    print u"里程碑：%s" % get_version(issue),
    print u"Epic link：%s" % get_story_epic_link(issue),
    print u"Story Points = %0.2f" % get_story_point(issue),
    _time = get_task_time(issue)
    print u"估计工时：%s，剩余工时：%s" % (_time['agg_time'], _time['org_time'])


def get_subtasks(issue):
    """
    收集issue的相关子任务的issue
    :param issue: 源issue
    :return: 相关issue字典
    """
    link = {}
    if not link.has_key(u"%s" % str(issue)):
        link[u"%s" % str(issue)] = []
    _task_issues = issue.fields.subtasks
    for _t in _task_issues:
        link[u"%s" % str(issue)].append(u"%s" % _t)
    return link


def get_link(issue):
    """
    收集issue的相关issue
    :param issue: 源issue
    :return: 相关issue字典
    """
    link = {}
    if not link.has_key(u"%s" % str(issue)):
        link[u"%s" % str(issue)] = []
    _task_issues = issue.fields.issuelinks
    for _t in _task_issues:
        #print dir(_t), _t.raw
        if "outwardIssue" in dir(_t):
            """该story相关的任务"""
            link[u"%s" % str(issue)].append(u"%s" % _t.outwardIssue)
        if "inwardIssue" in dir(_t):
            """该story相关的任务"""
            link[u"%s" % str(issue)].append(u"%s" % _t.inwardIssue)
    return link


def getIssue(pj_name, keys, version=None):

    """获取的最大记录数 1000 个"""
    jql_sql = u'project=%s' % pj_name
    jql_sql += u' AND created >= 2018-1-1 ORDER BY created DESC'
    total = 0
    kv = {}
    kv_link = {}
    task_link = {}

    while True:
        issues = jira.search_issues(jql_sql, maxResults=100, startAt=total)
        for issue in issues:

            if (u"%s" % issue.fields.issuetype) not in keys:
                continue

            if (len(issue.fields.fixVersions) > 0) and ((u"%s" % issue.fields.fixVersions[0]) == version):
                watcher = jira.watchers(issue)
                _user = {}
                for watcher in watcher.watchers:
                    if watcher.active:
                        _user['alias'] = watcher.name
                        _user['name'] = watcher.displayName
                        _user['email'] = watcher.emailAddress

                """收集story相关的任务"""
                task_link.update(get_link(issue))

                if not kv.has_key((u"%s" % issue.fields.issuetype)):
                    kv[(u"%s" % issue.fields.issuetype)] = 0
                    kv_link[(u"%s" % issue.fields.issuetype)] = {}

                kv[(u"%s" % issue.fields.issuetype)] += 1
                if not (kv_link[(u"%s" % issue.fields.issuetype)]).has_key(u'%s' % issue.fields.status):
                    (kv_link[(u"%s" % issue.fields.issuetype)])[u'%s' % issue.fields.status] = []

                (kv_link[(u"%s" % issue.fields.issuetype)])[u'%s' % issue.fields.status].append(u'%s' % str(issue))
                #showIssue(issue)

        if len(issues) == 100:
            total += 100
        else:
            break

    return kv, kv_link, task_link


def main():

    fast = jira.project('FAST')
    print(u"项目名称：%s，负责人：%s" % (fast.name,fast.lead.displayName))
    components = jira.project_components(fast)
    print [c.name for c in components]
    versions = jira.project_versions(fast)
    print [v.name for v in reversed(versions)]

    """获取项目版本信息
    """
    _versions = jira.project_versions('FAST')
    _version = {}
    task_link = {}
    for _v in _versions:
        if u"3.0 " not in u"%s" % _v:
            continue
        if not _version.has_key(u"%s" % _v):
            _version[u"%s" % _v] = {}
        _version[u"%s" % _v][u"id"] = _v.id
        if 'startDate' in dir(_v):
            _version[u"%s" % _v]['startDate'] = _v.startDate
        if 'releaseDate' in dir(_v):
            _version[u"%s" % _v]['releaseDate'] = _v.releaseDate
        print u"%s" % _v

        kv, kv_link, _task_link = getIssue('FAST', [u'story'], version=u"%s" % _v)
        task_link.update(_task_link)
        if not _version[u"%s" % _v].has_key(u"issues"):
            _version[u"%s" % _v][u"issues"] = {}
        _version[u"%s" % _v][u"issues"][u"key"] = kv
        _version[u"%s" % _v][u"issues"][u"link"] = kv_link

    for _v in _version:
        print u"里程碑：%s" % _v
        for _key in _version[_v][u"issues"][u"key"]:
            print(u"[类型：%s]: %d（个）" % (_key, _version[_v][u"issues"][u"key"][_key]))
            for __v in _version[_v][u"issues"][u"link"][_key]:
                print u'\t状态：%s' % __v
                for __story in _version[_v][u"issues"][u"link"][_key][__v]:
                    print u"\t\t- story：%s" % __story
                    show_issue(jira.issue(__story))
                    if task_link.has_key(__story):
                        for _task in task_link[__story]:
                            print u"\t\t\t 任务: %s" % _task
                            _issue = jira.issue(_task)
                            show_issue(_issue)
                            _link = get_link(_issue)
                            for _l in _link:
                                for __l in _link[_l]:
                                    __issue = jira.issue(__l)
                                    show_issue(__issue)


if __name__ == '__main__':
    main()
