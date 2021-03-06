#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#
#

from DataHandler import mongodb_class


class Personal:
    """
    构建个人信息库
    """

    def __init__(self, date=None, landmark=None):
        self.personal = {}
        self.mongodb = mongodb_class.mongoDB(None)
        self.landmark = landmark
        if date is None:
            self.st_date = "2018-01-01"
            self.ed_date = "2018-12-31"
        else:
            self.st_date = date['st_date']
            self.ed_date = date['ed_date']

    def _getTaskListByPersonal(self, project):
        """
        按组列出 人员-任务
        :param project: 项目
        :return:
        """

        """mongoDB数据库
        """
        self.mongodb.setDataBbase(project)

        if self.landmark is None:
            _search = {"issue_type": {"$ne": ["epic", "story"]},
                       "$or": [{"spent_time": {"$ne": None}},
                               {"org_time": {"$ne": None}}],
                       "$and": [{"created": {"$gte": "%s" % self.st_date}},
                                {"created": {"$lt": "%s" % self.ed_date}}]
                       }
        else:
            _search = {"issue_type": {"$ne": ["epic", "story"]},
                       "landmark": self.landmark,
                       "$or": [{"spent_time": {"$ne": None}},
                               {"org_time": {"$ne": None}}],
                       }
        _cur = self.mongodb.handler('issue', 'find', _search)

        if _cur.count() == 0:
            return

        for _issue in _cur:
            if _issue['users'] is None:
                continue
            if _issue['users'] not in self.personal:
                self.personal[_issue['users']] = {'issue': [], 'worklog': []}
            _task = {}
            for _i in ['issue', 'summary', 'status', 'org_time',
                       'agg_time', 'spent_time', 'sprint', 'created', 'updated']:
                _task[_i] = _issue[_i]
            self.personal[_issue['users']]['issue'].append(_task)

    def _getWorklogListByPersonal(self, project):
        """
        按组列出 人员-工作日志
        :param project: 项目
        :return:
        """

        """mongoDB数据库
        """
        self.mongodb.setDataBbase(project)

        for _personal in self.personal:
            _search = {"author": u'%s' % _personal,
                       "$and": [{"started": {"$gte": "%s" % self.st_date}},
                                {"started": {"$lt": "%s" % self.ed_date}}]
                       }
            _cur = self.mongodb.handler('worklog', 'find', _search)

            if _cur.count() == 0:
                continue

            _worklog = {}
            for _wl in _cur:
                for _i in ['issue', 'comment', 'timeSpentSeconds', 'created', 'started', 'updated']:
                    _worklog[_i] = _wl[_i]
                self.personal[_personal]['worklog'].append(_worklog)

    def getWorklogSpentTime(self, name):
        """
        获取个人工作日志中工时统计
        :param name: 名称
        :return:
        """
        worklogs = self.personal[name]['worklog']
        _spent_time = 0
        for _issue in worklogs:
            if _issue['timeSpentSeconds'] is None:
                continue
            _spent_time += float(_issue['timeSpentSeconds'])

        return _spent_time

    def getSpentTime(self, name):
        """
        获取个人任务中工时统计
        :param name: 名称
        :return:
        """
        issues = self.personal[name]['issue']
        _time = 0
        for _issue in issues:
            if _issue['spent_time'] is None:
                continue
            _time += float(_issue['spent_time'])

        return _time

    def getOrgTime(self, name):
        """
        获取个人任务中工时统计
        :param name: 名称
        :return:
        """
        issues = self.personal[name]['issue']
        _time = 0
        for _issue in issues:
            if _issue['org_time'] is None:
                continue
            _time += float(_issue['org_time'])

        return _time

    def getNumberDone(self, name):
        """
        获取个人已完成任务的统计
        :param name:
        :return:
        """
        issues = self.personal[name]['issue']
        _done = 0
        for _issue in issues:
            if _issue['status'] == u'完成':
                _done += 1
        return _done, float(_done*100)/float(len(issues))

    def dispPersonal(self, personal):
        """
        显示个人业绩
        :param persion: 个人数据汇总
        :return:
        """
        for _name in personal:
            _done, _ratio = self.getNumberDone(personal[_name]['issue'])
            print(u'>>> %s: Task=%d, Done=%d, R=%0.2f%%, Spent=%0.2f (工时)' % (
                _name,
                len(personal[_name]['issue']),
                _done,
                _ratio,
                self.getSpentTime(_name)/3600.))

    def clearPersonal(self):
        self.personal = {}

    def scanProject(self, project):
        self._getTaskListByPersonal(project)
        self._getWorklogListByPersonal(project)

    def getPersonal(self, name=None):
        if name is None:
            return self.personal
        else:
            return self.personal[name]

    def getNameList(self):
        return self.personal.keys()

    def getNumbOfTask(self, name):
        return len(self.personal[name]['issue'])

    def getNumbOfMember(self):
        return len(self.personal)
