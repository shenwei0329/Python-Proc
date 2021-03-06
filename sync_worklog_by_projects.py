#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
from jira import JIRA
from jira.client import GreenHopper
from DataHandler import jira_class_epic

import sys

"""设置字符集
"""
reload(sys)
sys.setdefaultencoding('utf-8')


options = {
    'server': 'http://jira.chinacloud.com.cn',
}

jira = JIRA('http://jira.chinacloud.com.cn', basic_auth=('shenwei','sw64419'))
gh = GreenHopper(options, basic_auth=('shenwei','sw64419'))


def main():

    project_list = jira.projects()
    for _p in project_list:
        try:
            _project = jira_class_epic.jira_handler(_p.key)
            print(u">>> [%s]{%s} ... ing" % (_p, jira_class_epic.config.get('DATE', 'jira_scan_start_date')))
            _project.scan_worklog(jira_class_epic.config.get('DATE', 'jira_scan_start_date'))
            print(u">>> .ed.")
        except Exception, e:
            print("\n>>> Error: %s" % e)


if __name__ == '__main__':
    main()
