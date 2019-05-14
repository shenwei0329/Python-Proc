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
            print(u">>> [%s] ... ing" % _p)
            my_jira = jira_class_epic.jira_handler(_p.key)
            _sprints = my_jira.get_sprints()
            jira_class_epic.do_with_epic(my_jira, _sprints)
            jira_class_epic.do_with_task(my_jira)
            print(u">>> .ed.")
        except Exception, e:
            print("\n>>> Error: %s" % e)


if __name__ == '__main__':
    main()
