#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#

import jenkins, types, time

server = jenkins.Jenkins('http://172.16.74.169:32766', username='manager', password='8RP-KnN-V5s-BzA')

print('> Jenkins version: %s' % server.get_version())

views = server.get_views()

for _view in views:
    print _view
    cnf = server.get_view_config(_view['name'])
    print cnf