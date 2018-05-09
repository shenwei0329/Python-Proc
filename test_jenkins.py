#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#

import jenkins, types, time

server = jenkins.Jenkins('http://172.16.74.169:32766', username='manager', password='8RP-KnN-V5s-BzA')

print('> Jenkins version: %s' % server.get_version())

nodes = server.get_nodes()
print nodes

print server.get_node_info(nodes[0]['name'])

try:
    pass
except:
    pass
finally:
    pass