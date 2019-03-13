#!/usr/local/bin/python2.7
# -*- coding: utf-8 -*-
#
#   Python服务程序集
#   ================
#   2018.5.9 @Chengdu
#
#   2018.5.16 @成都
#   1）明确本程序用于建立“公司”最基本的（原始的）信息。
#   2）后续的分析类数据，通过REST-API方式获取
#
#   2018.6.23 @成都
#   1）引入ConfigParser及配置文件
#   2）程序注释
#

from __future__ import unicode_literals

import sys
import ConfigParser
import os
import json

reload(sys)
sys.setdefaultencoding('utf-8')
print sys.getdefaultencoding()

conf = ConfigParser.ConfigParser()
conf.read(os.path.split(os.path.realpath(__file__))[0] + '/rdm.cnf')

"""项目状态："""
pj_state = [u'在建', u'验收', u'交付', u'发布', u'运维']

pd_list = json.loads(conf.get('LIST', 'products'))

pj_list = json.loads(conf.get('LIST', 'projects'))

rdm_list = json.loads(conf.get('LIST', 'rdm'))

