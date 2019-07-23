#-*- coding: UTF-8 -*-
#
#   接收邮件
#   ========
#
#

from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import datetime
import poplib
import base64
import sys
 

def parser_date(msg):
    _date = msg.get('date').split(',')[1][1:]
    utcstr = _date.replace('+00:00','').replace(' (CST)','')
    print utcstr
    try:
        utcdatetime = datetime.datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0000 (GMT)')
        localdatetime = utcdatetime + datetime.timedelta(hours=+8)
        # localtimestamp = localdatetime.timestamp()
    except:
        utcdatetime = datetime.datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0800')
        # localtimestamp = utcdatetime.timestamp()
    ti = utcdatetime
    return "%04d%02d%02d" % (ti.year, ti.month, ti.day)


def parser_subject(msg):
    subject = msg['Subject']
    value, charset = decode_header(subject)[0]
    if charset:
        value = value.decode(charset)
    print(u'邮件主题： {0}'.format(value))
    return value
 

def parser_address(msg):
    hdr, addr = parseaddr(msg['From'])
    name, charset = decode_header(hdr)[0]
    if charset:
        name = name.decode(charset)
    print(u'发送人: {0}，邮箱: {1}'.format(name, addr))


def parser_content(msg, idx):

    for par in msg.walk():
        name = par.get_param("name")
        if name:
            print name
            value, charset = decode_header(name)[0]
            if charset:
                value = value.decode(charset)
            fname = value
            print(u"文件名: %s" % fname)
            data = par.get_payload(decode=True)
            try:
                f = open("files/%s" % fname, "wb")
            except:
                f = open("a_file", "wb")
            f.write(data)
            f.close()

        else:
            _datas = par.get_payload(decode=True)
            if _datas is None:
                continue
            f = open("htmls/email-%d.html" % idx, "wb")
            for _data in _datas:
                f.write(_data)
            f.close()


_arg_date = sys.argv[1].replace('-','')
email = "shenwei@chinacloud.com.cn"
password = sys.argv[2]
pop3_server = "pop.chinacloud.com.cn"
 
server = poplib.POP3_SSL(pop3_server)
server.set_debuglevel(0)

server.user(email)
server.pass_(password)
 
print("信息数量：%s 占用空间 %s" % server.stat())
resp, mails, octets = server.list()

_index = len(mails)
_err = False
for _idx in range(_index,0,-1):

    print(u"第%d封邮件：\n" % _idx)
    print("="*80)

    _err = False
    try:
        resp, lines, ocetes = server.retr(_idx)
    except Exception, e:
        print(">>>Err: %s" % e)
        _err = True
        server.quit()
        server = poplib.POP3_SSL(pop3_server)
        server.user(email)
        server.pass_(password)
    finally:
        if not _err:
            msg_content = b"\r\n".join(lines).decode("gbk")
            try:
                msg = Parser().parsestr(msg_content)
            except Exception, e:
                print(">>>Err.parsestr: %s" % e)
                _err = True
            finally:
                if not _err:
                    parser_address(msg)
                    parser_subject(msg)
                    _date = parser_date(msg)
                    if _date >= _arg_date:
                        parser_content(msg, _idx)
                    else:
                        break

    print("-"*80)
 
# 关闭连接
server.quit()
 
#
