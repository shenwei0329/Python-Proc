#-*- coding: UTF-8 -*-
#
#   接收邮件
#   ========
#   2019-07-23 Created by shenwei @Chengdu
#   从指定邮箱上取出从昨天到现在的邮件及其附件
#
#

from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import poplib
import sys
import os
from datetime import datetime, date, timedelta
import platform
import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import base64


def send_mail(server, consignee, title, text, files=[], timeout=180, subtype='plain'):
    """
    发送邮件
    :param server: 邮件服务器属性
    :param consignee: 收件人
    :param title: 标题
    :param text: 正文
    :param files: 附件
    :param timeout: 超时
    :return:
    """
    msg = MIMEMultipart()
    msg['From'] = server['user']
    msg['Subject'] = title
    msg['To'] = ",".join(consignee)
    msg.attach(MIMEText(text, _subtype=subtype, _charset='utf-8'))
    if "Linux" == platform.system():
        coding = 'utf-8'
    else:
        coding = 'gbk'
    for file in files:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(file, 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename="=?utf-8?b?%s?=" % base64.b64encode(
            os.path.basename(file).decode(coding).encode('utf-8')))
        msg.attach(part)
    smtp = smtplib.SMTP(server['host'], server['port'] if server.get('port') else 0, timeout=timeout)
    smtp.login(server['user'], server['passwd'])
    smtp.sendmail(server['user'], consignee, msg.as_string())
    smtp.close()


def parser_date(msg):
    """
    解析邮件创建日期
    :param msg: 邮件
    :return: 创建日期（yyyymmdd）
    """
    _date = msg.get('date').split(',')[1][1:]
    if ',' in _date:
        _date = _date.split(', ')[1]
    utcstr = _date.replace('+00:00','').replace(' (CST)','')
    print utcstr
    try:
        utcdatetime = datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0000 (GMT)')
    except:
        utcdatetime = datetime.strptime(utcstr, '%d %b %Y %H:%M:%S +0800')
    ti = utcdatetime
    return "%04d%02d%02d" % (ti.year, ti.month, ti.day)


def parser_subject(msg):
    """
    解析邮件主题
    :param msg: 邮件
    :return: 主题
    """
    subject = msg['Subject']
    value, charset = decode_header(subject)[0]
    # print value, charset
    if charset:
        value = value.decode(charset)
    print(u'subject: {0}'.format(value))
    return value
 

def parser_address(msg):
    """
    解析邮件地址
    :param msg: 邮件
    :return: 邮件地址
    """
    hdr, addr = parseaddr(msg['From'])
    name, charset = decode_header(hdr)[0]
    if charset:
        name = name.decode(charset)
    # print('sender: {0}，email: {1}'.format(name, addr))
    return name, addr


def parser_content(msg, _cr_date, _sender):
    """
    解析邮件正文
    :param msg: 邮件
    :param _cr_date: 邮件创建日期（用于文件命名）
    :param _sender: 发信人（用于文件命名）
    :return:
    """
    for par in msg.walk():
        name = par.get_param("name")
        if name:
            """具有附件文件"""
            value, charset = decode_header(name)[0]
            if charset:
                value = value.decode(charset)
            f_name = value
            print(">>> %s <<<" % f_name)
            data = par.get_payload(decode=True)
            """以二进制方式写入数据"""
            try:
                f = open("files/%s" % f_name, "wb")
            except:
                f = open("files/%s-%s-ref" % (_cr_date, _sender), "wb")
            f.write(data)
            f.close()

        else:
            _datas = par.get_payload(decode=True)
            if _datas is None:
                continue
            """正文文件命名"""
            f = open("htmls/%s-%s.html" % (_cr_date, _sender), "wb")
            for _data in _datas:
                f.write(_data)
            f.close()


"""扫描昨天到今天的邮件"""
yesterday = date.today() + timedelta(days=-1)
_arg_date = yesterday.strftime("%Y%m%d")

email = "shenwei@chinacloud.com.cn"
password = sys.argv[1]
pop3_server = "pop.chinacloud.com.cn"
 
server = poplib.POP3_SSL(pop3_server)
server.set_debuglevel(0)

server.user(email)
server.pass_(password)
 
# print("信息数量：%s 占用空间 %s" % server.stat())
resp, mails, octets = server.list()

_index = len(mails)
_err = False
for _idx in range(_index, 0, -1):

    # print(u"第%d封邮件：\n" % _idx)
    # print("="*80)

    _err = False
    try:
        resp, lines, ocetes = server.retr(_idx)
    except Exception, e:
        print(">>>Err: %s" % e)
        _err = True

        """当出现server.retr错误时，需要释放它，并重新创建一个"""
        server.quit()
        server = poplib.POP3_SSL(pop3_server)
        server.user(email)
        server.pass_(password)
    finally:
        if not _err:
            if "Linux" == platform.system():
                msg_content = b"\r\n".join(lines).decode("utf-8")
            else:
                msg_content = b"\r\n".join(lines).decode("gbk")
            try:
                msg = Parser().parsestr(msg_content)
            except Exception, e:
                print(">>>Err.parsestr: %s" % e)
                _err = True
            finally:
                if not _err:
                    _name, _sender = parser_address(msg)
                    parser_subject(msg)
                    _date = parser_date(msg)
                    if _date >= _arg_date:
                        """收取正文和附件"""
                        parser_content(msg, _date, _sender)
                    else:
                        break

    print("-"*80)
 
# 关闭连接
server.quit()
 
#
