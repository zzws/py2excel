#!/usr/bin/python2.7
 # -*- coding: UTF-8 -*-
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr(( \
        Header(name, 'utf-8').encode(), \
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

print "begin"
from_addr = 'xxx@xxx.com'
password = 'xxx'

# 输入SMTP服务器地址:
smtp_server = 'smtp.exmail.qq.com'
# 输入收件人地址:
to_addr = 'zhangcan102@126.com'


msg = MIMEText('hello, send by Python...', 'plain', 'utf-8')
msg['From'] = _format_addr( from_addr)
msg['To'] = _format_addr(u'管理员 <%s>' % to_addr)
msg['Subject'] = Header(u'来自SMTP的问候……', 'utf-8').encode()

server = smtplib.SMTP(smtp_server, 25)
server.set_debuglevel(1)
server.login(from_addr, password)
server.sendmail(from_addr, [to_addr], msg.as_string())
server.quit()