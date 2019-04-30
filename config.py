# #!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time : 2019/4/16 9:54
# @Author : BruceLong
# @FileName: config.py
# @Email   : 18656170559@163.com
# @Software: PyCharm
# @Blog ：http://www.cnblogs.com/yunlongaimeng/

# from过滤池
FILTER_POOL = {
    'yanxuan2@service.netease.com', 'mailnotify@service.netease.com', 'fa163@service.netease.com',
    'yanxuan1@service.netease.com',
    'youqian@service.netease.com',
    'mail@service.netease.com',
    'yanxuan-jd@service.netease.com',
    'yanxuan6@service.netease.com',
    'yanxuan3@service.netease.com',
    'yanxuan4@service.netease.com',
    'yanxuan5@service.netease.com',
    'Postmaster@163.com',
    'youqian@service.netease.com',

}

# 需要下载的附件格式类型
FILE_TYPE = {'docx', 'doc', 'xlsx', 'eml', 'txt', 'pdf', 'xls', 'ppt'}

# 爬虫请求的URL
# 登录url
LOGIN_URL = "https://mail.{ty}.com/entry/cgi/ntesdoor?style=-1&df=mail{ty}_letter&net=&language=-1&from=web&race=&iframe=1&product=mail{ty}&funcid=loginone&passtype=1&allssl=true&url2=https://mail.{ty}.com/errorpage/error{ty}.htm"

# 首页url
INDEX_URL = 'https://mail.{ty}.com/js6/main.jsp?sid={sid}&df=mail{ty}_letter#module=welcome.WelcomeModule%7C%7B%7D'

# 获取邮件列表页url
MESSAGE_LIST_URL = 'http://mail.{ty}.com/js6/s?sid={sid}&func=mbox:listMessages&TopTabReaderShow=1&TopTabLofterShow=1&welcome_welcomemodule_mailrecom_click=1&LeftNavfolder1Click=1&mbox_folder_enter=1'

# 获取邮件详情url
DETAIL_URL = 'https://mail.{ty}.com/js6/read/readhtml.jsp?mid={mid}&font=15&color=064977'

# 获取发件方IP url
IP_URL = 'https://mail.{ty}.com/js6/s?func=mbox:getMessageData&sid={sid}&mode=text&mid={mid}&l=read&action=read_head'

# 下载附件 url
# ATTACH_URL = 'https://mail.163.com/js6/s?sid={sid}&func=mbox:readMessage&YxReadTopShow=1&TopNavMailMasterShow=1&read_toolbar_back_click=1&mbox_folder_enter=1&error=no%20Conext.module&l=read&action=read'
BF_ATTACH_URL = 'https://mail.163.com/js6/s?sid={sid}&func=mbox:readMessage&l=read&action=read'
# ATTACH_URL = 'https://mail.163.com/js6/read/readdata.jsp?sid={sid}&mid={mid}&part=3&mode=download&l=read&action=download_attach'
ATTACH_URL = 'https://mail.163.com/js6/read/readdata.jsp?sid={sid}&mid={mid}&part={id}&mode=download&l=read&action=download_attach'
