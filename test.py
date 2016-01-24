#!-*-coding:utf-8-*-
import urllib
import urllib2
import cookielib
from urllib import unquote
from urllib import quote
import re
import os
import getpass
import socket
import time

print unicode("************************************************","utf-8")
print unicode("SG系统增强工具v2.0","utf-8")
print unicode("功能：1.附件批量下载 2.工单批量回单 3.工单批量退单","utf-8")
print unicode("说明：1.SG附件默认保存在D:\SG中","utf-8")
print unicode("     2.SG附件不存在清单保存在D:\SG工单无附件清单.txt中","utf-8")
print unicode("     3.SG工单回单失败请单保存在D:\SG工单回单失败清单.txt中","utf-8")
print unicode("     4.win7下如遇闪退请使用兼容性运行","utf-8")
print unicode("************************************************","utf-8")

def cbk(a, b, c):
    per = 100.0 * a * b / c
    if per > 100:
        per = 100
    #print '%.2f%%' % per

#设置cookie
cookie = cookielib.CookieJar()
opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie))
#设置header
headers = {
    'User-Agent':'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/8.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)'
}

while(True):
    username = raw_input('Username:')
    password = getpass.getpass('Password:')
    #登陆需要POST的数据#
    postdata=urllib.urlencode({
        'password':password,
        'username':username,
        'x':'43',
        'y':'6'
    })
    #登陆系统，获取cookie
    req_checkUser = urllib2.Request(
        headers = headers,
        url = 'http://132.121.92.224:7001/WebSg/checkUser.action',
        data = postdata
    )
    result = opener.open(req_checkUser)
    #获取登陆者姓名
    try:
        loginUser = opener.open('http://132.121.92.224:7001/WebSg/mainTop.action')
        page_print_2 = loginUser.read()
        reg_u = re.compile('\<font color=\"\#ffffff\">&nbsp;([^&]*)&')
        q5 = reg_u.findall(page_print_2)
        print q5[0]
        s7 = q5[0].decode('gb2312').encode('utf-8')
        s8 = s7.split('：')
        loginUsername = s8[1]
        break
    except:
        print u'账号或密码错误！请重新输入！'
    
select = "exit"
while(select != ""):
    #选择操作类型
    print u'请选择操作类型：1.批量下载附件 2.SG批量回单 3.SG批量退单'
    str_select = '请选择：'
    select = raw_input(str_select.decode('utf-8').encode('gb2312'))

    #访问待处理工单列表
    response = opener.open('http://132.121.92.224:7001/WebSg/toShowAllWork.action?srcPage=allWorkArea&onlyOperator=true&sgWebMenuId=181')
    response3 = opener.open('http://132.121.92.224:7001/WebSg/listAllWorkFromMiddQuery.action?srcPage=allWorkArea')
    reg_n=re.compile(r'xPageBarBuild[^\]]*\]\)')
    q4=reg_n.findall(response3.read())
    s4=q4[0]
    s5=s4.split(',')
    str_perpage = s5[1]
    str_pageNum = int(s5[2])
    str_sgNumber =int(s5[3])
    print u'检测到' + str(str_sgNumber) + u'张SG工单...'

    if(int(select)==1):
        print u'请选择起止工单数（如1到10张）：'
        start ='起始单数：'
        end = '结束单数：'
        count_pre = raw_input(start.decode('utf-8').encode('gb2312'))
        count_suf = raw_input(end.decode('utf-8').encode('gb2312'))
        print u'请等待...'

        #设定每页展示的工单数量
        postdata_pageNumber=urllib.urlencode({
            'pagination.index':'1',
            'pagination.size':str_sgNumber
        })
        req_page = urllib2.Request(
            headers = headers,
            url = 'http://132.121.92.224:7001/WebSg/listAllWorkFromMiddQuery.action?srcPage=allWorkArea',
            data = postdata_pageNumber
        )
        response1 =opener.open(req_page)
        page_print_1 = response1.read()

        reg_s=re.compile(r'\s+\'[^\w]*TaskID[^\']*\'')
        q3=reg_s.findall(page_print_1)

        localDir = 'D:\\SG\\'
        if(os.path.exists(localDir)==False):
            os.mkdir(localDir)

        #创建txt文件并清空
        f = open(u'D:\\SG工单无附件清单.txt','a')
        f.truncate()

        if (str_sgNumber>0):
            for i in range(int(count_pre)-1,int(count_suf)):
                #分割块
                s1=q3[i].split('\'')
                s2=s1[1]
                s3=s2.split(',')
                str_taskName=quote(s3[0])
                str_taskID=quote(s3[1])
                str_orderTicketId=s3[3]
                str_workTicketId=s3[2]
                str_orderSerialNo=s3[4]
                str_tacheInstId=s3[5]
                str_flowInstId=s3[6]
                str_workDirection=quote(s3[7])
                str_sgWorkListInfoId=s3[8]
                str_codeCount=s3[9]
                str_childOrderCount=s3[10]


                #获取抓附件的url post
                postdata_download=urllib.urlencode({
                    'taskName':str_taskName,
                    'taskId':str_taskID,
                    'orderTicketId':str_orderTicketId,
                    'workTicketId':str_workTicketId,
                    'orderSerialNo':str_orderSerialNo,
                    'tacheInstId':str_tacheInstId,
                    'flowInstId':str_flowInstId,
                    'workDirection':str_workDirection,
                    'sgWorkListInfoId':str_sgWorkListInfoId,
                    'codeCount':str_codeCount,
                    'childOrderCount':str_childOrderCount
                })
                req_cl =urllib2.Request(
                    headers = headers,
                    url = 'http://132.121.92.224:7001/WebSg/infoAllWork.action?srcPage=allWorkArea',
                    data = postdata_download
                )
                try:
                    response3 =opener.open(req_cl)
                
                    url_file='http://132.121.92.224:7001/WebSg/showTicketAttachs.action?ticketId=' + str_workTicketId + '\''
                    response2 =opener.open(url_file)
                    reg1=re.compile(r'WebSg[^\"]*')
                    q1=reg1.findall(response2.read())
                except:
                    pass

                if (q1):
                    #分割文件名字符串
                    #下载url=主网址+附件url
                    url_base ='http://132.121.92.224:7001/'
                    url_create = url_base + q1[0]
                    str1=q1[0].split('.')
                    if(len(str1)>1):
                        str_postfix=str1[len(str1)-1]
                        str2=str1[0].split('/')
                        str_filename=unquote(str2[2]).decode('utf-8')
                        #sourceFilename= localDir + str_filename + '.' + str_postfix
                        sourceFilename= str_filename + '.' + str_postfix
                        filename= localDir + str_orderSerialNo + '.' + str_postfix
                        print sourceFilename
                        print filename
                        #print url_create
                        try:
                            urllib.urlretrieve(url_create, filename, cbk)
                            download_count = 1
                            while(os.path.getsize(filename)<1500 and download_count < 5):
                                urllib.urlretrieve(url_create, filename, cbk)
                            if(os.path.getsize(filename)>1500):
                                print sourceFilename + ': '+ filename + u' 下载完成!'
                            else:
                                print sourceFilename + ': '+ filename + u' 下载错误!'
                        except IOError as e:
                            pass
                    if(len(q1)>1):
                        for j in range(1,len(q1)):
                            url_create = url_base + q1[j]
                            str1=q1[j].split('.')
                            if(len(str1)>1):
                                str_postfix=str1[len(str1)-1]
                                str2=str1[0].split('/')
                                str_filename=unquote(str2[2]).decode('utf-8')
                                sourceFilename= str_filename + '.' + str_postfix
                                filename= localDir + str_orderSerialNo + '(' + str(j) + ')' + '.' + str_postfix                                  
                                try:
                                    urllib.urlretrieve(url_create, filename,cbk)
                                    download_count = 1
                                    while(os.path.getsize(filename)<1500 and download_count < 5):
                                        urllib.urlretrieve(url_create, filename, cbk)
                                    if(os.path.getsize(filename)>1500):
                                        print sourceFilename + ': '+ filename + u' 下载完成!'
                                    else:
                                        print sourceFilename + ': '+ filename + u' 下载错误!'
                                except IOError as e:
                                    pass
                else:
                    print str_orderSerialNo + ': '+ u'SG工单无附件！'
                    f.write(str_orderSerialNo + '\n')
            f.close()
        print u'所有附件下载完成!'
        str_select = '是否结束？按回车键退出，按任意键返回'
        select = raw_input(str_select.decode('utf-8').encode('gb2312'))


    elif(int(select)==2 or int(select)==3):
        #退单需输入退单原因
        if(int(select)==3):
            str_str2 = '请输入退单原因:'
            reason = raw_input(str_str2.decode('utf-8').encode('gb2312'))

        f = open(u'D:\\SG工单回单失败清单.txt','a')
        f.truncate()

        #输入工单号
        str_str1 = '请输入工单号（用,隔开）:'
        sgSerialNoList = raw_input(str_str1.decode('utf-8').encode('gb2312'))
        print u'请等待...'
        s6 = sgSerialNoList.split(',')
        n1 = len(s6)

        #设定每页展示的工单数量
        postdata_pageNumber=urllib.urlencode({
            'pagination.index':'1',
            'pagination.size':str_sgNumber
        })
        req_page = urllib2.Request(
            headers = headers,
            url = 'http://132.121.92.224:7001/WebSg/listAllWorkFromMiddQuery.action?srcPage=allWorkArea',
            data = postdata_pageNumber
        )
        response1 =opener.open(req_page)
        page_print_1 = response1.read()
        reg_t=re.compile(r'\length\]\s+\=\s+\[([^\]]*)\]')
        q3=reg_t.findall(page_print_1)

        #创建list存放字段信息
        datalist0 = []
        datalist1 = []
        datalist2 = []
        datalist3 = []
        datalist4 = []
        datalist5 = []
        datalist6 = []
        datalist7 = []
        datalist8 = []
        datalist9 = []
        datalist10 = []
        datalist11 = []
        datalist12 = [] #子工单/正常

        for i in range(0,str_sgNumber):
            s1=q3[i].split('\'')
            s2=s1[5]
            s3=s2.split(',')
            str_taskName=quote(s3[0]) #网管验收
            str_taskID=quote(s3[1]) #TaskID
            str_orderTicketId=s3[3]
            str_workTicketId=s3[2]
            str_orderSerialNo=s3[4]
            str_tacheInstId=s3[5]
            str_flowInstId=s3[6]
            str_workDirection=quote(s3[7])
            str_sgWorkListInfoId=s3[8]
            str_codeCount=s3[9]
            str_childOrderCount=s3[10]
            str_sgWorkTicketAttrValuesTopic=quote(s1[17])
            str_findZigongdan=s1[9]

            datalist0.append(quote(s3[0]))
            datalist1.append(quote(s3[1]))
            datalist2.append(s3[2])
            datalist3.append(s3[3])
            datalist4.append(s3[4])
            datalist5.append(s3[5])
            datalist6.append(s3[6])
            datalist7.append(quote(s3[7]))
            datalist8.append(s3[8])
            datalist9.append(s3[9])
            datalist10.append(s3[10])
            datalist11.append(quote(s1[17]))
            datalist12.append(s1[9])
        
        #根据SG工单号查找
        for j in range(0,n1):
            sgSerial = s6[j]
            if(sgSerial in datalist4):
                index_sg = datalist4.index(sgSerial)
                #构造post
                str_taskName = datalist0[index_sg]
                str_taskID = datalist1[index_sg]
                str_orderTicketId = datalist3[index_sg]
                str_workTicketId = datalist2[index_sg]
                str_orderSerialNo = datalist4[index_sg]
                str_tacheInstId = datalist5[index_sg]
                str_flowInstId = datalist6[index_sg]
                str_workDirection = datalist7[index_sg]
                str_sgWorkListInfoId = datalist8[index_sg]
                str_codeCount = datalist9[index_sg]
                str_childOrderCount = datalist10[index_sg]
                str_sgWorkTicketAttrValuesTopic = datalist11[index_sg]
                str_findZigongdan = datalist12[index_sg]

##                print 'str_taskName:' + str_taskName
##                print 'str_taskID:' + str_taskID
##                print 'str_orderTicketId:' + str_orderTicketId
##                print 'str_workTicketId:' + str_workTicketId
##                print 'str_orderSerialNo:' + str_orderSerialNo
##                print 'str_tacheInstId:' + str_tacheInstId
##                print 'str_flowInstId:' + str_flowInstId
##                print 'str_workDirection:' + str_workDirection
##                print 'str_sgWorkListInfoId:' + str_sgWorkListInfoId
##                print 'str_codeCount:' + str_codeCount
##                print 'str_childOrderCount:' + str_childOrderCount
##                print 'str_findZigongdan:' + str_findZigongdan

                postdata_download=urllib.urlencode({
                    'taskName':str_taskName,
                    'taskId':str_taskID,
                    'orderTicketId':str_orderTicketId,
                    'workTicketId':str_workTicketId,
                    'orderSerialNo':str_orderSerialNo,
                    'tacheInstId':str_tacheInstId,
                    'flowInstId':str_flowInstId,
                    'workDirection':str_workDirection,
                    'sgWorkListInfoId':str_sgWorkListInfoId,
                    'codeCount':str_codeCount,
                    'childOrderCount':str_childOrderCount
                })
                req_cl =urllib2.Request(
                    headers = headers,
                    url = 'http://132.121.92.224:7001/WebSg/infoAllWork.action?srcPage=allWorkArea',
                    data = postdata_download
                )
                response3 =opener.open(req_cl)
                #判断“单据操作监控”中是否带有“子工单”字样(此法不适用，已修改）
                response5=opener.open('http://132.121.92.224:7001/WebSg/listExecuteLog.action?serviceOrderTicketId='+str_orderTicketId)
                page_print_3 = response5.read()
                #如果是子工单：
                if(len(str_findZigongdan)==29):
                    data1 = 'sgWorkTicket.sgWorkTicketAttrValuesType.topic=1&sgWorkTicket.sgWorkTicketAttrValues.topic=&'
                    data2 = 'sgWorkTicket.sgWorkTicketAttrValuesType.projectNo=1&sgWorkTicket.sgWorkTicketAttrValues.projectNo=&'
                    data3 = 'sgWorkTicket.sgWorkTicketAttrValuesType.serialno=1&sgWorkTicket.sgWorkTicketAttrValues.serialno=&'
                    data4 = 'sgWorkTicket.exceptionReason=&sgWorkTicket.sgWorkTicketAttrValuesType.requireCheckDate=1&'
                    data5 = 'sgWorkTicket.sgWorkTicketAttrValues.requireCheckDate=&sgWorkTicket.sgWorkTicketAttrValuesType.xtysyqwcsj=1&'
                    data6 = 'sgWorkTicket.sgWorkTicketAttrValues.xtysyqwcsj=&sgWorkTicket.sgWorkTicketAttrValuesType.submitPerson_sz=1&'
                    data7 = 'sgWorkTicket.sgWorkTicketAttrValues.submitPerson_sz='
                    data8 = loginUsername + '&'
                    data9 = 'sgWorkTicket.sgWorkTicketAttrValuesType.submitPersonPhone_sz=1&sgWorkTicket.sgWorkTicketAttrValues.submitPersonPhone_sz=&'
                    data10 = 'sgWorkTicket.sgWorkTicketAttrValuesType.feedbackinfo=1&sgWorkTicket.sgWorkTicketAttrValues.feedbackinfo=&'
                    data11 = 'taskId=' + '' + '&'
                    data12 = 'workTicketId=' + str_workDirection + '&'
                    data13 = 'workTempletId=&orderTicketId=' + str_orderTicketId +'&'
                    data14 = 'flowInstId=&tacheInstId=' + str_tacheInstId + '&'
                    data15 = 'tacheInstPreTime=&tacheInstAlmTime=&tacheInstFinTime=&workDirection=%D5%FD%CF%F2%D7%F7%D2%B5&'
                    data16 = 'alarmColor=&taskName=%CD%F8%B9%DC%D1%E9%CA%D5&sgWorkListInfoId=' + str_workDirection +'&'
                    data17 = 'serviceLevelId=1&workState=2&dispatchRequire=&configStatus=&'
                    data18 = 'orderSerialNo=' + str_orderSerialNo + '&'
                    data19 = 'parentOrderId=' + str_orderTicketId +'&'
                    data20 = 'ticketType=2&manualProReason=&submitCount=0&needFillOutWorkTicketOvertimeReason=false&'
                    data21 = 'productType=%B9%A4%B3%CC%C8%EB%BF%E2%B9%DC%C0%ED%C1%F7%B3%CC&actionType=%D0%C2%D4%F6&orderType=%D5%FD%B3%A3%B5%A5'
                    data_sg = data1 + data2 + data3 + data4 + data5 + data6 + data7 + data8 + data9 + data10 + data11 + data12 + data13 + data14 + data15 + data16 + data17 + data18 + data19 + data20 + data21
                    if(int(select)==3):
                        req_cm =urllib2.Request(
                            #headers = headers,
                            url = 'http://132.121.92.224:7001/WebSg/backChildOrder.action?tempJson=' + str_workTicketId + '&backOrderReason=' + reason,
                            data = data_sg
                        )
                    else:
                        req_cm =urllib2.Request(
                            #headers = headers,
                            url = 'http://132.121.92.224:7001/WebSg/orderFPDoActionTask.action?tempJson=' + str_workTicketId,
                            data = data_sg
                        )
                    while(True):
                        try:
                            response4 = opener.open(req_cm)
                            page_print_2=response4.read()
                            reg_status_success = re.compile(r'\"contFrame\"\>\s+([^\<]*)\<')
                            reg_status_failure = re.compile(r'\<center\>\s+([^\s]*)\s+\<f')
                            q6 = reg_status_failure.findall(page_print_2)
                            q7 = reg_status_success.findall(page_print_2)
                            if(len(q7)==0):
                                print str_orderSerialNo + q6[0]
                                f.write(str_orderSerialNo + '\n')
                            else:
                                print str_orderSerialNo + q7[0]
                            break
                        except:
                            pass
                #如果是环节工单：
                else:
                    data22='sgWorkTicket.sgWorkTicketAttrValuesType.topic=1&sgWorkTicket.sgWorkTicketAttrValues.topic='
                    data23= str_sgWorkTicketAttrValuesTopic
                    data24='&sgWorkTicket.sgWorkTicketAttrValuesType.projectNo=1&sgWorkTicket.sgWorkTicketAttrValues.projectNo='+ str_orderSerialNo
                    data25='&sgWorkTicket.sgWorkTicketAttrValuesType.serialno=1&sgWorkTicket.sgWorkTicketAttrValues.serialno='+ str_orderSerialNo
                    data26='&sgWorkTicket.exceptionReason=346&sgWorkTicket.sgWorkTicketAttrValuesType.requireCheckDate=1'
                    data27='&sgWorkTicket.sgWorkTicketAttrValues.requireCheckDate=&sgWorkTicket.sgWorkTicketAttrValuesType.xtysyqwcsj=1'
                    data28='&sgWorkTicket.sgWorkTicketAttrValues.xtysyqwcsj=&sgWorkTicket.sgWorkTicketAttrValuesType.submitPerson_sz=1'
                    data29='&sgWorkTicket.sgWorkTicketAttrValues.submitPerson_sz='+loginUsername+'&sgWorkTicket.sgWorkTicketAttrValuesType.submitPersonPhone_sz=1'
                    data30='&sgWorkTicket.sgWorkTicketAttrValues.submitPersonPhone_sz=&sgWorkTicket.sgWorkTicketAttrValuesType.feedbackinfo=1'
                    data31='&sgWorkTicket.sgWorkTicketAttrValues.feedbackinfo=&taskId='+str_taskID
                    data32='&workTicketId='+str_workTicketId+'&workTempletId=&orderTicketId='+str_orderTicketId+'&flowInstId='+str_flowInstId+'&tacheInstId='+str_tacheInstId
                    data33='&tacheInstPreTime=&tacheInstAlmTime=&tacheInstFinTime=&workDirection=%D5%FD%CF%F2%D7%F7%D2%B5&alarmColor=&taskName=%B9%A4%B3%CC%CF%B5%CD%B3%D1%E9%CA%D5'
                    data34='&sgWorkListInfoId='+str_workDirection+'&serviceLevelId=&workState=2&dispatchRequire=&configStatus=&orderSerialNo='+str_orderSerialNo
                    data35='&parentOrderId='+str_orderTicketId+'&ticketType=&manualProReason=&submitCount=0&needFillOutWorkTicketOvertimeReason=false'
                    data36='&productType=%B9%A4%B3%CC%C8%EB%BF%E2%B9%DC%C0%ED%C1%F7%B3%CC&actionType=%D0%C2%D4%F6&orderType=%D5%FD%B3%A3%B5%A5'
                    data37='&sgWorkTicket.exceptionReason=&sgWorkTicket.sgWorkTicketAttrValuesType.requireCheckDate=1'
                    data_sg1= data22 + data23 + data24 + data25 + data26 + data27 + data28 + data29 + data30 + data31 + data32 + data33 + data34 + data35 + data36
                    data_sg2= data22 + data23 + data24 + data25 + data37 + data27 + data28 + data29 + data30 + data31 + data32 + data33 + data34 + data35 + data36
                    if(int(select)==3):
                        req_cm =urllib2.Request(
                            #headers = headers,
                            url = 'http://132.121.92.224:7001/WebSg/doWork.action?srcPage=allWorkArea&delayTime=5&manualProReason=',
                            data = data_sg1
                        )
                    else:
                        req_cm =urllib2.Request(
                    	    #headers = headers,
                            url = 'http://132.121.92.224:7001/WebSg/doWork.action?srcPage=allWorkArea&delayTime=5&manualProReason=',
                            data = data_sg2
                        )
                    while(True):
                        try:
                            response4 = opener.open(req_cm)
                            page_print_2 =response4.read()
                            reg_status_success = re.compile(r'\"contFrame\"\>\s+([^\<]*)\<')
                            reg_status_failure = re.compile(r'\<center\>\s+([^\s]*)\s+\<f')
                            q6 = reg_status_failure.findall(page_print_2)
                            q7 = reg_status_success.findall(page_print_2)
                            if(len(q7)==0):
                                print str_orderSerialNo + q6[0]
                                f.write(str_orderSerialNo + '\n')
                            else:
                                print str_orderSerialNo + q7[0]
                            break
                        except:
                            pass
            else:
                print sgSerial + u'工单已回或不存在！'
        print u'所有工单回单完成！'
        f.close()
        str_select = '是否结束？按回车键退出，按任意键返回'
        select = raw_input(str_select.decode('utf-8').encode('gb2312'))
    else:
        print u'输入错误！'
        str_select = '是否结束？按回车键退出，按任意键返回'
        select = raw_input(str_select.decode('utf-8').encode('gb2312'))

