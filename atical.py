# _*_ coding:utf-8 _*_
#GVIM ��Ӻ���ע��

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

#���϶����ַ��� ʹ�ô�����ú���

import json
import urllib2
import urllib
import xlrd
import xlwt 
import os
from bs4 import BeautifulSoup
import httplib  
  
httplib.HTTPConnection._http_vsn = 10  
httplib.HTTPConnection._http_vsn_str = 'HTTP/1.0'

num=1
row0=['ISSN','�ڿ���','Ӱ������','�п�Ժ����','����ѧ��','С��ѧ��','SCI/SCIE','�Ƿ����','¼�ñ�','�������','�鿴��','�о�����']

###����excel�ļ���д������
f=xlwt.Workbook()
sheet1=f.add_sheet('sheet1',cell_overwrite_ok=True)
for i in range(1,2):
    sheet1.write(0,i,row0[i].decode('GB2312'))

###���������о�����
for k in range(2,3):                              
    url='http://www.letpub.com.cn/index.php?page=journalapp&fieldtag='+str(k)+'&firstletter=&currentpage=1#journallisttable'
    headers = {'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}
    request=urllib2.Request(url,headers=headers)
    response=urllib2.urlopen(request,timeout=40)
    result=response.read()
    html=BeautifulSoup(result,'lxml')
    try:
        title=html.find_all('h2')[2]                #�˴��쳣������Ϊ�˷�ֹ�����еĸ�ʽ��������
    except:
        continue
    title=title.get_text()

    ####ȡĳ�о������ҳ��
    pages=title[-4:-1]                     
    pages=title.split('��'.decode('GB2312'))[-1]
    pages=pages[:-1]
    
    ###�����������Ϊ�˷���鿴�������
    print pages
    print title.encode('GBK','ignore')
    print k
    data=int(pages)
    if data==0:                                    #�˴�Ϊ�������հ�ҳ
        continue
 
   ###����ÿ���о����������ҳ
    for i in range(1,data/10+2):
        url='http://www.letpub.com.cn/index.php?page=journalapp&fieldtag='+str(k)+'&firstletter=&currentpage='+str(i)+'#journallisttable'
        headers = {'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}    #�˾�Ϊ���Ǵ�ʩ�����������������
        request=urllib2.Request(url,headers=headers)
        response=urllib2.urlopen(request,timeout=40)
        result=response.read()
        html=BeautifulSoup(result,'lxml')
        row_tr =html.find_all('tr')                               #��tr��ǩ��ʽ������õ���htm�ļ����õ�Ԫ��Ϊÿ��tr��ɵ��б�
       
        ###����ÿҳ�������У�����ǩtr
        for tr in range(7,len(row_tr)-1):
            row_td=row_tr[tr].find_all('td')                      #��td��ǩ��ʽ��ÿ��tr�����ݣ��õ�Ԫ��Ϊÿ��td��ɵ��б�
            rows=[]
            rows.append(row_td[0].get_text())                      #�������� 0,1,11����Ԫ��������Ԫ����ʽ��ͬһ����
            rows.append(row_td[1].find_all('a')[0].get_text())
            
            ###����ÿ�е�������,����ǩtd����ʱ����rows��
            for i in range(2,10):
                rows.append(row_td[i].get_text())
            rows.append(row_td[11].get_text())
            title=title.split('��'.decode('GB2312'))[0]            #���о����򼸸��ִ�title������������rows���Ա�һ��д����
            rows.append(title[5:])
            
            ###����������ֵд����
            for j in range(0,12):
                sheet1.write(num,j,rows[j])            
            print num
            num=num+1 
    f.save('demo5.xls')      #����excel�ļ�


