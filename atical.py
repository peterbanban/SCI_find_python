# _*_ coding:utf-8 _*_
#GVIM 添加汉字注释

import sys
reload(sys)
sys.setdefaultencoding('utf-8')

#以上定义字符集 使得代码可用汉字

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
row0=['ISSN','期刊名','影响因子','中科院分区','大类学科','小类学科','SCI/SCIE','是否分区','录用比','审稿周期','查看数','研究方向']

###创建excel文件并写入首行
f=xlwt.Workbook()
sheet1=f.add_sheet('sheet1',cell_overwrite_ok=True)
for i in range(1,2):
    sheet1.write(0,i,row0[i].decode('GB2312'))

###遍历所有研究方向
for k in range(2,3):                              
    url='http://www.letpub.com.cn/index.php?page=journalapp&fieldtag='+str(k)+'&firstletter=&currentpage=1#journallisttable'
    headers = {'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}
    request=urllib2.Request(url,headers=headers)
    response=urllib2.urlopen(request,timeout=40)
    result=response.read()
    html=BeautifulSoup(result,'lxml')
    try:
        title=html.find_all('h2')[2]                #此处异常处理是为了防止反扒中的格式错乱问题
    except:
        continue
    title=title.get_text()

    ####取某研究方向的页数
    pages=title[-4:-1]                     
    pages=title.split('：'.decode('GB2312'))[-1]
    pages=pages[:-1]
    
    ###以下三句输出为了方便查看处理进度
    print pages
    print title.encode('GBK','ignore')
    print k
    data=int(pages)
    if data==0:                                    #此处为了跳过空白页
        continue
 
   ###遍历每个研究方向的所有页
    for i in range(1,data/10+2):
        url='http://www.letpub.com.cn/index.php?page=journalapp&fieldtag='+str(k)+'&firstletter=&currentpage='+str(i)+'#journallisttable'
        headers = {'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'}    #此句为反扒措施，仿浏览器请求数据
        request=urllib2.Request(url,headers=headers)
        response=urllib2.urlopen(request,timeout=40)
        result=response.read()
        html=BeautifulSoup(result,'lxml')
        row_tr =html.find_all('tr')                               #以tr标签格式化请求得到的htm文件，得到元素为每个tr组成的列表
       
        ###遍历每页的所有行，即标签tr
        for tr in range(7,len(row_tr)-1):
            row_td=row_tr[tr].find_all('td')                      #以td标签格式化每个tr的内容，得到元素为每个td组成的列表
            rows=[]
            rows.append(row_td[0].get_text())                      #单独处理 0,1,11个单元格，其他单元格形式相同一起处理
            rows.append(row_td[1].find_all('a')[0].get_text())
            
            ###遍历每行的所有列,即标签td，暂时放入rows里
            for i in range(2,10):
                rows.append(row_td[i].get_text())
            rows.append(row_td[11].get_text())
            title=title.split('，'.decode('GB2312'))[0]            #将研究方向几个字从title里分离出来放在rows里以便一起写入表格
            rows.append(title[5:])
            
            ###将遍历所得值写入表格
            for j in range(0,12):
                sheet1.write(num,j,rows[j])            
            print num
            num=num+1 
    f.save('demo5.xls')      #保存excel文件


