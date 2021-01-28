#!/usr/bin/env python
# coding: utf-8

# In[453]:


import os
import re
import time
import yaml
import random
import logging
import xlwt
import xlrd
import xlsxwriter


# In[454]:


#安装selenium pip install selenium


# In[455]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException


# In[460]:


def TitleSpiderMain(papertitle):
    #本函数用于根据题名搜索进行信息收集，打开文章详细页进行收集
    #信息包含文章基本信息和期刊基本信息，如果有的话
    #返回两个结果，一个为查询信息结果，一个是未查询到文章列表。
    mylist=[]
    nofind=[]
    driver = webdriver.Chrome()
    url = 'http://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&SID=6FAEOvziD7rmWrbUJk6&search_mode=GeneralSearch'
    driver.implicitly_wait(20)#等待时间
    driver.get(url)
    driver.find_element_by_id("clearIcon1").click()
    driver.find_element_by_id("value(input1)").send_keys(papertitle)#模拟输入数字直接数字，文本需要双引号
    s1 = Select(driver.find_element_by_xpath('//*[@id="select1"]'))#选择检索主题
    s1.select_by_visible_text('标题')#常用 主题，标题，作者，地址
    driver.find_element_by_id("searchCell1").click()#模拟点击检索按钮，经常更换，需更新
    try:
        driver.find_element_by_xpath('//*[@id="noRecordsDiv"]/div[1]')
        nofind.append(papertitle)
    except:
        papernumber=driver.find_element_by_id("hitCount\.top").text#搜索文章数
        pagenumber=driver.find_element_by_id("pageCount\.top").text#选择页码数
        for i in range(1,int(papernumber)+1):
                path='//*[@id="RECORD_'+str(i)+'"]/div[3]/div/div[1]/div/a/value'
                driver.find_element_by_xpath(path).click()#打开文档详细页ID 更改条数
                driver.find_element_by_xpath('//*[@id="hidden_section_label"]').click()#文献详细页显示更多信息
                infor1=driver.find_element_by_xpath('//*[@id="records_form"]/div/div/div/div[1]/div').text#文献详细页内容
                time.sleep(1)
                infor1=infor1+driver.find_element_by_xpath('//*[@id="hidden_section"]').text#添加隐藏信息
                time.sleep(2)
                path='//*[@id="show_journal_overlay_link_'+str(i)+'"]/p/a'
                try:
                    driver.find_element_by_xpath(path).click()#打开期刊页面
                    path='//*[@id="show_journal_overlay_'+str(i)+'"]'
                    infor2=driver.find_element_by_xpath(path).text#提取期刊信息
                    time.sleep(2)
                    list1 = [infor1,infor2]
                    infor='\n'.join(list1)
                    mylist.append(infor)
                    driver.find_element_by_xpath('//*[@id="skip-to-navigation"]/ul[1]/li[2]/a').click()#返回检索页面
                except:
                    infor=infor1
                    mylist.append(infor)
                    driver.find_element_by_xpath('//*[@id="skip-to-navigation"]/ul[1]/li[2]/a').click()#返回检索页面
                #driver.find_element_by_xpath('//*[@id="skip-to-navigation"]/ul[1]/li[1]/a').click()    
    time.sleep(5)
    driver.close()                    
    return mylist,nofind  


# In[457]:


def data_write(file_path, datas):
    #列表保存成xls文件，需要import xlwt
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet    
    #将数据写入第 i 行，第 j 列
    i = 0
    for data in datas:
        for j in range(len(data)):
            sheet1.write(i,j,data[j])
        i = i + 1        
    f.save(file_path) #保存文件


# In[458]:


def paper_info_take(paperinfo):
    #从记录中提取文章信息，输入的为一条记录    
    title=paperinfo.split('\n',)[0]#题名    
    try:
        paperinfo.find('查看 Web of Science ResearcherID 和 ORCID')>-1
        journal=paperinfo[paperinfo.find('查看 Web of Science ResearcherID 和 ORCID'):].split('\n',)[1]#journal
    except:
        journal=paperinfo.split('\n',)[2]#journal
    try:
        paperinfo.find('作者:')>-1
        author=paperinfo[paperinfo.find('作者:'):].split('\n',)[0]# author
    except:
        author=''    
    try:
        paperinfo.find('卷:')>-1
        volume=paperinfo[paperinfo.find('卷:'):].split('\n',)[0]#卷期页
    except:
        volume=''
    try:
        paperinfo.find('DOI:')>-1
        doi=paperinfo[paperinfo.find('DOI:'):].split('\n',)[0]#doi
    except:
        doi=''
    try:
        paperinfo.find('出版年:')>-1
        year=paperinfo[paperinfo.find('出版年:'):].split('\n',)[0]#出版年
    except:
        year=''
    try:
        paperinfo.find('作者关键词:')>-1
        keyword=paperinfo[paperinfo.find('作者关键词:'):].split('\n',)[0]#关键词
    except:
        keyword=''
    try:
        paperinfo.find('摘要')>-1
        abstract=paperinfo[paperinfo.find('摘要'):].split('\n',)[1]#出版年
    except:
        abstract=''
    try:
        paperinfo.find('(通讯作者)')>-1
        institution=paperinfo[paperinfo.find('通讯作者地址:'):paperinfo.find('(通讯作者)')]#机构#通讯作者
        #Cauthor=paperinfo[paperinfo.find('通讯作者地址:'):].split('\n',)[2]
    except:
        institution=''
        #Cauthor=''
    try:
        paperinfo.find('入藏号:')>-1
        Collectionnumber=paperinfo[paperinfo.find('入藏号:'):].split('\n',)[0]#馆藏号
    except:
        Collectionnumber=''
    try:
        paperinfo.find('impact factor')>-1
        impactfactor=paperinfo[paperinfo.find('impact factor'):].split('\n',)[1]#影响因子
    except:
        impactfactor=''   
    try:
        paperinfo.find('ISSN:')>-1
        ISSN=paperinfo[paperinfo.find('ISSN:'):].split('\n',)[1]#影响因子
    except:
        ISSN=''  
    try:
        paperinfo.find('Web of Science 核心合集中的 "被引频次":')>-1
        WOFcite=paperinfo[paperinfo.find('Web of Science 核心合集中的 "被引频次":'):].split('\n',)[0]#被引次数
    except:
        WOFcite=''
    try:
        paperinfo.find('类别中的排序 JCR 分区')>-1
        JCR=paperinfo[paperinfo.find('类别中的排序 JCR 分区'):].split('\n',)[1]#jcr分区
    except:
        JCR=''
    try:
        paperinfo.find('授权号')>-1
        Fund=paperinfo[paperinfo.find('授权号'):paperinfo.find('查看基金资助信息')]
    except:
        Fund=''
    paperinformation=title+'$'+journal+'$'+author+'$'+volume+'$'+doi+'$'+year+'$'+keyword+'$'+abstract+'$'+Collectionnumber+'$'+impactfactor+'$'+ISSN+'$'+WOFcite+'$'+JCR+'$'+Fund+'$'+institution
    return paperinformation
    
    


# In[461]:


#主程序部分
#导入要搜索的文章信息

data = xlrd.open_workbook('2020文章列表.xlsx')#导入文章列表
table = data.sheet_by_name('paper')#具体是取哪个表格，填写名称
nrows = table.nrows
ncols = table.ncols
#TitleSpiderMain(table.cell(1, 0).value)
#table.cell(1, 0).value#第一个为行，第二个为列，从0开始计算
result=[]
nofind=[]
for k in range(1,nrows):
    resul,nofind1=TitleSpiderMain(table.cell(k, 0).value)
    result.append(resul)
    nofind.append(nofind1)
#保存原始结果，所有爬取下来的信息
data_write('D:/python/web-of-science-spider-main/resultdataall.xls',result)
data_write('D:/python/web-of-science-spider-main/nofind.xls',nofind)


# In[439]:


####本部分可以进行基本的信息提取，不一定全面


# In[443]:


#文献需要信息提取，生成excel文件
paperinformationAll=[]
for i in range(0,len(result)+1):
    try:
        len(result[i])==1
        paper=paper_info_take(result[i][0])
    except:
        paper=''
    paperinformationAll.append(paper)

#保存结果
workbook = xlsxwriter.Workbook('resultpaperall.xlsx')     #新建excel表 
worksheet = workbook.add_worksheet('sheet1')       #新建sheet（sheet的名称为"sheet1"） 
for i in range(0,len(paperinformationAll)):
    worksheet.write_string('A'+str(i+1),paperinformationAll[i])
workbook.close() 


# In[444]:


result=TitleSpiderMain('Review of Ampharete Malmgren, 1866 (Annelida, Ampharetidae) from China')


# In[446]:


driver = webdriver.Chrome()
url = 'http://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&SID=6FAEOvziD7rmWrbUJk6&search_mode=GeneralSearch'
driver.implicitly_wait(20)#等待时间
driver.get(url)
driver.find_element_by_id("clearIcon1").click()
driver.find_element_by_id("value(input1)").send_keys('Review of Ampharete Malmgren, 1866 (Annelida, Ampharetidae) from China')#模拟输入数字直接数字，文本需要双引号
s1 = Select(driver.find_element_by_xpath('//*[@id="select1"]'))#选择检索主题
s1.select_by_visible_text('标题')#常用 主题，标题，作者，地址
driver.find_element_by_id("searchCell1").click()#模拟点击检索按钮，经常更换，需更新


# In[ ]:




