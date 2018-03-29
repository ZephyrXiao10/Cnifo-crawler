# -*- coding: utf-8 -*-
"""
Created on Wed Feb  7 14:26:01 2018

@author: xiaoziliang_sx
"""




#%%
from WindPy import *
w.start()
import requests
from urllib.parse import urlencode
import xlsxwriter
import itertools
import numpy as np
from datetime import datetime,date,timedelta


URL = 'http://www.cninfo.com.cn/cninfo-new/announcement/query' 
HEADER = {
            'Referer':'http://www.cninfo.com.cn/cninfo-new/announcement/show',
            'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36',
            'X-Requested-With':'XMLHttpRequest'}
RESPONSE_TIMEOUT =10


def get_page(start_date:str,end_date:str,category,max_page = 200):
    
    rp = []
    for page in range(max_page):
        query = {
            'stock':'',
            'searchkey':'',
            'plate':'',
            'category':category,
            'trade':'',
            'column':'sse',
            'columnTitle':'历史公告查询',
            'pageNum':page+1,
            'pageSize':30,
            'tabName':'fulltext',
            'sortName':'',
            'sortType':'',
            'limit':'',
            'showTitle':'',
            'seDate': start_date+' ~ '+end_date
            }
        try:
            response = requests.post(URL, query, HEADER, timeout=RESPONSE_TIMEOUT)
            if response.status_code == 200:
                rp = rp+response.json().get('announcements')
                
        except requests.ConnectionError as e:
            print(e)
    return rp
    
    
def get_info(json,categoryKey):
    
    for ann in json:
        #adjunctUrl = ann.get('adjunctUrl')
        #print(adjunctUrl)
        if categoryKey == '其它重大事项':
            if '复牌' in ann.get('announcementTitle'):
                category = '复牌公告'
            elif '停牌' in ann.get('announcementTitle'):
                category = '停牌公告'
            else:
                continue
        
        else:
            category = categoryKey        
        
        
        tm = ann.get('announcementTime')/1000
        dtt = str(datetime.fromtimestamp(tm))  
        aid = ann.get('announcementId') 
        view_url = "http://www.cninfo.com.cn/cninfo-new/disclosure/sse/bulletin_detail/true/"+aid+"?announceTime="+dtt[:10]
        

        
        code = ann.get('secCode')
        if code[0] > '3':
            code = code+'.SH'
        else:
            code = code+'.SZ'
        rst = w.wss(code, "sec_name,industry_citic,industry_citiccode","industryType=1")
        print('%s processing...' % ann.get('secName'))
        industry = rst.Data[1][0]
        
        
        #--------第一列
        fir = dtt[:10]
        sec = ann.get('secCode')   
        thi = ann.get('announcementTitle')
        
        fc = fir+'_'+sec+'_'+thi
        
        
        yield {            
            'index':fc,    
            'announcementTime':dtt,
            'secCode':code,
            'secName':ann.get('secName'),
            'industry':industry,
            'announcementTitle': ann.get('announcementTitle'),
            'categoryKey':category,
            'view_url':view_url            
        }

  
    
    
    
#%%
        
#from urllib.request import urlopen
#from pdfminer.pdfinterp import PDFResourceManager, process_pdf
#from pdfminer.converter import TextConverter
#from pdfminer.layout import LAParams
#from io import StringIO
#
#
#
#
#def readPDF(pdfFile):
#    rsrcmgr = PDFResourceManager()
#    retstr = StringIO()
#    laparams = LAParams()
#    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
#
#    process_pdf(rsrcmgr, device, pdfFile)
#    device.close()
#
#    content = retstr.getvalue()
#    retstr.close()
#    return content

    






def ToExcel(info_generator,workbook,op_date):
    worksheet = workbook.add_worksheet('标题整理')
    column_format = workbook.add_format({'bold': False, 'pattern': 1, 'bg_color':'#70ad47','font_name': 'Arial MT','font_color':'white'})
    allformat = workbook.add_format({'bold': False, 'font_name': 'Arial'})
    titleformat = workbook.add_format({'bold': False, 'font_name': 'Arial','text_wrap':True})
    
    red_fmt = workbook.add_format({'bg_color':'#FFC7CE','font_name': 'Arial','font_color': '#9C0006'})
    green_fmt = workbook.add_format({'bg_color':'#C6EFCE', 'font_name': 'Arial','font_color': '#006100'})                                
    columns_val = np.array(['report_id','公告时间','index_id','index_name','industry','公告标题','公告类型','URL','op_date'])
    worksheet.write_row('A1',columns_val, column_format) 
    worksheet.set_column('A:A', 18)
    worksheet.set_column('B:B', 20.3)
    worksheet.set_column('C:E', 11.75)
    worksheet.set_column('F:F', 80.)
    worksheet.set_column('G:G', 32.7)
    worksheet.set_column('H:H', 90.)
    worksheet.set_column('I:I', 20.3)
    for val_num, info in enumerate(info_generator):
#        if '更正' in info['announcementTitle'] or '修正' in info['announcementTitle'] or '补充' in info['announcementTitle']:
#            worksheet.write(val_num+1,0,info['secCode'], allformat)
#            worksheet.write(val_num+1,1,info['secName'], red_fmt)
#            worksheet.write(val_num+1,2,info['announcementTitle'], allformat)      
#        if '停牌' in info['announcementTitle']:
#            val_str = np.array([info['announcementTime'],info['secCode'],info['secName'],'停牌公告'])
#            worksheet.write_row(val_num+1,0,val_str,red_fmt)
#        elif '复牌' in info['announcementTitle']:
#            val_str = np.array([info['announcementTime'],info['secCode'],info['secName'],'复牌公告'])
#            worksheet.write_row(val_num+1,0,val_str,green_fmt)            
            
        val_str = np.array([info['index'],info['announcementTime'],info['secCode'],info['secName'],info['industry'],info['announcementTitle'],info['categoryKey'],info['view_url']])
        worksheet.write_row(val_num+1,0,val_str,allformat)  
        worksheet.write(val_num+1,8,op_date,allformat)
    return workbook






if __name__ == '__main__':
    outputStringList = []
    MAX_PAGE = 30
    # START_DATE END_DATE = "yyyy-MM-dd"   
    # please note that END_DATE could be 1 or 2 days ahead of currrent date
    
#    today = str(date.today())
#    tomorrow = str(date.today()+timedelta(1))
#    START_DATE = '2018-01-01'
#    END_DATE = today
#    tradedays=w.tdays(START_DATE, END_DATE, "")  
#    yestertradep1 = str(tradedays.Data[0][-2]+timedelta(1))[:10]
#    op_date = str(datetime.now())[:19]
    
    today = str(date.today()) 
    tomorrow = str(date.today()+timedelta(1))
    END_DATE = today
    op_date = str(datetime.now())[:19]
    
    
    
    #print("start date:%s----------end date: %s" %(yestertradep1,tomorrow))
    print("start date:%s----------end date: %s" %(today,tomorrow))
    cat_list = ['category_qyfpxzcs_szsh;',
                'category_bcgz_szsh;',
                'category_cqfxyj_szsh;',
                'category_qtzdsx_szsh;']
    cat_key = ['权益分派与限制出售股份上市','补充及更正','澄清、风险提示、业绩预告事项','其它重大事项']
    
    #wkbk = xlsxwriter.Workbook('巨潮业绩预告爬虫标题整理'+str(END_DATE)+'.xlsx')    
    
    it=[]
    for i,cat in enumerate(cat_list):
        print('category %s processing...' % cat_key[i])
        r = get_page(today,tomorrow,cat,max_page=100)
        info_generator = get_info(r,cat_key[i])   
        it = itertools.chain(it,info_generator)
        
    wkbk = xlsxwriter.Workbook('巨潮公告爬虫整理'+str(END_DATE)+'.xlsx')
    wkbk = ToExcel(it, wkbk, op_date)    
    wkbk.close()  
     
    
    

    
######### pdf reader
#    for info in info_generator:
#        pdf_url = 'http://www.cninfo.com.cn/'+ info.get('adjunctUrl')
#        print(pdf_url)
#        pdfFile = urlopen(pdf_url)
#        outputString = readPDF(pdfFile)
#        outputStringList.append(outputString)
#        secName= info['secName'].replace('*','s')
#        with open('./业绩预告/'+info['secCode']+secName+info['announcementTitle']+'.txt', 'w',encoding='utf8') as f:
#            f.write(outputString)        
#        print(outputString)
#            
#    pdfFile.close()  
        
    