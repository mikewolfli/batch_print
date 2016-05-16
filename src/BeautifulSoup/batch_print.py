#coding=utf-8
'''
Created on 2016年1月11日

@author: 10256603
'''

import os
import sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox 
from tkinter import scrolledtext
from tkinter.font import Font
import psycopg2
import logging
from suds.transport.http import *
import urllib
import gzip 
from bs4 import BeautifulSoup
from ntlm import HTTPNtlmAuthHandler
import clr
import configparser
#from System.Reflection import Assembly
#dllpath=os.path.join(os.getcwd(),'Seagull.BarTender.Print.dll')
#Assembly.LoadFile(dllpath)
clr.AddReference("Seagull.BarTender.Print")
import  Seagull.BarTender.Print as printer

id_pre='ctl00_PlaceHolderMain_dataListCELabel_ctl01_lb'
dic_name1=['ProjectName','WBSNo','ElevatorNo','CarWeight','CwtFrameWeight','ReserveDecorationWeight',
          'CarBalanceBlockQty','CarWeightBlockQty','BalanceRate','CwtBlockQtyBeforeDc','CwtBlockQtyAfterDc',
          'CwtBlockQtyMat1','CwtBlockQtyMat2']
dic_name2=['TmBrakeSpringLeft','TmBrakeSpringRight','BufferProductionNoCar',
          'BufferProductionNoCw','GovernorProductionNoCar','GovernorProductionNoCw','SaftyGearProductionNoCar',
          'SaftyGearProductionNoCw','MotherBoardProductionNo','TmProductionNo']

head_name={'TmBrakeSpringLeft':'曳引机制动弹簧 左',
           'TmBrakeSpringRight':'曳引机制动弹簧 右',
           'BufferProductionNoCar':'缓冲器制造编号 轿厢',
           'BufferProductionNoCw':'缓冲器制造编号 对重',
           'GovernorProductionNoCar':'限速器制造编号 轿厢',
           'GovernorProductionNoCw':'限速器制造编号 对重',
           'SaftyGearProductionNoCar':'安全钳制造编号 轿厢',
           'SaftyGearProductionNoCw':'安全钳制造编号 对重',
           'MotherBoardProductionNo':'主板制造编号',
           'TmProductionNo':'曳引机制造编号'}
force_str='请见合格证'
none_str='/'

class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert(END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)# Scroll to the bottom
            
       
class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()
        try:
            self.conn=psycopg2.connect(database="pgworkflow", user="postgres", password="1q2w3e4r", host="10.127.144.62", port="5432")
        except :
            messagebox.showerror(title="连接错误", message="无法连接")
            sys.exit()
            
        
    def quit_func(self):
        if self.conn is not None:
            self.conn.close()
        self.quit()
        
    def get_governor_flag(self,wbs):
        cur_unit=self.conn.cursor()
        cur_unit.execute("select has_governor from s_unit_info_attach where wbs_no='"+wbs+"';")
        row_unit = cur_unit.fetchall()
        
        if not row_unit:
            return None
        
        return row_unit[0][0]   

    def get_info_table(self, wbs):
        config=configparser.ConfigParser()
        config.read(os.path.join(os.getcwd(),'login.ini'))
        self.user_id=config.get('login','userid')
        password=config.get('login','password')  
    
        #用户域验证
        url='http://spcn.tkeasia.com/sites/chinamfg/CELabel/_layouts/15/CELabel/SearchPage.aspx'  
        passman = urllib.request.HTTPPasswordMgrWithDefaultRealm()
        passman.add_password(None, url, self.user_id, password)
        # create the NTLM authentication handler
        auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
        #cj=http.cookiejar.CookieJar()  
        #ck_hander=urllib.request.HTTPCookieProcessor(cj) 
        opener=urllib.request.build_opener(auth_NTLM)
        urllib.request.install_opener(opener)    
        req=urllib.request.Request(url)    
        resp=urllib.request.urlopen(req)
             
        if resp.code==401:
            self.logger.error("登陆失败，请更新login.ini中的账号密码执行")
            return None            
        
        data_pool=[]
        
        self.logger.info("正在网页抓却信息表数据...")
        for wbs_item in wbs:
            data={'MSOGallery_FilterString':'',
                  'MSOGallery_SelectedLibrary':'',
                  'MSOLayout_InDesignMode':'',
                  'MSOLayout_LayoutChanges':'',
                  'MSOSPWebPartManager_DisplayModeName':'Browse',
                  'MSOSPWebPartManager_EndWebPartEditing':'false',
                  'MSOSPWebPartManager_ExitingDesignMode':'false',
                  'MSOSPWebPartManager_OldDisplayModeName':'Browse',
                  'MSOSPWebPartManager_StartWebPartEditingName':'false',
                  'MSOTlPn_Button':'none',
                  'MSOTlPn_SelectedWpId':'',
                  'MSOTlPn_ShowSettings':'False',
                  'MSOTlPn_View':'0',
                  'MSOWebPartPage_PostbackSource':'',
                  'MSOWebPartPage_Shared':'',
                  '__EVENTARGUMENT':'',
                  '__EVENTTARGET':'',
                  '__EVENTVALIDATION':'/wEdABI9aNyskv96U2YHfKrkWRos0mdSE9/ejYYPXEQDvuz8jTE2ZPaixp3A4yJqsaaZ9TNw93j7+RM9oh0dUTsxL+1jtQbDR+ch1L1p4GNAeaLPS7CfyGGwdgKwbM6Oy0Q3Vwk+TMgdQTOIOcxEfagsXAJAa4Szvm9LsC6IPHBGqLlB1FfFbbEIk3I4Ch5NNWRxduxjAwI5raT+ZEUkah07GrWvfTuITGYKWYvzyTpH8j/veqtHEO0ZArytCDynmC10JyV216uA9/JUvxWbS/N6JcqvBmFgmF8YCyGBQWrzffUx4dRsRXpYAxeRMDyvNBUaswyx2ZFXhiAzFTaaG275wiB9F675wgOL/edqy+XoTRmB4Bvof46P+dv4BBRgmctafO2rH5BYJVrcWgrp/9HygbN9puM1yacwhDwdEBwseT1PbQ==',
                  #'__REQUESTDIGEST':'0x0904B9DB5BAEFAE0CA2F8A013678D173C272DA79D860D4784A30242AB10D6328B19210F732B1024CFF05526FCD1179D253EE2473CF4AB717BFA5694261BA73AE,28 Jan 2016 01:22:52 -0000',
                  '__SCROLLPOSITIONX':'0',
                  '__SCROLLPOSITIONY':'0',
                  '__VIEWSTATE':'/wEPBSpWU0tleTo5NzEzMTEzMi04YjgyLTRjNDAtYTg1OS03NjBiMzQzNTM0NzJkjy8vgwgj8rMCnkqnvoK8FNbj3NqYV4eUPbnwSqDJOA8=',
                  '_maintainWorkspaceScrollPosition':'0',
                  'ctl00$PlaceHolderMain$AspNetPagerCELabel_input':'1',
                  'ctl00$PlaceHolderMain$DropDownListPageSize':'10',
                  'ctl00$PlaceHolderMain$HiddenField':self.user_id,
                  'ctl00$PlaceHolderMain$btnSearch':'查询',
                  'ctl00$PlaceHolderMain$ddlStatus':'9',
                  'ctl00$PlaceHolderMain$tbProjectName':'',
                  'ctl00$PlaceHolderMain$tbWBSNo':wbs_item,}
            postdata=urllib.parse.urlencode(data, 'utf-8')
            binary_data = postdata.encode('utf-8')
            req=urllib.request.Request(url,binary_data)
            
            req.add_header('Host', 'spcn.tkeasia.com')
            req.add_header('User-Agent','Mozilla/5.0 (Windows NT 6.1; rv:43.0) Gecko/20100101 Firefox/43.0')
            req.add_header('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8')
            req.add_header('Accept-Language','en-US,en;q=0.5')
            req.add_header('Accept-Encoding','gzip, deflate')
            req.add_header('Referer','http://spcn.tkeasia.com/sites/chinamfg/CELabel/_layouts/15/CELabel/SearchPage.aspx')
            req.add_header('Cookie','WSS_FullScreenMode=false')
            req.add_header('Connection','keep-alive')
                   
            resp=urllib.request.urlopen(req)
            result=gzip.decompress(resp.read()).decode('utf-8')
            data_row=self.parse_xhtml(result)
            
            if not data_row:
                self.logger.error(wbs_item+":数据抓取失败，无法打印信息表标签") 
                continue
            
            data_row['gonvernorflag']=self.get_governor_flag(wbs_item)
            data_pool.append(data_row)
        #print(data_pool)
        self.logger.info("信息表数据抓取完成!")
        return data_pool
            
               
    def parse_xhtml(self, result):
        data={}
        xhtml_content=BeautifulSoup(result)
                
        xhtml_list=xhtml_content.find_all(id=re.compile(id_pre))
        if not xhtml_list:
            return None
        
        dic_name=dic_name1+dic_name2
        for item in dic_name:            
            for line in xhtml_list:
                if line.attrs['id']==id_pre+item:
                    data[item]=line.string
                    break
        return data   
                                  
    def print_wbs(self,wbs):
        btEngine = printer.Engine()
    
        if not btEngine.IsAlive:
            btEngine.Start()
    
        wbs_file_path = os.path.join(os.getcwd(),'6030.btw')
        
        self.logger.info("启动小标签打印机，同时打开模板文件")
        btformt = btEngine.Documents.Open(wbs_file_path,'Gprinter  GP-1125T*')
        
        self.logger.info("小标签打印开始...")
        for wbs_item in wbs:
            btformt.SubStrings.SetSubString('project_id',wbs_item[0:9])
            btformt.SubStrings.SetSubString('unit_id', wbs_item[11:])
            result=btformt.Print()
            
            if result:
                self.logger.info(wbs_item+":小标签打印成功")
            else:
                self.logger.error(wbs_item+":小标签打印失败")
       
        self.logger.info("小标签打印结束")
    
    
    def print_info(self,wbs): 
        data_pool=self.get_info_table(wbs) 
        
        if not data_pool:
            pass
         
        btEngine = printer.Engine()
    
        if not btEngine.IsAlive:
            btEngine.Start()
    
        wbs_file_path = os.path.join(os.getcwd(),'180100.btw')
        
        self.logger.info("启动信息表标签打印机，同时打开模板文件")
        btformt = btEngine.Documents.Open(wbs_file_path,'SLO-P3106(U)*')
        
        force_print = self.force_print_value.get()
        
        self.logger.info("信息表标签打印开始...")
        
        value=''
        flag=False
        str_info=''
        
        for data_line in data_pool:
            wbs_info = data_line['WBSNo']
            str_info=wbs_info+':'
            for dic1 in dic_name1:
                value = data_line[dic1]
                if not value:
                    break;
                if len(value) >34 and dic1=='ProjectName':
                    p_value = value[:34]
                else:
                    p_value=value
                    
                btformt.SubStrings.SetSubString(dic1,p_value)
            
            if not value:
                self.logger.error(wbs_info+'基本参数未上传(CE),中止打印')
                continue
            
            flag = data_line['gonvernorflag']
            for dic2 in dic_name2:
                if not flag and (dic2=='GovernorProductionNoCar' or dic2=='GovernorProductionNoCw'):
                    value=none_str
                elif not data_line[dic2] and force_print: 
                    value=force_str
                    str_info=str_info+head_name[dic2]+','
                elif not data_line[dic2] and not force_print:
                    str_info=wbs_info+':'
                    value=''
                    break
                else:
                    value=data_line[dic2]
                btformt.SubStrings.SetSubString(dic2,value)
            
            if str_info != wbs_info+':':
                self.logger.warn(str_info+'制造编号信息不全')
            
            if len(value)==0:
                self.logger.warn(str_info+'制造编号都未上传（质量), 中止打印')
                continue

            result=btformt.Print()
            
            if result:
                self.logger.info(data_line['WBSNo']+":信息表标签打印成功")
            else:
                self.logger.error(data_line['WBSNo']+":信息表标签打印失败")
       
        self.logger.info("信息表标签打印结束")            
                                                               
    '''  
    is no use  
    def clear_list(self):
        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row) 
    '''   
   
      
    def start_print(self):
        sel_lines = self.wbs_list.selection()
        
        if not sel_lines:
            self.logger.warn("请在列表中选择需打印的Unit!")
            pass
        
        sel_units=[]
        
        for line in sel_lines:
            sel_units.append(self.wbs_list.item(line,'values')[2])
            
        if self.wbs_check_value.get():
            self.print_wbs(sel_units)
        
        if self.infotable_check_value.get():
            self.print_info(sel_units)         
                      
    def stop_print(self):
        pass                      
               
    def display_list(self, event):
        for row in self.wbs_list.get_children():
            self.wbs_list.delete(row)
        cur_unit=self.conn.cursor()
        search_var = self.search_string.get()
        if not search_var.strip():
            pass            
        cur_unit.execute("select contract_id,wbs_no,project_name,lift_no,elevator_type from v_unit_info where contract_id like '%"+search_var+"%' or wbs_no like '%"+search_var+"%' ORDER BY wbs_no ASC;")
        row_unit = cur_unit.fetchall()
        
        i=-1;
        for row in row_unit:
            i=i+1;
            self.wbs_list.insert('',i,values=(str(i+1),row[0],row[1],row[2],row[3],row[4]))  
              
                               
    def createWidgets(self):
        self.list_label=Label(self)
        self.list_label["text"]="WBS NO(Unit清单)"
        self.list_label.grid(row=0, column=0, sticky=NSEW)
        
        self.wbs_list=ttk.Treeview(self,show='headings',columns=('col0','col1','col2','col3','col4','col5'))
        self.wbs_list.heading('col1', text='合同号')
        self.wbs_list.heading('col2', text='WBS No')
        self.wbs_list.heading('col3', text='项目名称')
        self.wbs_list.heading('col4', text='梯号')
        self.wbs_list.heading('col5', text='梯型')
        self.wbs_list.column('col0', width=20,anchor='sw')
        self.wbs_list.column('col1', width=50, anchor='sw')
        self.wbs_list.column('col2', width=100, anchor='sw')
        self.wbs_list.column('col3', width=200, anchor='sw')
        self.wbs_list.column('col4', width=50, anchor='sw')
        self.wbs_list.column('col5', width=60, anchor='sw')
        self.wbs_list.grid(row=1, column=0, rowspan=6,columnspan=4,sticky=NSEW)
        
        self.y_wbs_scroll=ttk.Scrollbar(self,orient=VERTICAL,command=self.wbs_list.yview)
        self.wbs_list.configure(yscrollcommand=self.y_wbs_scroll.set)
        self.y_wbs_scroll.grid(row=1,column=4,rowspan=6,sticky=NS)
        self.x_wbs_scroll=ttk.Scrollbar(self,orient=HORIZONTAL,command=self.wbs_list.xview)
        self.wbs_list.configure(xscrollcommand=self.x_wbs_scroll.set)
        self.x_wbs_scroll.grid(row=7,column=0, columnspan=4,sticky=EW)
        
        self.log_label=Label(self)
        self.log_label["text"]="operate log"
        self.log_label.grid(row=8, column=0, sticky=NSEW)
                     
        self.log_text=scrolledtext.ScrolledText(self, state='disabled')
        #self.log_text.configure(font='TkFixedFont')
        self.log_text.grid(row=9, column=0, rowspan=4, columnspan=6,sticky=NSEW)
        # Create textLogger
        text_handler = TextHandler(self.log_text)        
        # Add the handler to logger
        self.logger = logging.getLogger()
        self.logger.addHandler(text_handler)
        
        self.search_string=StringVar()
        self.input_entry=Entry(self, textvariable=self.search_string)
        #self.input_entry["width"]=50
        self.input_entry.bind('<Return>', self.display_list)
        self.input_entry.grid(row=1, column=5,padx=5,pady=10, sticky=NSEW)   
                    
        self.print_button=Button(self, width=20)
        self.print_button["text"]="打印"
        self.print_button.grid(row=2, column=5, sticky=NSEW)
        self.print_button["command"]=self.start_print
        
        self.wbs_check_value=IntVar()
        self.wbs_check=Checkbutton(self,width=20, variable=self.wbs_check_value)
        self.wbs_check["text"]="WBS NO"
        self.wbs_check.grid(row=4, column=5, sticky=NSEW)
        self.wbs_check_value.set(1)
        
        self.infotable_check_value=IntVar()
        self.infotable_check=Checkbutton(self, width=20, variable=self.infotable_check_value)
        self.infotable_check["text"]="项目信息表"
        self.infotable_check.grid(row=5,column=5, sticky=NSEW)
        self.infotable_check_value.set(1)
        
        self.force_print_value=IntVar()
        self.force_print_check=Checkbutton(self,width=20,variable=self.force_print_value)
        self.force_print_check["text"]="项目信息表强制打印"
        self.force_print_check.grid(row=6,column=5,sticky=NSEW)
        self.force_print_value.set(0)
                
        self.quit_button=Button(self, width=20)
        self.quit_button["text"]="退出"
        self.quit_button.grid(row=0, column=5, sticky=NSEW)
        self.quit_button["command"]=self.quit_func
        
                      
if __name__ == '__main__':
    root=Tk()
    Application(root)
    root.title("资料室批量打印程序")
    root.option_add('*Font', 'Verdana 10 bold')
    root.option_add('*EntryField.Entry.Font', 'Courier 10')
    root.geometry('700x580')  #设置了主窗口的初始大小960x540  
    root.mainloop()
    root.destroy()