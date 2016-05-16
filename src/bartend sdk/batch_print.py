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
import threading
from threading import Thread as thread
import time
from win32com import client

bt=client.gencache.EnsureModule("{D58562C1-E51B-11CF-8941-00A024A9083F}",0x0,10,0)
#bt=client.Dispatch("BarTender.Application")

class TextHandler(logging.Handler):
    """This class allows you to log to a Tkinter Text or ScrolledText widget"""
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text.configure(state='normal')
            self.text.insert(END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)# Scroll to the bottom
        

class print_wbs_thread(thread):
    def __init__(self, threadID, wbs):
        thread.__init__(self)
        self.threadID = threadID
        self.wbs = wbs
        
    def run(self):
        threadLock.acquire()
        print_wbs(self.wbs)
        time.sleep(1)
        print_info(self.wbs)
        time.sleep(3)
        threadLock.release()
       
def print_wbs(wbs):    
    wbs_file_path = os.path.join(os.getcwd(),'6030.btw')
    btFormt=bt.Open(wbs_file_path)
    bt.Visible=0
    
    for wbs_item in wbs:
       # btformt.SubStrings("project_id").Value = wbs_item[0:9]
        #btformt.SubStrings("unit_id").Value = wbs_item[11:]
        btformt.SubStrings.SetSubString('project_id',wbs_item[0:9])
        btformt.Print()
    
def print_info(wbs):
    pass

threadLock=threading.Lock()
threads=[]

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
            print_wbs(sel_units)
        
        if self.infotable_check_value.get():
            print_info(sel_units) 
                            
        pass
          
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
        self.wbs_list.grid(row=1, column=0, rowspan=5,columnspan=4,sticky=NSEW)
        
        self.y_wbs_scroll=ttk.Scrollbar(self,orient=VERTICAL,command=self.wbs_list.yview)
        self.wbs_list.configure(yscrollcommand=self.y_wbs_scroll.set)
        self.y_wbs_scroll.grid(row=1,column=4,rowspan=5,sticky=NS)
        self.x_wbs_scroll=ttk.Scrollbar(self,orient=HORIZONTAL,command=self.wbs_list.xview)
        self.wbs_list.configure(xscrollcommand=self.x_wbs_scroll.set)
        self.x_wbs_scroll.grid(row=6,column=0, columnspan=4,sticky=EW)
        
        self.log_label=Label(self)
        self.log_label["text"]="operate log"
        self.log_label.grid(row=7, column=0, sticky=NSEW)
                     
        self.log_text=scrolledtext.ScrolledText(self, state='disabled')
        #self.log_text.configure(font='TkFixedFont')
        self.log_text.grid(row=8, column=0, rowspan=4, columnspan=6,sticky=NSEW)
        
        # Create textLogger
        self.text_handler = TextHandler(self.log_text)
        # Add the handler to logger
        self.logger = logging.getLogger()
        self.logger.addHandler(self.text_handler)
        
        self.search_string=StringVar()
        self.input_entry=Entry(self, textvariable=self.search_string)
        #self.input_entry["width"]=50
        self.input_entry.bind('<Return>', self.display_list)
        self.input_entry.grid(row=1, column=5,padx=5,pady=10, sticky=NSEW)   
    
        '''
        no use
        self.clear_button=Button(self, width=20)
        self.clear_button["text"]="清空列表"
        self.clear_button.grid(row=2, column=5, sticky=NSEW)
        self.clear_button["command"]=self.clear_list
        
        
        self.stop_button=Button(self, width=20)
        self.stop_button["text"]="中止打印"
        self.stop_button.grid(row=3,column=5,sticky=NSEW)
        self.stop_button["command"]=self.stop_print
        '''
                
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