# -*- coding: utf-8 -*-
"""
Created on Sat Mar 24 14:31:38 2018

@author: x00428488

updated on 25-Mar-2018
make this GUI independent from other person files
all funcitons and Class is included in this single file

"""

import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
import tkinter.colorchooser
import sys, os
sys.path.append('./lib')
sys.path.append('./templateSourceData')
import _thread as thread  ##### GUI的button 事件处理时间较长时，用线程
import logging
from general_functions_for_transfer_report_v4 import transfer_report_data
from general_functions_for_ML import ML_generate, compare_MLfiles
from ISDP_JFSheet1 import create_JF_Sheet1
from POfile_create import PO_generate
from POfile_compare import compare_PO_file



#####################自定义Frame类， 用于生成不同功能的界面######################
class func_frame(tk.Frame):
#    padx=5, pady=5, borderwidth=2, relief='groove'
    buttontexts = ['Select Manual MLfile', 'Select Auto MLfile', 'Result save path']
    
#    progress_bar = self.
#    global status_info
    info_text = '***** Processing Information Will Be Shown Here *****\n'
    def __init__(self, row=0, column=0, parent=None, func_name='Undefined', buttontexts=buttontexts):
        self.style1 = ttk.Style()
        self.style1.configure("TButton", padding=6, relief='raised', font=('times',12,'bold'), 
                  background='darkgrey', foreground='midnightblue')
        self.style1.configure("C.TButton", font=('times',20,'italic bold'),
                   foreground='green',
                   background='yellow'
                   )
        
        tk.Frame.__init__(self, parent, padx=5, pady=5, borderwidth=1, relief='groove')
        ttk.Sizegrip(self).grid(column=999, row=999, sticky=('se'))
        self.func_name= func_name
        self.buttontexts = buttontexts
        self.path_texts = [None]*len(self.buttontexts)
        self.btns = [None]*len(self.buttontexts)
        self.grid(row=row, column=column, sticky=('news'))
        
        self.make_func_lable()
        self.make_filepath()

        self.status_info = self.make_status_info()
        self.progress_bar = self.make_progress_bar()
        self.make_start_btn()
        self.scroll_bar = self.make_scroll_bar()
        self.status_info['yscrollcommand'] = self.scroll_bar.set
        
        self.rowconfigure(8, weight=1)
        self.columnconfigure(0, weight=4)
        self.columnconfigure(1, weight=1)
        self.show = True
    
    def make_func_lable(self):
        func_name_label = tk.Label(self, width=40, height=2, relief='flat', 
                                   text=self.func_name, font=('Helvetica', 20, 'bold italic'), 
                                   bg='honeydew', fg='olive')
        func_name_label.grid(row=0, column=0, columnspan=2, sticky=('we'))
        
    def make_filepath(self):

        for i, buttontext in enumerate(self.buttontexts, start=0):
            self.path_texts[i] = tk.Text(self, bg='ghostwhite', fg='black', height=1,  
                           width=60, font=('times',10,'normal'))
            self.path_texts[i].grid(row=i+1, column=0, padx=5, pady=5, sticky=('wens'))
            if self.func_name == 'Check Materials List Files':
                self.btns[i] = ttk.Button(self,text=self.buttontexts[i], command=(lambda 
                         x=i:self.selectfiles(x)), width=15, style='TButton')
            else:
                self.btns[i] = ttk.Button(self,text=self.buttontexts[i], command=(lambda 
                     x=i:self.selectfile(x)), width=15, style='TButton')
            self.btns[i].grid(row=i+1, column=1, padx=5, pady=5, sticky=('we'))
        self.btns[-1].config(command=lambda:self.select_save_path(i))
        
            
#        path_texts[-1].insert('0.0', 'D:/MLfile_result') 
    
    def make_start_btn(self):
        generateML_btn = ttk.Button(self, text='Start', width=10, cursor='hand2', style="C.TButton")
        generateML_btn.grid(row=6, column=1, sticky=('ns'), padx=5, pady=5)
        cwd = os.getcwd().replace(u'\\', '/')
        if self.func_name == 'Generate Materials List Files':
#            generateML_btn.config(command=lambda:thread.start_new_thread(ML_generate, 
#            (self.get_text_contents, self.progress_bar, self.status_info)))
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, ML_generate, self.get_text_contents, 
                                                    self.progress_bar, self.status_info)) 
            self.status_info.insert('2.0', '\nkeep CCM shipped file at the second sheet')
            self.path_texts[-2].insert('1.0', cwd+'/templateSourceData/ItemDetails.xlsx')
            self.path_texts[-1].insert('1.0', cwd+'/Results/MLfiles')
        elif self.func_name == 'Check Materials List Files':
#            generateML_btn.config(command=lambda:thread.start_new_thread(compare_MLfiles, 
#            (self.get_text_contents, self.progress_bar, self.status_info)))
#            generateML_btn.config(command=lambda:thread.start_new_thread(testfunc, 
#            (self.status_info, '2.0', '\n'+str(self.get_text_contents()))))
            self.path_texts[-1].insert('1.0', cwd+'/Results/CheckMLfile')
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, compare_MLfiles, self.get_text_contents, 
                                                    self.progress_bar, self.status_info))   
        elif self.func_name == 'Generate Report':
            self.status_info.insert('2.0', 
                                    '\nOriginal file should be like "All implementation plan and progress  2018-04-01.xlsx"')
            self.path_texts[-1].insert('1.0', cwd+'/Results/CustomerReport')
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, transfer_report_data, self.get_text_contents, 
                                                    self.progress_bar, self.status_info))        
        elif self.func_name == 'Generate Report Sheet1':
            self.path_texts[-1].insert('1.0', cwd+'/Results/ReportSheet1')
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, create_JF_Sheet1, self.get_text_contents, 
                                                    self.progress_bar, self.status_info))   
        elif self.func_name == 'PO file create':
            self.status_info.insert('2.0', '\nOriginal file should be CCM site shipped materials')
            self.path_texts[-2].insert('1.0', cwd+'/templateSourceData/ItemDetails.xlsx')
            self.path_texts[-1].insert('1.0', cwd+'/Results/POfile')
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, PO_generate, self.get_text_contents, 
                                                    self.progress_bar, self.status_info))   
        elif self.func_name == 'PO file compare':
            self.path_texts[-1].insert('1.0', cwd+'/Results/ComparePOfile')
            generateML_btn.bind('<Button-1>', 
                                self.handlerAdaptor(self, compare_PO_file, self.get_text_contents, 
                                                    self.progress_bar, self.status_info))   
        else:
            pass
            
    def make_status_info(self):
#        global status_info 
        status_info = tk.Text(self, bg='ghostwhite', fg='black')
        status_info.grid(row=8, column=0, columnspan=2, padx=5, pady=5, sticky=('news'))
#        status_info.insert('insert', info_text)
        status_info.insert('end', self.info_text, ('highlightline', 'recent', 'warning'))
        status_info.tag_configure('highlightline', background='oldlace', 
                                  font='timesnewroman 12 bold', relief='groove')
        return status_info
        
    def make_scroll_bar(self):
        s = ttk.Scrollbar(self, orient='vertical', command=self.status_info.yview)
        s.grid(row=8, column=999, sticky=('nse'))
        return s
    
    def make_progress_bar(self):
        progress_bar = ttk.Progressbar(self, orient=tk.HORIZONTAL, mode='indeterminate')
        progress_bar.grid(row=10, column=0, sticky=('wes'))
        return progress_bar
        
    def selectfile(self, num):
        path = self.path_texts[num].get('1.0', 'end')
        filename = tkinter.filedialog.askopenfilename(filetypes=[('Excel file', ('.xls', '.xlsx'))])
        self.path_texts[num].delete('1.0',tk.END)
        if filename:
            self.path_texts[num].insert('insert', filename)
        else:
            self.path_texts[num].insert('insert', path)
    
    def selectfiles(self, num):
        path = self.path_texts[num].get('1.0', 'end')
        filenames = tkinter.filedialog.askopenfilenames(filetypes=[('Excel file', ('.xls', '.xlsx'))])
        self.path_texts[num].delete('1.0',tk.END)
        if filenames:
            for i, filename in enumerate(filenames, start=1):
                position = str(i)+'.0'
                self.path_texts[num].insert(position, filename+'\n')
        else:
            self.path_texts[num].insert('insert', path)

      
    def select_save_path(self, num):
        path = self.path_texts[num].get('1.0', 'end')
        path_name = tkinter.filedialog.askdirectory()
        self.path_texts[num].delete('1.0', tk.END)
        if path_name:
            self.path_texts[num].insert('insert', path_name)
        else:
            self.path_texts[num].insert('insert', path)
        
    def handlerAdaptor(event, self, fun, *args):  
        '''''事件处理函数的适配器，相当于中介，那个event是从那里来的呢，我也纳闷，这也许就是python的伟大之处吧'''  
        return lambda event, fun=fun, args=args: thread.start_new_thread(fun, args)
#        lambda args=(self.get_text_contents, self.progress_bar, self.status_info): 
#                           thread.start_new_thread(fun, args)

    def get_text_contents(self):
        file = []
        for i in range(len(self.buttontexts)):
            file.append(self.path_texts[i].get('0.0','end'))
        return file



############## 生成主窗口的类#########################
class main_win():
    def __init__(self, parent=None):
        self.root = tk.Tk()
        self.root.title('Raven Creator')
        self.root.geometry('800x600+400+100')
        self.root.option_add('*tearOff', False)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.framename = None
        self.frame = None
        self.make_menu_bar()
        self.root.mainloop()
            
    def make_menu_bar(self):
        ###窗口菜单创建    
        menufont= ('helitica', 10, 'normal')
        menubar = tk.Menu(self.root, font=menufont, bg='yellow', fg='black')
        self.root.config(menu=menubar)
        
        filemenu = tk.Menu(menubar, tearoff=0, font=menufont)
        filemenu.add_command(label="New", command=self.donothing)
        filemenu.add_command(label="Open", command=self.donothing)
        filemenu.add_command(label="Save", command=self.donothing)
        filemenu.add_command(label="Save as...", command=self.donothing)
        filemenu.add_command(label="Close", command=self.donothing)
        filemenu.add_separator() 
        filemenu.add_command(label="Exit", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=filemenu)
        
        editmenu = tk.Menu(menubar, tearoff=0, font=menufont)
        editmenu.add_command(label="Undo", command=self.donothing)
        editmenu.add_separator()
        editmenu.add_command(label="Cut", command=self.donothing)
        editmenu.add_command(label="Copy", command=self.donothing)
        editmenu.add_command(label="Paste", command=self.donothing)
        editmenu.add_command(label="Delete", command=self.donothing)
        editmenu.add_command(label="Select All", command=self.donothing)
        menubar.add_cascade(label="Edit", menu=editmenu)
        
        
        ### 核心功能菜单，实现功能的各个模块
        menu_function = tk.Menu(menubar, font=menufont)
        ## material list相关的功能菜单
        menu_ML = tk.Menu(menu_function)
        menu_ML.add_command(label='Genrate MLfile', command=lambda: self.generate_frame())
        menu_ML.add_separator()
        menu_ML.add_command(label='Check MLfile', command=lambda: self.check_frame())
        menu_ML.add_separator()
        menu_ML.add_command(label='Generate PO file', command=lambda: self.PO_create_frame())
        menu_ML.add_separator()
        menu_ML.add_command(label='Compare PO file', command=lambda: self.PO_compare_frame())
        menu_function.add_cascade(menu=menu_ML, label='ML/PO/AF')
        ## 报告转换的功能菜单
        menu_report = tk.Menu(menu_function)
        menu_report.add_command(label='Generate Report', command=lambda: self.report_frame())
        menu_report.add_separator()
        menu_report.add_command(label='Generate Report Sheet1', 
                                  command=lambda: self.report_Sheet1_frame())
        menu_function.add_cascade(menu=menu_report, label='Report Transfer')
        menubar.add_cascade(menu=menu_function, label='Function')
        
        
        
        menu_preference = tk.Menu(menubar, font=menufont)
        menu_pref_colorset = tk.Menu(menu_preference)
        menu_pref_colorset.add_command(label='Choose background color', 
                                    command=self.set_background_color)
        menu_preference.add_cascade(menu=menu_pref_colorset, label='Color Set')
        menu_preference.add_command(label='Font', 
                                    command=self.donothing)
        menubar.add_cascade(menu=menu_preference, label='Preference')
        
        helpmenu = tk.Menu(menubar, tearoff=0, font=menufont)
        helpmenu.add_command(label="Help Index", command=self.donothing)       
        helpmenu.add_command(label='Guide', command = lambda: 
            tk.messagebox.showinfo(title='Guide', message='Ha Ha, there is no guide!'))
        helpmenu.add_command(label="About...", command=lambda:
            tk.messagebox.showinfo('About the program', 'Made by Nobody, 24-Mar-2018'))
        menubar.add_cascade(label="Help", menu=helpmenu)
    
    def set_background_color(self):
        pass
        bgcolor = tkinter.colorchooser.askcolor()
        if self.frame:
            logging.debug(bgcolor)
            self.frame.config(bg=bgcolor[1])
            logging.debug(bgcolor)

    
    def generate_frame(self):   
        # 生成material list文件的界面
        if self.framename == 'frame_generate':
            pass
        else:
            try:
                self.frame.destroy()
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                generate_btntexts = ['Select CCM data', 'Select SN data', 'Select Item Details', 'Save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='Generate Materials List Files', 
                                        buttontexts=generate_btntexts)
                self.framename = 'frame_generate'
                logging.debug(self.framename)
#        return 'frame_generate'
    
    def check_frame(self):
        # 检查material list的界面
        if self.framename == 'frame_check':
            pass
        else:
            try:
#                self.framename == 'frame_generate':
                self.frame.destroy()  
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                check_btntexts = ['Select Manual MLfile', 'Select Auto MLfile', 'Result save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='Check Materials List Files', 
                                        buttontexts=check_btntexts)
                self.framename = 'frame_check'
                logging.debug(self.framename)
#        return 'frame_check'
    
    def report_frame(self):
        # 从Sheet1 生成客户模板报告的界面
        if self.framename == 'frame_report':
            pass
        else:
            try:
#                self.framename == 'frame_generate':
                self.frame.destroy()  
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                report_btntexts = ['Select original file', 'Result save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='Generate Report', 
                                        buttontexts=report_btntexts)
                self.framename = 'frame_report'
                logging.debug(self.framename)
#        return 'frame_check'
    
    def report_Sheet1_frame(self):
        # 从ISDP生成Sheet1的界面
        if self.framename == 'frame_report_Sheet1':
            pass
        else:
            try:
#                self.framename == 'frame_generate':
                self.frame.destroy()  
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                report_btntexts = ['Select ISDP file', 'Result save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='Generate Report Sheet1', 
                                        buttontexts=report_btntexts)
                self.framename = 'frame_report_Sheet1'
                logging.debug(self.framename)
#        return 'frame_check'
    
    def PO_create_frame(self):
        # 从ISDP生成Sheet1的界面
        if self.framename == 'frame_PO_file_create':
            pass
        else:
            try:
#                self.framename == 'frame_generate':
                self.frame.destroy()  
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                report_btntexts = ['Select CCM file', 'Select Item-PO mapping file', 'Result save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='PO file create', 
                                        buttontexts=report_btntexts)
                self.framename = 'frame_PO_file_create'
                logging.debug(self.framename)
#        return 'frame_check'
    
    def PO_compare_frame(self):
        # 从ISDP生成Sheet1的界面
        if self.framename == 'frame_PO_file_compare':
            pass
        else:
            try:
#                self.framename == 'frame_generate':
                self.frame.destroy()  
            except Exception as e:
                logging.debug('first time generate frame')
            finally:
                report_btntexts = ['Select Mannual PO file', 'Select Auto PO file', 'Result save path']
                self.frame = func_frame(parent=self.root, 
                                        func_name='PO file compare', 
                                        buttontexts=report_btntexts)
                self.framename = 'frame_PO_file_compare'
                logging.debug(self.framename)
#        return 'frame_check'
    
    def donothing(self):
        tk.messagebox.showinfo('Info', 'Undefined funciton')
   
if __name__ == '__main__':
    main_win()
