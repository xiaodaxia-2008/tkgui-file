# -*- coding: utf-8 -*-
"""
Created on Sun Apr  1 00:12:15 2018

@author: x00428488
"""
import time, logging, os, pandas as pd, datetime, re, xlsxwriter
import tkinter as tk

#######################定义处理ML文件的函数#################
        ##################################
def compare_MLfiles(get_text_contents, progressBar, status_info):
    try:
        starttime = time.time()
        progressBar['value']=1
        logging.debug('start check')
        
        fileA_path, fileB_path, result_save_path = get_text_contents()
        fileA_path = fileA_path.strip().splitlines()
        fileB_path = fileB_path.strip().splitlines()
        result_save_path = result_save_path.replace('\n','')
        progressBar['value']=5
        logging.debug(fileA_path)
        logging.debug(fileB_path)
        logging.debug(result_save_path)
        if not os.path.exists(result_save_path):
            os.makedirs(result_save_path)
            
        sitename_Regex = re.compile(r'\w+[-_]\d+') ####匹配站点，可以设置为匹配DU ID
        totalsteps = len(fileA_path)
        for step_i, fileAname in enumerate(fileA_path, start=1):
            logging.debug(fileAname)
            progressBar['value']= 5+round(step_i/totalsteps)*95
            sitename = sitename_Regex.search(fileAname)
            if sitename:
                sitename = sitename.group()
                status_info.insert('2.0', '\nthe '+str(step_i)+'th file ' + sitename+' is being checking...')
                logging.debug(sitename)
            for fileBname in fileB_path:
                if sitename in fileBname:
                    
                    logging.debug('start compare')
                    result, detailed_result = compare_excel_file(fileAname, fileBname)
                    xlsxname = result_save_path + '/'+ sitename +'__'+ datetime.datetime.now().strftime(
                            '%Y-%m-%d_%H-%M-%S') +'_ML_check.xlsx'
                    logging.debug('start write data')
                    logging.debug(xlsxname)
                    logging.debug(str(detailed_result)+'\n')
                    detailed_result.to_excel(xlsxname)
                    
                    status_info.insert('2.0', '\n'+str(detailed_result))
                    status_info.insert('2.0', '\n\n'+sitename + ' : '+str(result))     
        endtime = time.time()
        mesg = 'Cost '+str(endtime-starttime)+' seconds!'+'  Finished '+str(totalsteps)+' files'                   
        status_info.insert('2.0', '\n'+mesg)
        tk.messagebox.showinfo(title='Finished', message=mesg)
                
    except Exception as e:
        logging.debug(e)
        status_info.insert('2.0', '\n'+str(e))
        tk.messagebox.showerror(title='Error', message=e)        

def compare_excel_file(file_A, file_B):
    A_df = pd.read_excel(file_A)
    B_df = pd.read_excel(file_B)
    A_df.fillna(value=0, inplace=True)
    B_df.fillna(value=0, inplace=True)
    if A_df.shape == B_df.shape:
        A_df.sort_values(['Item Description','Serial Number'], inplace = True)
        B_df.sort_values(['Item Description','Serial Number'], inplace = True)
        B_df.index = A_df.index
        if (A_df == B_df).all(axis=1).all():
            return True, pd.DataFrame(['same value'])
        else:
            C_df = A_df[(A_df != B_df).any(axis=1)]
            return False, C_df
    else:
        return False, pd.DataFrame(['different size'])
        
#######################################     
        
#################### 生成material list的函数 ###################

   
def ML_generate(get_text_contents, progressBar, status_info):
    try:
        progressBar['value'] = 1
        
        CCM_shipped_file, SN_file, item_detail_file, save_file_path = get_text_contents()
    
        CCM_shipped_file = CCM_shipped_file.replace('\n','')
        SN_file = SN_file.replace('\n','')
        item_detail_file = item_detail_file.replace('\n','')
        save_file_path = save_file_path.replace('\n','')
        logging.debug(CCM_shipped_file + '\n' + SN_file + '\n' + save_file_path)
        status_info.insert('2.0', '\nCCM shipped file: '+CCM_shipped_file+'\nSN file: '+SN_file)

        if not os.path.exists(save_file_path):
            os.makedirs(save_file_path)
            
        starttime = time.time()
        ##处理Item Details文件，去掉重复，以Item Code 为索引
        logging.debug('读取Item Detials 文件。。。')
        status_info.insert('2.0', '\nloading data...')
    #    item_detail_file = 'ItemDetails.xlsx'  ##设置Item Details映射信息的文件路径
        Item_Details = pd.read_excel(item_detail_file, 0, header=0)
        Item_Details.drop_duplicates('Item Code',inplace=True)
        Item_Details.set_index('Item Code',inplace=True)
        
        ##读取CCM中的站点发货配置数据
        logging.debug('读取CCM 发货数据。。。')
        progressBar['value'] = 2
    #    CCM_shipped_file = '5640048_DCConfiguration__29546648_20180320232150941.xlsx' ##设置CCM数据文件的路径
        shipped_items = pd.read_excel(CCM_shipped_file, 1, header=0)
        shipped_items.drop(['Contract Register Date', 
                           'DU Name', 'Customer Site Name', 'Product',
                           'UOM', 'Bpart Line ID',
                           'Substitute ID', 'Survey Aux.', 'Current Qty.', 'Unshipped DN Qty.',
                           'Delivery Config Engineer', 'Spart Number',
                           'Split Flag', 'Delivery Mode'], axis=1, inplace=True)  ##去掉无用的列
        shipped_items.dropna(axis=0,subset=['Part Number*'], inplace=True)  ##去掉有nan值的行，CCM里面的数据有的Item Code是空的，去掉这样的行
        logging.debug('CCM 发货数据清理中。。。')
        status_info.insert('2.0', '\nclearing CCM data...')
        progressBar['value'] = 5
        shipped_items = DUconfigsum(shipped_items)  ##去掉Item Code重复的行，对Item Code的总数汇总，利用了DUconfigsum()函数
        
        ##读取站点存货的Serial Number 表格
        logging.debug('读取SN 文件。。。')
    #    lable_progress.config(text='loading SN file...')
    #    SN_file = 'baseSNInventoryExport_29546748_20180320232328241.xlsx'  ##设置SN文件的路径
        SN = pd.read_excel(SN_file, 0, header=0)
        SN.drop(['Project Code', 'Contract No.', 'Original Bill No.',
               'C/L No.', 'Box Name', 'Box No.', 'Box Status', 'Location Type',
               'Transaction No.', 'Transaction Date', 'Entity'], axis=1, inplace=True)  ##去掉无用的列
        
        
        ##分别一次选取CCM数据中的一个DU的发货数据进行处理
        logging.debug('开始生成ML文件。。。')
    #    lable_progress.config(text='starting generating ML file...')
        DUs = shipped_items['DU ID'].drop_duplicates()
        DU_total = len(DUs)
        for DU_num, DU in enumerate(DUs, start=1):
            mesg = 'Total ' + str(DU_total) + ' DUs, the ' + str(DU_num) + 'th DU '+str(DU) + ' is being processed...'
            logging.debug(mesg)
            status_info.insert('2.0', '\n'+mesg)
            progressBar['value'] = 5+round(DU_num/DU_total*95)
            DU_shipped_items = shipped_items[shipped_items['DU ID']==DU]  ##选择单个DU的发货数据
            row_num = DU_shipped_items.shape[0]  ##这个DU发货有多少个item
            sitename = DU_shipped_items['Customer Site ID'].iloc[0]
#            sitename = get_sitename_from_DUID(DU)  ##获取DU对应的站点名，SN 码表中有有时有站点名，没有DU ID
            item_SN = SN[(SN['Location Code']==sitename) | (SN['Location Code']==DU)]  ##选择这个DU ID对应的SN码数据
            
            ##利用CCM的站点发货配置数据初始化Material List 文件
            MLfile = pd.DataFrame({'Site ID':DU_shipped_items['DU ID'], 'USID':[str(sitename)+'_PH1_PSTN']*row_num, 
                                  'Related PO':DU_shipped_items['Customer PO NO.'].astype(str), 
                                   'BOQ scope':['PSTN']*row_num,
                                  'Configuration/Description':DU_shipped_items['Product Instance'],
                                   'Item Description':DU_shipped_items['Part Description'],
                                  'Qty per 1 Site':DU_shipped_items['Qty Sum'],
                                  'Serial Number':[None]*row_num,
                                  'Item Code':DU_shipped_items['Part Number*']})
            
            
            ##计算每个Item Code是哪种Item Details：Main Equipment，Site Materials， Liscense/SW
            Item_details_list = [None]*row_num
            j=0
            for i in MLfile.index.values:
                try:
                    Item_details_list[j]=Item_Details.loc[MLfile.loc[i,'Item Code'],'Item Details']
                except KeyError:
        #            logging.debug('No such item code in Item Details table')
                    pass
                j += 1
            MLfile['Item Details'] = Item_details_list  ##将计算出来的Item Details信息添加到MLfile
                
            ##读取每个Item Code对应的SN码，如果有的话，存到Serial_number字典里
            Serial_number = {}
            for itemcode in MLfile['Item Code']:
                SNcollect = []
                for Ser_Num in item_SN['SN']:
                    if itemcode[-5:] in Ser_Num:
                        SNcollect.append(Ser_Num)
                if SNcollect:
                    Serial_number[str(itemcode)] = SNcollect
            
            ##根据Serial Number的数量，将ML表格中的相应行进行拆分
            for itemcode in MLfile['Item Code']:
                if itemcode in Serial_number.keys():
                    group_item_SN = pd.DataFrame([], columns=MLfile.columns)
                    for value in Serial_number[itemcode]:
                        row_item = MLfile[MLfile['Item Code']==itemcode].copy()
                        row_item['Qty per 1 Site']=1
                        row_item['Serial Number'] = value
                        group_item_SN = group_item_SN.append(row_item)
                    if all(MLfile.loc[MLfile['Item Code']==itemcode]['Qty per 1 Site']==len(Serial_number[itemcode])):
                        MLfile.drop(MLfile[MLfile['Item Code']==itemcode].index, inplace=True)
                        MLfile = MLfile.append(group_item_SN, ignore_index=True)
                    
            ##对ML的列重新排序，去除ML表格中不需要的Item Code列
            MLheaders = ['Site ID', 'USID', 'Item Details', 'Related PO', 'BOQ scope', 'Configuration/Description', 
                         'Item Code', 'Item Description', 'Qty per 1 Site', 'Serial Number']
            MLfile = MLfile.loc[:,MLheaders]
            xlsxname = save_file_path + '/'+str(DU) + '__' + datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S') +'_ML.xlsx'
            write_df_to_excel(MLfile, xlsxname, filetype='Material List')
            status_info.insert('2.0', '\nfile saved to: '+ xlsxname)
                        
        endtime = time.time()
        nfiles = str(len(DUs))
        costtime = str(endtime-starttime)
        result_mesg = 'Generated '+nfiles+' MLfiles, cost '+costtime+' seconds!'
        logging.debug(result_mesg)
        status_info.insert('2.0', '\n'+ str(result_mesg))
        choice = tk.messagebox.askyesno(title='Finished', 
                                        message=result_mesg+'\nDo you want to open it?')
        if choice:
            os.startfile(save_file_path)
        
    except Exception as e:
        logging.debug(e)
        tk.messagebox.showerror(title='Error', message=e)
        status_info.insert('2.0', '\n'+ str(e))


###############################################################        
######写ML Dateframe到excel的函数    ############################
#写excel
def write_df_to_excel(data, filepath, filetype=None):
    
    workbook = xlsxwriter.Workbook(filepath, 
                                   {'nan_inf_to_errors': False,
                                    'strings_to_numbers': False})
    headerformat_1 = workbook.add_format({'font_name': 'Arial', 'font_size': 11,
                                              'bold': 1, 'text_wrap': 1, 
                                              'bg_color': 'yellow',
                                          'align': 'left', 'valign': 'vcenter',
                                          'border': 1})
    normalformat = workbook.add_format({'font_name':'Arial', 'font_size':11,
                                          'align':'left', 'valign':'vcenter',
                                          'border':1})
    sheet1 = workbook.add_worksheet(name='Sheet1')
    for colnum, column in enumerate(data.columns):
        sheet1.write(0, colnum, column, headerformat_1)
        for rownum, value in enumerate(data[column], start=1):
            sheet1.write(rownum, colnum, value, normalformat)
    if filetype == 'Material List':
        sheet1.set_column('A:A', 10)
        sheet1.set_column('B:B', 20)
        sheet1.set_column('C:C', 15)
        sheet1.set_column('D:E', 12)
        sheet1.set_column('F:F', 30)
        sheet1.set_column('G:G', 11)
        sheet1.set_column('H:H', 50)
        sheet1.set_column('I:I', 10)
        sheet1.set_column('J:J', 27)
        sheet1.freeze_panes(1, 1)
    elif filetype == 'PO File':
        POfile_brief = data[data['PO Item Code']!=''][
                            ['DU ID', 'PO Item Code', 'PR Line Description', 
                             'PR Line Quantity', 'UOM', 'Supplier Name']]
        sheet2 = workbook.add_worksheet(name='Sheet2')
        for colnum, column in enumerate(POfile_brief.columns):
            sheet2.write(0, colnum, column, headerformat_1)
            for rownum, value in enumerate(POfile_brief[column], start=1):
                sheet2.write(rownum, colnum, value, normalformat)
        sheet2.set_column('A:B', 25)
        sheet2.set_column('F:F', 40)
        sheet2.set_column('C:C', 80)
        sheet2.set_column('D:E', 10)
        sheet2.freeze_panes(1, 1)
        
        sheet1.set_column('A:C', 17)
        sheet1.set_column('D:D', 27)
        sheet1.set_column('E:E', 15)
        sheet1.set_column('F:F', 30)
        sheet1.set_column('G:G', 12)
        sheet1.set_column('H:H', 20)
        sheet1.set_column('I:I', 30)
        sheet1.set_column('J:J', 12)
        sheet1.set_column('L:L', 25)
        sheet1.freeze_panes(1, 1)
    workbook.close()


###############################################################            

###############################################################        
#################### CCM 数据清理函数#####################
def DUconfigsum(siteconfig):
    DUconfig = {}
    DUs = siteconfig['DU ID'].drop_duplicates()
    A = pd.DataFrame([])
    for DU in DUs:
        Partnums = siteconfig[siteconfig['DU ID']== DU]['Part Number*'].drop_duplicates()
        Partdict = {}
        for Partnum in Partnums: 
            partnumsum = siteconfig[(siteconfig['DU ID']==DU) & (siteconfig['Part Number*']==Partnum)]['Actual Qty.*'].sum()            
            Partdict[Partnum] = partnumsum
            
        DUconfig[DU] = Partdict
        A = A.append(pd.DataFrame([Partdict.keys(), Partdict.values(), [DU]*len(Partdict)]).T)
    siteconfig.drop_duplicates(['Part Number*', 'DU ID'], inplace=True)
    A.columns = ['Part Number*', 'Qty Sum', 'DU ID']
    C = siteconfig.merge(A, on=['Part Number*', 'DU ID'], sort=False, copy=False)
    return C        

#################从DU ID获取站点名的函数#########################
def get_sitename_from_DUID(DU):
    '''
    根据DU ID 提取站点名，比如输入'YE-21_TE'或者'YE-21_OSP'或者'YE-21_CW'，输出的就是'YE-21'
    DU ID的其他形式：AHL-1_PH2.1 Rep_TE,Myeik-1_PH2.1_TE等
    站点名是 字母-数字 或者 字母_数字
    利用正则表达式寻找输入DU的站点名
    '''
    ###利用去掉后缀来匹配
#    suffixRegex = re.compile(r'_OSP$|_TE$|_CW$|_PH2.1 Rep_TE$|_PH2.1_TE$')
#    sitename = suffixRegex.sub('', DU)
    
    ####利用前缀格式来匹配，似乎更简单一点
    prefixRegex = re.compile(r'^\w+[-_]\d+')
    sitename = prefixRegex.search(DU)
    if sitename:
        return sitename.group()
    else:
        return 'XXXNO DU IDXXX'
