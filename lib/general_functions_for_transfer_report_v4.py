# -*- coding: utf-8 -*-
"""
Created on Mon Mar 19 22:53:31 2018

@author: x00428488
"""
import pandas as pd, numpy as np
import datetime, logging, os, time, xlsxwriter
import tkinter as tk
from map_milestone_v4 import map_milestone

def transfer_report_data(get_text_contents, progressBar, status_info):
    
    try:
        progressBar.start()
        
        origin_file, save_file_path = get_text_contents()
    
        origin_file = origin_file.replace('\n','')
        save_file_path = save_file_path.replace('\n','')
        logging.debug(origin_file)
        logging.debug(save_file_path)
        status_info.insert('2.0', '\norigin file: ' + origin_file)

        if not os.path.exists(save_file_path):
            os.makedirs(save_file_path)
            
        starttime = time.time()
        status_info.insert('2.0', '\nloading data...')
#        origin_file = 'D:/MLtool/MLPO/Results/ReportSheet1/Sheet1_2018-04-08_20-40-49.xlsx'
        origin_data = pd.read_excel(origin_file, sheet_name='Sheet1', 
                                    header=[0, 1])
        status_info.insert('2.0', '\nformating index...')
        set_mutiIndex(origin_data) # 格式化muti index
        status_info.insert('2.0', '\nformating finished...')
        # 设置行的索引为Site ID
        origin_data.set_index('Customer Site ID', drop=False, inplace=True)
        status_info.insert('2.0', '\nDrop duplicates...')      
#        origin_data.drop_duplicates(subset='Customer Site ID', keep='first', inplace=True)
        origin_data.drop_duplicates(subset=('Customer Site ID', ''), inplace=True)
        origin_data.dropna(axis=0,subset=[('Customer Site ID', '')], inplace=True)  ##去掉有nan值的行
        origin_data.fillna(' ', inplace=True)
        origin_data['MSAN Type'] = origin_data['MSAN Type'].astype(str)
        
        status_info.insert('2.0', '\nsetting index...')             
        exchange_set = origin_data['Exchange'].drop_duplicates()
        status_info.insert('2.0', '\ncreating exchange_set...')
        exchanges_data = {'DUs':{}}
        for exchange in exchange_set:
            exchanges_data['DUs'][exchange] = origin_data[origin_data['Exchange']==exchange]['Customer Site ID'].values
            status_info.insert('2.0', '\n'+str(exchange))
        
        DUs_data = {}
        #DU_ID = 'AHL-15'
        status_info.insert('2.0', '\nstarting computing...')
        for DU_ID in origin_data.index:
            DUdata = {}
            DUdata['Exchange'] = origin_data.loc[DU_ID, 'Exchange']
            DUdata['Site Type'] = origin_data.loc[DU_ID, 'Site Type']
            if origin_data.loc[DU_ID, 'MSAN Type'][0].strip() == 'S200':
                DUdata['MSAN Type'] = 'Pole'
#            DUdata['Rollout Date'] = compute_report_DU_data(DU_ID, origin_data, columnnames)
            DUdata['Rollout Date'] = compute_report_DU_data_mapmilestone(DU_ID, 
                                                  origin_data, map_milestone)
            DUs_data[DU_ID] = DUdata
            status_info.insert('2.0', '\ncomputing '+ str(DU_ID))
        
        # 新建excel文件
#        save_file_path = 'D:/MLtool/ISDPreport/Publish_verson/version_3'
        timenow = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        excel_name = save_file_path + '/Report_for_customer_' + timenow +'.xlsx'
        workbook = xlsxwriter.Workbook(excel_name, {'nan_inf_to_errors': True})
        
        
        
#        sheet_num = 1
        worksheet = workbook.add_worksheet(name='Site Details')
#        dateformat = workbook.add_format({'font_name':'Arial', 'font_size':9,
#                                          'num_format': 'dd/mm/yy', 
#                                          'align':'center', 'valign':'vcenter',
#                                          'border':1})
        normalformat = workbook.add_format({'font_name':'Arial', 'font_size':9,
                                          'align':'center', 'valign':'vcenter',
                                          'border':3})
        dateformat =  workbook.add_format({'font_name':'Arial', 'font_size':9,
                                          'align':'center', 'valign':'vcenter',
                                          'border':3})
        dateformat.set_num_format('d/m/yyyy')
        headerformat_1 = workbook.add_format({'font_name':'Arial', 'font_size':11,
                                              'bold': 1, 'text_wrap': 1, 
                                              'align':'left', 'valign':'vcenter',
                                              'border':3})
        headerformat_2 = workbook.add_format({'font_name':'Arial', 'font_size':11,
                                              'bold':1, 'text_wrap': 1, 
                                              'align':'left', 'valign':'vcenter',
                                              'border':3})
        headerformat_2.set_bold(False)
        
        write_sheet_row_labels(worksheet, headerformat_1, headerformat_2)
        status_info.insert('2.0', '\n'+ 'starting dumping files...')
        k = 0
        for DU_ID, DU_data in DUs_data.items():
            write_report_DU_data(worksheet, DU_ID, DU_data['Rollout Date'], k, headerformat_1, 
                                 dateformat, normalformat)
            logging.debug('Writing ' + DU_ID)
            status_info.insert('2.0', '\n'+'Writing ' + DU_ID)
            k += 6

        worksheet.freeze_panes(4, 2)
        
        
        # 下面开始写summary Sheet
        status_info.insert('2.0', '\n'+'starting computing summary sheet...')
        sheet_summary = workbook.add_worksheet(name='Summary')
        # 写行标签
        write_sheet_row_labels(sheet_summary, headerformat_1, headerformat_2, sumsheet=True)
        # 写每个区域的汇总
        exchange = 'PZD(PHASE1)'
        k = 0
        for exchange, DUs in exchanges_data['DUs'].items():
            status_info.insert('2.0', '\n'+'computing summary sheet: '+str(exchange))
#            exchange_data = np.random.rand(15, 43).astype(np.object)
            exchange_data = np.repeat(np.repeat(' ', 43), 15).reshape(15, 43).astype(np.object)
            exchange_data[0] = np.repeat(origin_data[(origin_data['Exchange']==exchange) & 
                                          (origin_data['Site Type']=='Indoor')].shape[0], 43) # 计算indoor站点数
            exchange_data[2] = np.repeat(origin_data[(origin_data['Exchange']==exchange) &
                                (origin_data['MSAN Type']=='S200')].shape[0], 43)  # 计算Pole的数量
            exchange_data[1] = np.repeat(origin_data[(origin_data['Exchange']==exchange) & 
                                          (origin_data['Site Type']=='Outdoor')].shape[0], 43) - \
                                            exchange_data[2]  # 计算outdoor 的数量
            exchange_data[6] = np.array([2, 1, 1, 4, 1, 2, 2, 7, 7, 3, 3, 2, 2, 3, 1, 1, 
                                 3, 3, 3, 5, 2, 2, 5, 1, 2, 1, 2, 2, 2, 1, 1, 1, 
                                 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]).astype(np.str_).astype(np.object)  # standard deviation
            exchange_data[14] = exchange_data[6] + '/' + exchange_data[6] # 计算duration
            DU = DUs[0]
            
            
            # 整合每个区域的计划、实际日期
            sites = np.array([])
            exchange_date = {'PlanStart':pd.DataFrame(),
                             'PlanEnd':pd.DataFrame(),
                             'ActualStart':pd.DataFrame(),
                             'ActualEnd':pd.DataFrame()}
            for DU in DUs:
                sites = np.r_[sites, np.array([DU])]
#                origin_data[origin_data['Exchange']==exchange]['TSSR approval'].max()
                planstart = pd.DataFrame(DUs_data[DU]['Rollout Date'][0][1:]).T
                planend = pd.DataFrame(DUs_data[DU]['Rollout Date'][1][1:]).T
                actualstart = pd.DataFrame(DUs_data[DU]['Rollout Date'][2][1:]).T
                actualend = pd.DataFrame(DUs_data[DU]['Rollout Date'][3][1:]).T
                
                exchange_date['PlanStart'] = exchange_date['PlanStart'].append(planstart, 
                             ignore_index=True)
                exchange_date['PlanEnd'] = exchange_date['PlanEnd'].append(planend, 
                             ignore_index=True)
                exchange_date['ActualStart'] = exchange_date['ActualStart'].append(actualstart, 
                             ignore_index=True)
                exchange_date['ActualEnd'] = exchange_date['ActualEnd'].append(actualend, 
                             ignore_index=True)
            
            
            # 计算每个区域的开始结束日期
            exchange_min_plan_start = []
            exchange_max_plan_end = []
            exchange_min_actual_start = []
            exchange_max_actual_end = []
            num_of_milestiones = exchange_date['PlanStart'].shape[1]
            for col in range(0, num_of_milestiones):
                exchange_min_plan_start.append(exchange_date['PlanStart'] \
                                    [not_blank_index(exchange_date['PlanStart'][col])][col].min())
                exchange_max_plan_end.append(exchange_date['PlanEnd'] \
                                    [not_blank_index(exchange_date['PlanEnd'][col])][col].max())
                exchange_min_actual_start.append(exchange_date['ActualStart'] \
                                    [not_blank_index(exchange_date['ActualStart'][col])][col].min())
                exchange_max_actual_end.append(exchange_date['ActualEnd'] \
                                    [not_blank_index(exchange_date['ActualEnd'][col])][col].max())
#                print(col)
            
            i = 0
            for datestart, dateend in zip(exchange_min_plan_start, exchange_max_plan_end):
                if isinstance(datestart, datetime.datetime) :
                    exchange_data[3][i] = str(datestart.day)   # 计算PlanStartDay/PlanEndDay
                    exchange_data[4][i] = str(datestart.month)   # 计算PlanStartMonth/PlanEndMonth
                    exchange_data[5][i] = str(datestart.year)  # 计算PlanStartYear/PlanEndYear
                exchange_data[3][i] += '/'
                exchange_data[4][i] += '/' 
                exchange_data[5][i] += '/'
                if isinstance(dateend, datetime.datetime):
                    exchange_data[3][i] += str(dateend.day)
                    exchange_data[4][i] += str(dateend.month) 
                    exchange_data[5][i] += str(dateend.year)
                i += 1
            
            i = 0
            for datestart, dateend in zip(exchange_min_actual_start, exchange_max_actual_end): 
                if isinstance(datestart, datetime.datetime) :
                    exchange_data[11][i] = str(datestart.day)   # 计算ActualStartDay/AcutalEndDay
                    exchange_data[12][i] = str(datestart.month)   # 计算ActualStartMonth/ActualEndMonth
                    exchange_data[13][i] = str(datestart.year)  # 计算ActualStartYear/ActualEndYear
                exchange_data[11][i] += '/'
                exchange_data[12][i] += '/'  
                exchange_data[13][i] += '/'
                if isinstance(dateend, datetime.datetime):
                    exchange_data[11][i] += str(dateend.day)
                    exchange_data[12][i] += str(dateend.month)  
                    exchange_data[13][i] += str(dateend.year)
                i += 1
            
            # 开始计算每个区域的finish/pending site count/list  exchange_data[7,8,9,10]
            total_sites = len(DUs)
            # 计算finish sites count
            exchange_data[7] = [exchange_date['ActualEnd'] \
                                    [not_blank_index(exchange_date['ActualEnd'][i])][i].shape[0] 
                                    for i in range(0, num_of_milestiones)]  
            exchange_data[9] = total_sites - exchange_data[7] # 计算pending sites count
                        
            exchange_data[8] = [sites[not_blank_index(exchange_date['ActualEnd'][col])]  
                                    for col in range(num_of_milestiones)]
            exchange_data[10] = [' '.join(np.setdiff1d(sites, exchange_data[8][col])) 
                                    for col in range(num_of_milestiones)]
            exchange_data[8] = [' '.join(exchange_data[8][col]) for col in range(num_of_milestiones)]
            
            # 往summary sheet 写入这个exchange 的数据
            write_summary_exchange_data(sheet_summary, exchange, exchange_data, 
                                        headerformat_1, normalformat, k)
            exchanges_data[exchange] = exchange_data
            k += 16
        
        sheet_summary.freeze_panes(5, 3)
        workbook.close()
        status_info.insert('2.0', '\nfile saved to: '+ excel_name)
        progressBar.stop()
        endtime = time.time()
        DU_total = str(len(DUs_data.keys()))
        costtime = str(endtime-starttime)
        result_mesg = 'Generated '+DU_total+' DUs report, cost '+costtime+' seconds!'
        logging.debug(result_mesg)
        status_info.insert('2.0', '\n'+ str(result_mesg))
        choice = tk.messagebox.askyesno(title='Finished', 
                                        message=result_mesg+'\nDo you want to open it?')
        if choice:
            os.startfile(excel_name)
    except Exception as e:
        logging.debug(e)
        progressBar.stop()
        tk.messagebox.showerror(title='Error', message=e)
        status_info.insert('2.0', '\n'+ str(e))



# days 可以是正或者负
def add_date(origin_date, days=0):
    if not isinstance(origin_date, (pd._libs.tslib.Timestamp, datetime.datetime)) :
        return ' '
    delta_days = datetime.timedelta(days)
    return origin_date + delta_days


def get_index_of_next_column(lst, columnname):
    return lst[lst.index(columnname)+1]


def date_str_format(date):
    return date.strftime('%d-%b-%Y')

def norm_read_date(date_lst):
    date_temp = []
    for date in date_lst:
        # 如果date=NaT
#         type(date) = pandas._libs.tslib.NaTType 但是isinstance(date, datetime.datetime) 竟然是True
        if not (type(date) == pd._libs.tslib.Timestamp or type(date) == datetime.datetime):
#        if not isinstance(date, (pd._libs.tslib.Timestamp, datetime.datetime)):
            date = ' '
        date_temp.append(date)
    return date_temp


def duration_cac(startdate, enddate):
    if isinstance(startdate, (pd._libs.tslib.Timestamp, datetime.datetime)) and \
    isinstance(enddate, (pd._libs.tslib.Timestamp, datetime.datetime)):
        return str((enddate - startdate).days + 1)
    else:       
        return ' '


def end_date_cac(start_date_lst, no):
    if isinstance(start_date_lst[no], (pd._libs.tslib.Timestamp, datetime.datetime)):
        for i in range(1, len(start_date_lst)-no):
            if isinstance(start_date_lst[no+i], (pd._libs.tslib.Timestamp, datetime.datetime)):
                return start_date_lst[no+i]
        return ' '
    else:
        return ' '


def start_date_cac(end_date_lst, no):
    if isinstance(end_date_lst[no], (pd._libs.tslib.Timestamp, datetime.datetime)):
        for i in range(1, no):
            if isinstance(end_date_lst[no-i], (pd._libs.tslib.Timestamp, datetime.datetime)):
                return end_date_lst[no-i]
        return ' '
    else:
        return ' '



# 计算一列中是日期的元素的索引
def not_blank_index(ser):
    index = []
    for i in range(0, ser.shape[0]):
        tf = True if ser[i] != ' ' else False
        index.append(tf)
    return index

# 格式化一下muti index， 去掉那些二层索引没意义的索引项
def set_mutiIndex(df, level2_values=None):
    '''
    df should be a pandas Dataframe, level2_values should be the second column labels 
    you want to keep.
    '''
    if not level2_values:
        level2_values = ['Plan Start Date', 'Actual Start Date', 
                         'Plan End Date', 'Actual End Date', 'Actual', 'Plan']
    df.index.name = df.columns.names[0]
    df.columns.names = [None, None]
    df.reset_index(inplace=True)
    levels = df.columns.levels
    level2 = levels[1].values
    level1 = [ele.strip() if type(ele)==str else ele for ele in levels[0].values]
    level2[[ False if level2[i] in level2_values else True 
                for i in range(len(level2))]] = ''
    newlevels = [level1, level2]
    df.columns.set_levels(newlevels, inplace=True)
    
def compute_report_DU_data_mapmilestone(DU_ID, origin_data, map_milestone):
    # 初始化，日期从位置1开始存储
    plan_start_date = [' ']*(len(map_milestone)+1)
    actual_start_date = [' ']*(len(map_milestone)+1)
    plan_end_date = [' ']*(len(map_milestone)+1)
    actual_end_date = [' ']*(len(map_milestone)+1)
#    # 开始日期
#    for i, hwmilestone in enumerate(map_milestone, start=1):
#        if type(hwmilestone[1]) == str:
#            plan_start_date[i] = origin_data.loc[DU_ID, (hwmilestone[1],'Plan End Date')]
#            actual_start_date[i] = origin_data.loc[DU_ID, (hwmilestone[1],'Actual End Date')]
#        elif type(hwmilestone[1]) == list:
#            plan_start_date[i] = add_date(origin_data.loc[DU_ID, 
#                           (hwmilestone[1][0],'Plan End Date')], hwmilestone[1][1])
#            actual_start_date[i] = add_date(origin_data.loc[DU_ID, 
#                             (hwmilestone[1][0],'Actual End Date')], hwmilestone[1][1])
    # 结束日期
    for i, hwmilestone in enumerate(map_milestone, start=1):
        if type(hwmilestone[1]) == str:
            plan_end_date[i] = origin_data.loc[DU_ID, (hwmilestone[1],'Plan End Date')]
            actual_end_date[i] = origin_data.loc[DU_ID, (hwmilestone[1],'Actual End Date')]
        elif type(hwmilestone[1]) == list:
            plan_end_date[i] = add_date(origin_data.loc[DU_ID, 
                           (hwmilestone[1][0],'Plan End Date')], hwmilestone[1][1])
            actual_end_date[i] = add_date(origin_data.loc[DU_ID, 
                             (hwmilestone[1][0],'Actual End Date')], hwmilestone[1][1])
        
    plan_end_date = norm_read_date(plan_end_date)
    actual_end_date = norm_read_date(actual_end_date)
     
    
    # 定义duration (plan_end_date - plan_start_date) / (actual_end_date - actual_start_date)
    duration = [' ']*(len(map_milestone)+1)
    
#    # 计算End date， duration
#    for i in range(1, (len(map_milestone)+1)-1):
#    #    plan_end_date[i] = plan_start_date[i+1]
#    #    actual_end_date[i] = actual_start_date[i+1]
#        plan_end_date[i] = end_date_cac(plan_start_date, i)
#        actual_end_date[i] = end_date_cac(actual_start_date, i)
#        
#        plan_duration = duration_cac(plan_start_date[i], plan_end_date[i])
#        actual_duration = duration_cac(actual_start_date[i], actual_end_date[i])
#        duration[i] = plan_duration + '/' + actual_duration
    
     # 计算start date， duration
    for i in range(2, (len(map_milestone)+1)):
        plan_start_date[i] = start_date_cac(plan_end_date, i)
        actual_start_date[i] = start_date_cac(actual_end_date, i)
        
        plan_duration = duration_cac(plan_start_date[i], plan_end_date[i])
        actual_duration = duration_cac(actual_start_date[i], actual_end_date[i])
        duration[i] = plan_duration + '/' + actual_duration
    
     
    # 第一行的start date， duration
    plan_start_date[1] = add_date(plan_end_date[1], -3)
    actual_start_date[1] = add_date(actual_end_date[1], -3)
    plan_duration = duration_cac(plan_start_date[1], plan_end_date[1])
    actual_duration = duration_cac(actual_start_date[1], actual_end_date[1])
    duration[1] = plan_duration + '/' + actual_duration
    
#    # 最后一行的end date， duration
#    plan_end_date[-1] = add_date(plan_start_date[-1], 3)
#    actual_end_date[-1] = add_date(actual_start_date[-1], 3)
#    plan_duration = duration_cac(plan_start_date[-1], plan_end_date[-1])
#    actual_duration = duration_cac(actual_start_date[-1], actual_end_date[-1])
#    duration[-1] = plan_duration + '/' + actual_duration
    
    plan_start_date = norm_read_date(plan_start_date)
    actual_start_date = norm_read_date(actual_start_date)

    
    return [plan_start_date, plan_end_date, actual_start_date, actual_end_date, duration]   
        
   
def write_report_DU_data(worksheet, DU_ID, DU_data, k, headerformat_1, dateformat, normalformat):
    
    plan_start_date = DU_data[0]
    plan_end_date = DU_data[1]
    actual_start_date = DU_data[2]
    actual_end_date = DU_data[3]
    duration = DU_data[4]
    
    # 写单个DU ID的列标签
    headerformat_1.set_align('center')
    worksheet.merge_range(1, 2+k, 1, 6+k, DU_ID, headerformat_1)
    worksheet.merge_range(2, 2+k, 2, 3+k, 'Plan', headerformat_1)
    worksheet.merge_range(2, 4+k, 2, 5+k, 'Actual', headerformat_1)
    worksheet.write(2, 6+k, 'Duration', headerformat_1)
    worksheet.write(3, 2+k, 'Start Date', headerformat_1)
    worksheet.write(3, 3+k, 'End Date', headerformat_1)
    worksheet.write(3, 4+k, 'Start Date', headerformat_1)
    worksheet.write(3, 5+k, 'End Date', headerformat_1)
    worksheet.write(3, 6+k, 'Plan/Actual', headerformat_1)
    
    
    # 写单个DU的start date, end date，duration
    rownspans = [(4, 8), (9, 12), (13, 19), (20, 23), (24, 28), (29, 33), (34, 40), 
                (41, 45), (46, 49), (50, 53), (54, 57)]
    i = 1
    for rownspan in rownspans:
        for rown in range(rownspan[0], rownspan[1]):
            worksheet.write(rown, 2+k, plan_start_date[i], dateformat)
            worksheet.write(rown, 3+k, plan_end_date[i], dateformat)
            worksheet.write(rown, 4+k, actual_start_date[i], dateformat)
            worksheet.write(rown, 5+k, actual_end_date[i], dateformat)
            worksheet.write(rown, 6+k, duration[i], normalformat)
            i += 1
    
    # 设置列的宽度
    date_column_width = 11
    worksheet.set_column(2+k, 6+k, date_column_width)
    worksheet.set_column(7+k, 7+k, 2)

def write_sheet_row_labels(worksheet, headerformat_1, headerformat_2, sumsheet=False):
    # 写表格的前两列行标签
    # 设置行标签
    column_0 = [None, None, 'Sr No.', None, '1', '2', '3', '4', '5', 
                None, None, None, '6', None, None, None, None, '7', '8', 
                '9', None, None, None, '10', None, None, None, None, '11',
                None, None, None, None, '12', None, None, None, None, 
                None, '13', '14', None, None, None, '15', '16', None, None, 
                None, '17', None, None, None, '18', None, '19', '20']
    
    column_1 =[None, 'Site ID', 'Checking Item', None, 'Survey',
               'TSSR', 'HLD', 'LLD/ Change', 'Site Acquisition', '(i) EPC', 
               '(ii) Cable ', '(iii) Pole', 'Civil Work', '(i) Foundation', 
               '(ii) Manhole', '(iii) Cable Laying', '(iv) Pole', 
               'Splicing (Uplink & Cable)', 'MSAN TI (Power & Facility)', 
               'Uplink', '(i) Equipment & Route', '(ii) IP Connectivity', 
               '(iii) Termination & Testing', 'Network Auditing', '(i) MDF', 
               '(ii) DC', '''(iii) Subscriber's Information''', '(iv) OSP Info', 
               'Cable Work', 'Cable Laying', 'Splicing (Uplink & Cable)', 
               'Jumpering/Jointing', 'Termination', 'Test & Configure', 
               '(i) Line / ADSL Testing', '(ii) Configuration /Commission', 
               '(iii) Subscriber/ Line Correctness', '(iv) Line Quality Measurement ', 
               "[Vertical/Numerical / Line & Station Card /MDF Card, etc…]", 
               'Confirmation', 'Plan & Announcement', '(i) Notification', 
               '(ii) Announcement', '(iii) Report', 'MOP', 'Mobilization', 
               '(i) MPT (CTE, OP, IT)', '(ii) State & Region', 
               '(iii) Vendor / LSP/ Subcon', 'Monitoring', '(i) Design & Specification', 
               '(ii) Equipment readiness', '(iii) Progess', 'Migration', 
               'Record and Report', 'Rectification', 'Documentation']
    
    # 写入行标签
    k = 1 if sumsheet else 0
    headerformat_1.set_align('left')
    for rown in range(1, 57):
        worksheet.write(rown+k, 0+k, column_0[rown], headerformat_1)
        if column_0[rown]:
            worksheet.write(rown+k, 1+k, column_1[rown], headerformat_1)
        else:
            worksheet.write(rown+k, 1+k, column_1[rown], headerformat_2)
   
    headerformat_1.set_align('center')
    if not sumsheet:
        worksheet.merge_range(1, 0, 1, 1, 'Site Id---->', headerformat_1)
    worksheet.merge_range(2, 0+k, 3+k, 0+k, 'Sr No.', headerformat_1)
    worksheet.merge_range(2, 1+k, 3+k, 1+k, 'Checking Item', headerformat_1)
    worksheet.set_column(0+k, 0+k, 8)
    worksheet.set_column(1+k, 1+k, 25)
    
    if sumsheet: 
        worksheet.write(1, 2, 'Exchange Area', headerformat_1)
        worksheet.set_row(4, 40)
        worksheet.set_column(0, 0, 2)

def write_summary_exchange_data(sheet_summary, exchange, exchange_data, headerformat_1, normalformat, k):
        
    sheet_summary.merge_range(1, 3+k, 1, 17+k, exchange, headerformat_1)
    sheet_summary.merge_range(2, 3+k, 2, 9+k, 'Plan', headerformat_1)
    sheet_summary.merge_range(2, 10+k, 2, 16+k, 'Actual', headerformat_1)
    sheet_summary.write(2, 17+k, 'Duration', headerformat_1)
    sheet_summary.merge_range(3, 3+k, 3, 5+k, 'Scope totally', headerformat_1)
    sheet_summary.merge_range(3, 6+k, 3, 8+k, 'Start/End date', headerformat_1)
    sheet_summary.merge_range(3, 10+k, 3, 13+k, ' ', headerformat_1)
    sheet_summary.merge_range(3, 14+k, 3, 16+k, 'Start/End date', headerformat_1)
    sheet_summary.write(3, 17+k, ' ', headerformat_1)
    col = ['Indoor', 'Outdoor', 'Pole', 'Day', 'Month', 'Year', 
           'Standard Deviation', 'Finished Sites Count', 'Finished Sites List',
           'Pending Site Count', 'Pending Site List', 'Day', 'Month', 'Year', 
           'Plan/Actual']
    for i in range(3, 18):
        sheet_summary.write(4, i+k, col[i-3], headerformat_1)
    normal_column_width = 10
    sheet_summary.set_column(3+k, 17+k, normal_column_width)    
    sheet_summary.set_column(18+k, 18+k, 2)    
    
    rownspans = [(4, 8), (9, 12), (13, 19), (20, 23), (24, 28), (29, 33), (34, 40), 
            (41, 45), (46, 49), (50, 53), (54, 57)]
    i = 1
    for rownspan in rownspans:
        for rown in range(rownspan[0], rownspan[1]):
            for j in range(0, 15):
                sheet_summary.write(rown+1, 3+k+j, exchange_data[j][i-1], normalformat)
            i += 1