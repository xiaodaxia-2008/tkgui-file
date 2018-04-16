# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd, numpy as np
from  general_functions_for_transfer_report_v4 import set_mutiIndex
import datetime, logging, time, os
import tkinter as tk


logging.basicConfig(level=logging.DEBUG, format='%(message)s')

def merge_site_date_eles(col):
    '''
    返回 col 中的属于日期的那个，如果都是日期， 返回最后一个
    '''
    for i in range(len(col)):
        
        if type(col[i]) == datetime.datetime or type(col[i]) == pd.Timestamp:
            date = col[i]
        else:
            date = ' '
        
    return date

def merge_site_date_rows(df, *status_info):
    level2_values=['Plan End Date',	 'Actual End Date']
    dfdate_set = []
#    dfcopy = pd.DataFrame([], index=df.index, columns=df.columns)
    dfcopy = df.copy().astype(np.object)
    for colname in level2_values:
        dfdate_set.append(df.xs(colname, level=1, axis=1, drop_level=False))
    
    for dfdate in dfdate_set:
        for datecol in dfdate.columns.values:
            dfdatecol = dfdate.loc[:, datecol]
            for site in dfdatecol.index.levels[0]:
                dfcopy.loc[site, datecol] = merge_site_date_eles(dfdatecol.loc[site]) 
            logging.debug(datecol)   
            if status_info:
                status_info[0].insert('2.0', '\n'+str(datecol))
    dfcopy.drop_duplicates(subset=('Customer Site ID', ''), inplace=True)
    return dfcopy





origin_file = 'D:/MLtool/MLPO/templateSourceData/56A03KN_客户报表_20180408.xlsx'


def create_JF_Sheet1(get_text_contents, progressBar, status_info):
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
#        origin_file = 'D:/MLtool/ISDPreport/x00428488_56A03KN_20180403.xlsx'
        ISDP_data = pd.read_excel(origin_file, sheet_name='Rollout Plan',
                                  header=[0, 1])
        status_info.insert('2.0', '\nclearing data...')
        set_mutiIndex(ISDP_data, level2_values=['Plan Start Date', 'Plan End Date',	
                                                'Actual Start Date', 'Actual End Date',
                                                'Owner'])
        ISDP_data.set_index(['Customer Site ID', 'DU ID'], drop=False, inplace=True)
        status_info.insert('2.0', '\nmerge DUs...')
        df_mergedate = merge_site_date_rows(ISDP_data, status_info)
#        df_mergedate = merge_site_date_rows(ISDP_data)
        status_info.insert('2.0', '\ncreating sheet1...')
#        colnames_jianfang = [('Exchange', ''), ('Customer Site ID', ''), ('Site Type', ''), 
#                    ('Site Survey', 'Plan End Date'), ('Site Survey', 'Actual End Date'),
#                    ('TSSR Ready', 'Plan Start Date'), ('TSSR Ready', 'Actual Start Date')
#                    ]   ## 在此填入对应的ISDP的标题
        keepcolumns = [('Customer Site ID', ''), ('Customer Site Name', ''),
                       ('DU Name', ''), ('Delivery Area', ''),
                       ('Phase', ''), ('City', ''), ('Township', ''), ('Exchange', ''),
                       ('Latitude', ''), ('Longitude', ''), ('Site Address', ''),
                       ('Capacity', ''), ('Site Type', ''), ('MSAN Type', ''),
                       ('Exchange solution', ''), ('HQ solution', ''),
                       ('LLD Solution', ''), ('LLD HQ approval-date', ''),
                       ('Site owner', ''), ('Subcontractor', ''), ('Zone Manager', ''),
                       ('site remark', ''), ('site status', ''), ('CW OR TE', ''),
                       ('Site Survey', 'Plan End Date'),
                       ('Site Survey', 'Actual End Date'),
                       ('TSSR Ready', 'Plan End Date'), ('TSSR Ready', 'Actual End Date'),
                       ('TSSR Approve', 'Plan End Date'),
                       ('TSSR Approve', 'Actual End Date'),
                       ('CDC Township Officer Visit', 'Plan End Date'),
                       ('CDC Township Officer Visit', 'Actual End Date'),
                       ('XCDC Application Submit', 'Plan End Date'),
                       ('XCDC Application Submit', 'Actual End Date'),
                       ('xCDC approve', 'Plan End Date'),
                       ('xCDC approve', 'Actual End Date'),
                       ('Draft LLD Finish', 'Plan End Date'),
                       ('Draft LLD Finish', 'Actual End Date'),
                       ('LLD Huawei Approve', 'Plan End Date'),
                       ('LLD Huawei Approve', 'Actual End Date'),
                       ('LLD Exchange Approve', 'Plan End Date'),
                       ('LLD Exchange Approve', 'Actual End Date'),
                       ('LLD HQ Approval', 'Plan End Date'),
                       ('LLD HQ Approval', 'Actual End Date'),
                       ('Finding Copper', 'Plan End Date'),
                       ('Finding Copper', 'Actual End Date'),
                       ('Copper Cable Laying', 'Plan End Date'),
                       ('Copper Cable Laying', 'Actual End Date'),
                       ('Inventory Check', 'Plan End Date'),
                       ('Inventory Check', 'Actual End Date'),
                       ('PA Application Submit', 'Plan End Date'),
                       ('PA Application Submit', 'Actual End Date'),
                       ('Power application approve', 'Plan End Date'),
                       ('Power application approve', 'Actual End Date'),
                       ('Power Pole Installation', 'Plan End Date'),
                       ('Power Pole Installation', 'Actual End Date'),
                       ('Meter installation', 'Plan End Date'),
                       ('Meter installation', 'Actual End Date'),
                       ('Power Connect', 'Plan End Date'),
                       ('Power Connect', 'Actual End Date'),
                       ('CW Start', 'Plan End Date'), ('CW Start', 'Actual End Date'),
                       ('Excavation', 'Plan End Date'), ('Excavation', 'Actual End Date'),
                       ('Lean Concrete', 'Plan End Date'),
                       ('Lean Concrete', 'Actual End Date'),
                       ('MH Construction', 'Plan End Date'),
                       ('MH Construction', 'Actual End Date'),
                       ('Rebar installation and Form work', 'Plan End Date'),
                       ('Rebar installation and Form work', 'Actual End Date'),
                       ('Casting', 'Plan End Date'), ('Casting', 'Actual End Date'),
                       ('MSAN foundation complete', 'Plan End Date'),
                       ('MSAN foundation complete', 'Actual End Date'),
                       ('Back filling', 'Plan End Date'),
                       ('Back filling', 'Actual End Date'),
                       ('Civil Work Completed', 'Plan End Date'),
                       ('Civil Work Completed', 'Actual End Date'),
                       ('Smart QC for CW', 'Plan End Date'),
                       ('Smart QC for CW', 'Actual End Date'),
                       ('DN Ready', 'Plan End Date'), ('DN Ready', 'Actual End Date'),
                       ('IP Network configuration', 'Plan End Date'),
                       ('IP Network configuration', 'Actual End Date'),
                       ('IP Uplink site Ready', 'Plan End Date'),
                       ('IP Uplink site Ready', 'Actual End Date'),
                       ('Material On Site', 'Plan End Date'),
                       ('Material On Site', 'Actual End Date'),
                       ('Equiment Installation', 'Plan End Date'),
                       ('Equiment Installation', 'Actual End Date'),
                       ('Installation Completed', 'Plan End Date'),
                       ('Installation Completed', 'Actual End Date'),
                       ('Termination', 'Plan End Date'),
                       ('Termination', 'Actual End Date'),
                       ('Y Splicing', 'Plan End Date'), ('Y Splicing', 'Actual End Date'),
                       ('Inventory Clearance', 'Plan End Date'),
                       ('Inventory Clearance', 'Actual End Date'),
                       ('Software Commisioning', 'Plan End Date'),
                       ('Software Commisioning', 'Actual End Date'),
                       ('Jumper wire', 'Plan End Date'),
                       ('Jumper wire', 'Actual End Date'), ('Dial Up', 'Plan End Date'),
                       ('Dial Up', 'Actual End Date'),
                       ('Migration ready', 'Plan End Date'),
                       ('Migration ready', 'Actual End Date'),
                       ('Smart QC for TE', 'Plan End Date'),
                       ('Smart QC for TE', 'Actual End Date'),
                       ('Migration Approval', 'Plan End Date'),
                       ('Migration Approval', 'Actual End Date'),
                       ('Migration', 'Plan End Date'), ('Migration', 'Actual End Date'),
                       ('Call test after migration', 'Plan End Date'),
                       ('Call test after migration', 'Actual End Date')]
#        df_jianfang = pd.DataFrame([], index=df_mergedate.index, columns=[['A'],[1]])
#        for colname in keepcolumns:
#            df_jianfang = df_jianfang.join(df_mergedate[colname])
#        df_jianfang.drop('A', axis=1, level=0, inplace=True)
        df_jianfang = df_mergedate[keepcolumns]
        df_jianfang.reset_index(level=[0, 1], drop=True, inplace=True)
        df_jianfang.index.name = None
        
#        save_file_path = 'D:/MLtool/ISDPreport/Publish_verson/version_5/'
        filename = save_file_path + '/Sheet1_'+ \
                    datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S') +'.xlsx'
        df_jianfang.to_excel(filename, index=True, freeze_panes=[2, 2])
        status_info.insert('2.0', '\nfile saved to: '+ filename)
        progressBar.stop()
        endtime = time.time()
        costtime = str(endtime-starttime)
        result_mesg = 'Finished create Sheet1 file, cost '+costtime+' seconds!'
        logging.debug(result_mesg)
        status_info.insert('2.0', '\n'+ str(result_mesg))
        choice = tk.messagebox.askyesno(title='Finished', 
                                        message=result_mesg+'\nDo you want to open it?')
        if choice:
            os.startfile(filename)
    
    except Exception as e:
        logging.debug(e)
        progressBar.stop()
        tk.messagebox.showerror(title='Error', message=e)
        status_info.insert('2.0', '\n'+ str(e))




nondatecols = ['Customer Site ID', 'Customer Site Name', 'DU ID', 'DU Name', 
               'Site Owner', 'Daily Plan', 'site status', 'site remark', 
               'Zone Manager', 'Exchange', 'Site Type', 'Subcontractor']


ISDP_colnames = [('Customer Site ID', ''), ('Customer Site Name', ''), ('DU ID', ''),
       ('DU Name', ''), ('Site Owner', ''), ('Daily Plan', ''),
       ('site status', ''), ('site remark', ''), ('Zone Manager', ''),
       ('Exchange', ''), ('Site Type', ''), ('Subcontractor', ''),
       ('Site Scope', 'Plan End Date'), ('Site Scope', 'Actual End Date'),
       ('Site Scope', 'Owner'), ('Site Survey', 'Plan End Date'),
       ('Site Survey', 'Actual End Date'), ('Site Survey', 'Owner'),
       ('TSSR Ready', 'Plan Start Date'), ('TSSR Ready', 'Plan End Date'),
       ('TSSR Ready', 'Actual Start Date'),
       ('TSSR Ready', 'Actual End Date'), ('TSSR Ready', 'Owner'),
       ('TSSR Approve', 'Plan End Date'),
       ('TSSR Approve', 'Actual End Date'), ('TSSR Approve', 'Owner'),
       ('Draft LLD Finish', 'Plan End Date'),
       ('Draft LLD Finish', 'Actual End Date'),
       ('Draft LLD Finish', 'Owner'),
       ('LLD Huawei Approve', 'Plan End Date'),
       ('LLD Huawei Approve', 'Actual End Date'),
       ('LLD Huawei Approve', 'Owner'),
       ('LLD Exchange Approve', 'Plan End Date'),
       ('LLD Exchange Approve', 'Actual End Date'),
       ('LLD Exchange Approve', 'Owner'),
       ('LLD HQ Approval', 'Plan End Date'),
       ('LLD HQ Approval', 'Actual End Date'),
       ('LLD HQ Approval', 'Owner'), ('HLD Approve', 'Plan End Date'),
       ('HLD Approve', 'Actual End Date'), ('HLD Approve', 'Owner'),
       ('Ready for Delivery', ''),
       ('XCDC Application Submit', 'Plan End Date'),
       ('XCDC Application Submit', 'Actual End Date'),
       ('XCDC Application Submit', 'Owner'),
       ('CDC Township Officer Visit', 'Plan End Date'),
       ('CDC Township Officer Visit', 'Actual End Date'),
       ('CDC Township Officer Visit', 'Owner'),
       ('xCDC approve', 'Plan End Date'),
       ('xCDC approve', 'Actual End Date'), ('xCDC approve', 'Owner'),
       ('Inventory Check', 'Plan End Date'),
       ('Inventory Check', 'Actual End Date'),
       ('Inventory Check', 'Owner'), ('CW Start', 'Plan End Date'),
       ('CW Start', 'Actual End Date'), ('CW Start', 'Owner'),
       ('Excavation', 'Plan End Date'), ('Excavation', 'Actual End Date'),
       ('Excavation', 'Owner'), ('Lean Concrete', 'Plan End Date'),
       ('Lean Concrete', 'Actual End Date'), ('Lean Concrete', 'Owner'),
       ('Rebar installation and Form work', 'Plan End Date'),
       ('Rebar installation and Form work', 'Actual End Date'),
       ('Rebar installation and Form work', 'Owner'),
       ('Casting', 'Plan End Date'), ('Casting', 'Actual End Date'),
       ('Casting', 'Owner'), ('Back filling', 'Plan End Date'),
       ('Back filling', 'Actual End Date'), ('Back filling', 'Owner'),
       ('Civil Work Completed', 'Plan End Date'),
       ('Civil Work Completed', 'Actual End Date'),
       ('Civil Work Completed', 'Owner'),
       ('Smart QC for CW', 'Plan End Date'),
       ('Smart QC for CW', 'Actual End Date'),
       ('Smart QC for CW', 'Owner'),
       ('PA Application Submit', 'Plan End Date'),
       ('PA Application Submit', 'Actual End Date'),
       ('PA Application Submit', 'Owner'),
       ('Power application approve', 'Plan End Date'),
       ('Power application approve', 'Actual End Date'),
       ('Power application approve', 'Owner'),
       ('Power Installation', 'Plan End Date'),
       ('Power Installation', 'Actual End Date'),
       ('Power Installation', 'Owner'), ('Power Connect', 'Plan End Date'),
       ('Power Connect', 'Actual End Date'), ('Power Connect', 'Owner'),
       ('Meter installation', 'Plan End Date'),
       ('Meter installation', 'Actual End Date'),
       ('Meter installation', 'Owner'),
       ('Finding Copper', 'Plan End Date'),
       ('Finding Copper', 'Actual End Date'), ('Finding Copper', 'Owner'),
       ('Material On Site', 'Plan End Date'),
       ('Material On Site', 'Actual End Date'),
       ('Material On Site', 'Owner'),
       ('Copper Cable Laying', 'Plan End Date'),
       ('Copper Cable Laying', 'Actual End Date'),
       ('Copper Cable Laying', 'Owner'),
       ('Equiment Installation', 'Plan End Date'),
       ('Equiment Installation', 'Actual End Date'),
       ('Equiment Installation', 'Owner'),
       ('Termination', 'Plan End Date'),
       ('Termination', 'Actual End Date'), ('Termination', 'Owner'),
       ('Y Splicing', 'Plan End Date'), ('Y Splicing', 'Actual End Date'),
       ('Y Splicing', 'Owner'), ('Inventory Clearance', 'Plan End Date'),
       ('Inventory Clearance', 'Actual End Date'),
       ('Inventory Clearance', 'Owner'),
       ('Software Commisioning', 'Plan End Date'),
       ('Software Commisioning', 'Actual End Date'),
       ('Software Commisioning', 'Owner'),
       ('Jumper wire', 'Plan End Date'),
       ('Jumper wire', 'Actual End Date'), ('Jumper wire', 'Owner'),
       ('Installation Completed', 'Plan Start Date'),
       ('Installation Completed', 'Plan End Date'),
       ('Installation Completed', 'Actual Start Date'),
       ('Installation Completed', 'Actual End Date'),
       ('Installation Completed', 'Owner'),
       ('Smart QC for TE', 'Plan End Date'),
       ('Smart QC for TE', 'Actual End Date'),
       ('Smart QC for TE', 'Owner'), ('Dial Up', 'Plan Start Date'),
       ('Dial Up', 'Plan End Date'), ('Dial Up', 'Actual Start Date'),
       ('Dial Up', 'Actual End Date'), ('Dial Up', 'Owner'),
       ('Migration ready', 'Plan End Date'),
       ('Migration ready', 'Actual End Date'),
       ('Migration ready', 'Owner'),
       ('Migration Approval', 'Plan End Date'),
       ('Migration Approval', 'Actual End Date'),
       ('Migration Approval', 'Owner'), ('Migration', 'Plan Start Date'),
       ('Migration', 'Plan End Date'), ('Migration', 'Actual Start Date'),
       ('Migration', 'Actual End Date'), ('Migration', 'Owner'),
       ('IP Network configuration', 'Plan End Date'),
       ('IP Network configuration', 'Actual End Date'),
       ('IP Network configuration', 'Owner'),
       ('DN Ready', 'Plan End Date'), ('DN Ready', 'Actual End Date'),
       ('DN Ready', 'Owner'), ('CW OR TE', ''), ('Phase', '')]


Jianfang_colnames = [('Exchange', ''), ('Site ID', ''), ('Site Type', ''),
       ('Equipment Type', ''), ('MSAN Type', ''), ('Target Date', ''),
       ('Priority', ''), ('Lat', ''), ('Long', ''), ('Township', ''),
       ('Site Address', ''), ('Subcon', ''), ('HW Zone Engineer', ''),
       ('Contact Number', ''), ('Site Survey Completion', 'Plan'),
       ('Site Survey Completion', 'Actual'), ('TSSR approval', 'Plan'),
       ('TSSR approval', 'Actual'), ('LLD Ready', 'Plan'),
       ('LLD Ready', 'Actual'), ('LLD Approval', 'Plan'),
       ('LLD Approval', 'Actual'), ('ROW Application', 'Plan'),
       ('ROW Application', 'Actual'), ('ROW Approval', 'Plan'),
       ('ROW Approval', 'Actual'), ('PA Application', 'Plan'),
       ('PA Application', 'Actual'), ('PA Approval', 'Plan'),
       ('PA Approval', 'Actual'), ('PA Ready', 'Plan'),
       ('PA Ready', 'Actual'), ('Copper finding', 'Plan'),
       ('Copper finding', 'Actual'), ('CW Start Plan', 'Plan'),
       ('CW Start Plan', 'Actual'),
       ('MSAN/DC/Manhole foundation Complete', 'Plan'),
       ('MSAN/DC/Manhole foundation Complete', 'Actual'),
       ('Subscriber information checking', 'Plan'),
       ('Subscriber information checking', 'Actual'),
       ('DN send to MPT', 'Plan'), ('DN send to MPT', 'Actual'),
       ('Delivery Plan to Site [MOS]', 'Plan'),
       ('Delivery Plan to Site [MOS]', 'Actual'),
       ('Fiber Pole Eruption Completion', 'Plan'),
       ('Fiber Pole Eruption Completion', 'Actual'),
       ('OSP Route Completion', 'Plan'),
       ('OSP Route Completion', 'Actual'),
       ('Line Quality Measurement(TE start)', 'Plan'),
       ('Line Quality Measurement(TE start)', 'Actual'),
       ('TE Completion', 'Plan'), ('TE Completion', 'Actual'),
       ('Copper Cable laying', 'Plan'), ('Copper Cable laying', 'Actual'),
       ('Y-splicing', 'Plan'), ('Y-splicing', 'Actual'),
       ('Termination', 'Plan'), ('Termination', 'Actual'),
       ('Software Commisioning', 'Plan'),
       ('Software Commisioning', 'Actual'), ('Service Test', 'Plan'),
       ('Service Test', 'Actual'), ('Jumpering wire', 'Plan'),
       ('Jumpering wire', 'Actual'), ('Quality check', 'Plan'),
       ('Quality check', 'Actual'), ('Migration Ready', 'Plan'),
       ('Migration Ready', 'Actual'), ('Migration', 'Plan'),
       ('Migration', 'Actual'), ('TE PAC Completion', 'Plan'),
       ('TE PAC Completion', 'Actual')]


