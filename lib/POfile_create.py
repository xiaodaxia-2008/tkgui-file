# -*- coding: utf-8 -*-
"""
Created on Wed Apr  4 10:39:31 2018

@author: x00428488
"""
import pandas as pd, numpy as np
import tkinter as tk
import datetime
import time
import os, logging, sys
from general_functions_for_ML import DUconfigsum, write_df_to_excel

logging.basicConfig(level=logging.DEBUG, format='%(msg)s')
## 定义常量，和用户数有关的PO
user_related_po = np.array([[8818068409, 8818068441, 8818068444, 8818068432, 8818068428,
                                8818068486, 8818068397, 8818068447, 8818068468],
                           [8818068427, 8818068512, 8818068473, 8818068467, 8818068420,
                            8818068425, 8818068412, 8818068507, 8818068503],
                           [8818068504, 8818068509, 8818068460, 8818068430, 8818068418,
                            8818068459, 8818068404, 8818068442, 8818068415],
                           [8818068489, 8818068482, 8818068399, 8818068471, 8818068436,
                            8818068493, 8818068479, 8818068433, 8818068391]]).astype(str)


def PO_generate(get_text_contents, progressBar, status_info):
    try:
        progressBar.start()
        paths = get_text_contents()
        for i, path in enumerate(paths):
            paths[i] = path.replace('\n', '')
            logging.debug(str(path[i]))
            
        CCM_shipped_file, item_detail_file, save_file_path = paths
        
        if not os.path.exists(save_file_path):
            os.makedirs(save_file_path)
            
        starttime = time.time()
        
        ##处理Item Details文件，去掉重复，以Item Code 为索引
        logging.debug('读取Item Detials 文件。。。')
        status_info.insert('2.0', '\nloading data...')
#        item_detail_file = './templateSourceData/ItemDetails.xlsx'  
        ##设置Item Details映射信息的文件路径
        Item_Details = pd.read_excel(item_detail_file, sheet_name='Sheet1', header=0)
#        Item_Details.drop_duplicates('Item Code',inplace=True) # 如果一个物料对应多个PO，不应该删除重复
        Item_Details.dropna(axis=0, subset=['Item Code'], inplace=True)
        Item_Details.set_index('Item Code',inplace=True)
        Item_Details.fillna('', inplace=True)
        
        ##读取CCM中的站点发货配置数据
        logging.debug('读取CCM 发货数据。。。')       
        CCM_shipped_file = './templateSourceData/56A03KN_DCConfiguration__29993670_20180410131941157.xlsx' ##设置CCM数据文件的路径
#        CCM_shipped_file = 'D:/MLtool/MLPO/templateSourceData/5640048_DCConfiguration__29993557_20180410130815547.xlsx' ##设置CCM数据文件的路径
        shipped_items = pd.read_excel(CCM_shipped_file, 1, header=0)
        shipped_items.drop(['Contract Register Date', 
                           'DU Name', 'Customer Site Name', 'Product',
                           'UOM', 'Bpart Line ID',
                           'Substitute ID', 'Survey Aux.', 'Current Qty.', 'Unshipped DN Qty.',
                           'Delivery Config Engineer', 'Spart Number',
                           'Split Flag', 'Delivery Mode'], axis=1, inplace=True)  ##去掉无用的列
        ##去掉有nan值的行，CCM里面的数据有的Item Code是空的，去掉这样的行
        shipped_items.dropna(axis=0,subset=['Part Number*'], inplace=True)         
        logging.debug('CCM 发货数据清理中。。。')
        status_info.insert('2.0', '\nclearing CCM data...')        
        ##去掉Item Code重复的行，对Item Code的总数汇总，利用了DUconfigsum()函数
        shipped_items = DUconfigsum(shipped_items)  
        shipped_items.fillna('', inplace=True)
        
        
        ## 读取ISDP中站点的基本信息
        ISDP_file = './templateSourceData/ISDP_site_information.xlsx'
        bpa_file = './templateSourceData/BPA_PSTN_Phase2.xlsx'
        site_info = pd.read_excel(ISDP_file, sheet_name=0, header=0, index_col='DU ID',
                                  skiprows=[1], usecols=[x for x in range(23)])
        site_info.fillna('', inplace=True)
        bpa_info = pd.read_excel(bpa_file, sheet_name=2)
        bpa_info.rename(columns={'Item': 'PO Item Code', 'Description*': 'PR Line Description',
                                 'Qty': 'PR Line Quantity', 'UOM*': 'UOM'}, inplace=True)
        bpa_info['PO Item Code'] = bpa_info['PO Item Code'].astype(str)
        bpa_info.set_index('PO Item Code', drop=False, inplace=True)
        bpa_info['PR Line Quantity'] = 1
        delivery_po = bpa_info[bpa_info['Classification*'].str.contains('Logistics')]
        
        
        ##分别一次选取CCM数据中的一个DU的发货数据进行处理
        logging.debug('开始生成PO文件。。。')        
        status_info.insert('2.0', '\n'+ 'Start creating PO files...')
        
#        DU_IDs = shipped_items['DU ID'].unique()
        DU_IDs = site_info.index.unique()
        PO_file_result = pd.DataFrame([])
        for DU_ID in DU_IDs:
            DUdata = shipped_items[shipped_items['DU ID']==DU_ID].copy()
            DUdata.set_index('Part Number*', drop=False, inplace=True)
            POinfo = pd.DataFrame([])   
            # 下面计算根据站点类型信息带出来的分包PO
            # 根据站点所属exchange关联出运输PO，默认为1000kg
            exchange = site_info.loc[DU_ID, 'Exchange']
            if exchange:
                DU_delivery_po = delivery_po[delivery_po.Remark.str.contains(exchange[1:],
                                                                             na=False)]
                POinfo = POinfo.append(DU_delivery_po[['PO Item Code', 'PR Line Description', 
                                                       'PR Line Quantity', 'UOM']])
            # 根据用户数量关联出Subscriber cable installation，
            # User Inventory Clearance of MDF，Subscriber Cut-over，Acceptance  test         
            user_num = site_info.loc[DU_ID, 'Capacity']
            if index_po(user_num):
                po_index = user_related_po[:, index_po(user_num)]
                POinfo = POinfo.append(bpa_info.loc[po_index, ['PO Item Code', 'PR Line Description',
                                                               'PR Line Quantity', 'UOM']])
            
            # 根据铜缆适量，确定铜缆铺放PO：Laying of  Copper Cable(50 pairs-0.4) in aerial
            # Laying of  Copper Cable(100 pairs-0.4) in duct
            
            
            # 以下计算可以根据站点物料关联出的分包PO
            if DUdata['Part Number*'].any():
                for PartNum in DUdata['Part Number*']:
                    try:
                        POinfo = POinfo.append(Item_Details.loc[PartNum,
                            ['PO Item Code', 'PR Line Description', 
                             'PR Line Quantity', 'UOM']])
                        
                    except KeyError:
                        logging.debug('No such item code in Item Details table')
                        pass
            status_info.insert('2.0', '\n'+ 'computing '+str(DU_ID))
            logging.debug('computing '+str(DU_ID))
            DUdata.rename(columns={'DU ID': 'DU_ID'}, inplace=True)
            POfile = DUdata[['HW Contract NO.', 'Customer Site ID', 'DU_ID', 
                             'Product Instance', 'Part Number*', 
                             'Part Description', 'Actual Qty.*']]
            if POinfo.any().any():
                POinfo['DU ID'] = DU_ID
                POinfo['Supplier Name'] = site_info.loc[DU_ID, 'Subcontractor']
                POfile = POfile.join(POinfo, how='outer', sort=False)
                PO_file_result = PO_file_result.append(POfile)
#            if PO_file_result.columns.values[0] != 'HW Contract NO.':
#                break
        
#        save_file_path = './Results'
        POfile_path = save_file_path + '/POfile_' + \
                        datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')+'.xlsx'
        # POfile.to_excel(POfile_path, index=False, freeze_panes=[1,1])
     
        PO_file_result.fillna('', inplace=True)
        write_df_to_excel(PO_file_result, POfile_path, 'PO File')
        
        
        status_info.insert('2.0', '\nfile saved to: '+ POfile_path)
        progressBar.stop()
        endtime = time.time()
        costtime = str(endtime-starttime)
        result_mesg = 'Finished creating PO file, cost '+costtime+' seconds!'
        logging.debug(result_mesg)
        status_info.insert('2.0', '\n'+ str(result_mesg))
        choice = tk.messagebox.askyesno(title='Finished', 
                                        message=result_mesg+'\nDo you want to open it?')
        if choice:
            os.startfile(POfile_path)
        
    except Exception as e:
        logging.debug(e)
        progressBar.stop()
        tk.messagebox.showerror(title='Error', message=e)
        status_info.insert('2.0', '\n'+ str(e))
        
def index_po(user_num):
    if user_num:
        index_po = int((user_num-1)/128)
        index_po = 4 if user_num > 512  else index_po
        index_po = 5 if user_num > 768  else index_po
        index_po = 6 if user_num > 1024 else index_po
        index_po = 7 if user_num > 2000 else index_po
        index_po = 8 if user_num > 3000 else index_po
        return index_po
    else:
        return None