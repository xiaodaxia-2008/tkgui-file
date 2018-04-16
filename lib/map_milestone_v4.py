# -*- coding: utf-8 -*-
"""
Created on Thu Mar 29 17:58:10 2018

@author: x00428488
"""
"""
用于修改客户的mile stone和华为mile stone的映射关系
"""    

## 建立客户milestone和hw milestone 的映射关系, 请用hw milestone替换下面的数值
## 左侧是客户的milestone， 右侧是对应的华为milestone，如果需要加减几天，请按照格式操作
###################### 注意按照以下格式修改 #########################
## ('TSSR', 'TSSR approval') 表示客户的TSSR对应华为的TSSR Approval
## ('LLD/ Change', ['LLD Ready', 1]),  表示客户的LLD/Change 对应着华为的LLD Ready 加1天

map_milestone = [('Survey', 'Site Survey'),
                 ('TSSR', 'TSSR Approve'),
                 ('HLD', 'Draft LLD Finish'),
                 ('LLD/ Change', 'LLD HQ Approval'),  # 某个值加1天
                 ('(i) EPC', 'Finding Copper'),  
                 ('(ii) Cable ', 'Finding Copper'),
                 ('(iii) Pole(SA)', 'Finding Copper'),
                 ('(i) Foundation', 'CW Start'),
                 ('(ii) Manhole', 'CW Start'),
                 ('(iii) Cable Laying', 'CW Start'), #某个值减1天
                 ('(iv) Pole(CW)', 'CW Start'),
                 ('Splicing (Uplink & Cable)', 'MSAN foundation complete'),
                 ('MSAN TI (Power & Facility)', 'MSAN foundation complete'),
                 ('(i) Equipment & Route', 'Smart QC for CW'),
                 ('(ii) IP Connectivity', 'Smart QC for CW'),
                 ('(iii) Termination & Testing', 'Smart QC for CW'),
                 ('(i) MDF', 'IP Network configuration'),
                 ('(ii) DC', 'IP Network configuration'),
                 ("(iii) Subscriber's Information", 'IP Network configuration'),
                 ('(iv) OSP Info', 'IP Network configuration'),
                 ('(i) Cable Laying', 'IP Uplink site Ready'),
                 ('(ii) Splicing (Uplink & Cable)', 'IP Uplink site Ready'),
                 ('(iii) Jumpering/Jointing', 'IP Uplink site Ready'),
                 ('(iv) Termination', 'IP Uplink site Ready'),
                 ('(i) Line / ADSL Testing', 'Equiment Installation'),
                 ('(ii) Configuration /Commission', 'Equiment Installation'),
                 ('(iii) Subscriber/ Line Correctness', 'Equiment Installation'),
                 ('(iv) Line Quality Measurement ', 'Equiment Installation'),
                 ('(v) [Vertical/Numerical / Line & Station Card /MDF Card, etc…]', 'Equiment Installation'),
                 ('Confirmation', 'Termination'),
                 ('(i) Notification', 'Termination'),
                 ('(ii) Announcement', 'Termination'),
                 ('(iii) Report', 'Termination'),
                 ('MOP', 'Termination'),
                 ('(i) MPT (CTE, OP, IT)', 'Y Splicing'),
                 ('(ii) State & Region', 'Y Splicing'),
                 ('(iii) Vendor / LSP/ Subcon', 'Software Commisioning'),
                 ('(i) Design & Specification', 'Jumper wire'),
                 ('(ii) Equipment readiness', 'Migration ready'),
                 ('(iii) Progess', 'Smart QC for TE'),
                 ('(i) Record and Report', 'Migration Approval'),
                 ('Rectification', 'Migration'),
                 ('Documentation', 'Call test after migration')]



## 以下分别是客户和华为的milestone

## 定义客户的mile stone， 除非客户模板改变，否则请勿修改
customer_milestone =['Survey', 'TSSR', 'HLD', 'LLD/ Change', 
                     '(i) EPC', '(ii) Cable ', '(iii) Pole(SA)', 
                     '(i) Foundation', '(ii) Manhole', '(iii) Cable Laying', 
                     '(iv) Pole(CW)', 
                   'Splicing (Uplink & Cable)', 
                   'MSAN TI (Power & Facility)', 
                   '(i) Equipment & Route', '(ii) IP Connectivity', 
                   '(iii) Termination & Testing', 
                   '(i) MDF', '(ii) DC', '''(iii) Subscriber's Information''', 
                   '(iv) OSP Info', 
                   '(i) Cable Laying', '(ii) Splicing (Uplink & Cable)', 
                   '(iii) Jumpering/Jointing', '(iv) Termination',
                   '(i) Line / ADSL Testing', '(ii) Configuration /Commission', 
                   '(iii) Subscriber/ Line Correctness', 
                   '(iv) Line Quality Measurement ', 
                   "(v) [Vertical/Numerical / Line & Station Card /MDF Card, etc…]", 
                   'Confirmation', 
                   '(i) Notification', '(ii) Announcement', '(iii) Report', 
                   'MOP',
                   '(i) MPT (CTE, OP, IT)', '(ii) State & Region', 
                   '(iii) Vendor / LSP/ Subcon', '(i) Design & Specification', 
                   '(ii) Equipment readiness', '(iii) Progess',
                   '(i) Record and Report', 
                   'Rectification', 'Documentation']


# 定义华为的mile stone， 除非模板改变，否则请勿修改
hw_milestone= ['Site Survey Completion', 'TSSR approval', 'LLD Ready', 
               'LLD Approval', 'ROW Application', 
             'ROW Approval', 'PA Application', 'PA Approval', 'PA Ready', 
             'Copper finding', 'CW Start Plan', 'MSAN/DC/Manhole foundation Complete', 
             'Subscriber information checking', 'DN send to MPT', 
             'Delivery Plan to Site [MOS]', 'Fiber Pole Eruption Completion', 
             'OSP Route Completion', 'Line Quality Measurement(TE start)', 
             'TE Completion', 'Copper Cable laying', 'Y-splicing', 'Termination', 
             'Software Commisioning', 'Service Test', 'Jumpering wire', 
             'Quality check', 'Migration Ready', 'Migration', 'TE PAC Completion']











