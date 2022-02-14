#!/usr/bin/env python
# coding: utf-8

# In[80]:



import pandas as pd
import numpy as np


# In[81]:



apr1 = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\APR\one.xlsx")
apr1 = apr1.iloc[3:]
#apr1.apr1[:-1,:]
apr1 = apr1.rename(columns=apr1.iloc[0]).drop(apr1.index[0])
apr1 = apr1.iloc[:-1 , :]
apr1.head()


# In[82]:


apr2 = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\APR\two.xlsx")

apr2 = apr2.iloc[3:]
#apr2.iloc[:-1,:]

apr2 = apr2.rename(columns=apr2.iloc[0]).drop(apr2.index[0])

apr2 = apr2.iloc[:-1 , :]

apr2.head()


# In[83]:


Result = pd.merge(apr1, apr2, how = 'outer', on = 'ID')
Result.head()


# In[84]:


Result = Result[['ID', 'USER NAME_x', 'CALLS', 'TOTAL','TALK', 'WAIT', 'PAUSE_y', 'LB', 'SB', 'TB', 'WB']]
Result = Result.rename(columns={'USER NAME_x' : 'Advisor Name', 'PAUSE_y' : 'PAUSE', 'TOTAL': 'TOTAL LOGIN'})

Result1 = Result
Result1


# In[85]:



Result1['Thour'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.hour
Result1['Tminutes'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.minute
Result1['Tsecond'] = pd.to_datetime(Result1['TALK'], format='%H:%M:%S').dt.second


Result1['Thour'] = Result1['Thour'] * 3600
Result1['Tminutes'] = Result1['Tminutes'] * 60

Result1['Ttalk']  = Result1['Thour'] + Result1['Tminutes'] + Result1['Tsecond']


Result1['Whour'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.hour
Result1['Wminutes'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.minute
Result1['Wsecond'] = pd.to_datetime(Result1['WAIT'], format='%H:%M:%S').dt.second



Result1['Whour']= Result1['Whour'] * 3600
Result1['Wminutes'] = Result1['Wminutes'] * 60

Result1['Wwait']  = Result1['Whour'] +Result1['Wminutes'] + Result1['Wsecond']



Result1['Phour'] = pd.to_datetime(Result1['PAUSE'], format='%H:%M:%S').dt.hour
Result1['Pminutes'] = pd.to_datetime(Result1['PAUSE'], format='%H:%M:%S').dt.minute
Result1['Psecond'] = pd.to_datetime(Result1['PAUSE'], format='%H:%M:%S').dt.second



Result1['Phour']= Result1['Phour'] * 3600
Result1['Pminutes'] = Result1['Pminutes'] * 60

Result1['Ppause']  = Result1['Phour'] +Result1['Pminutes'] + Result1['Psecond']

Result1['Total Handling Time1'] = Result1['Ttalk'] + Result1['Wwait'] + Result1['Ppause']


import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')

Result1['Total Handling Time'] = pd.to_datetime(Result1['Total Handling Time1'], unit='s').map(fmt)

Result1 = Result1[['ID', 'Advisor Name', 'CALLS', 'TOTAL LOGIN','TALK', 'WAIT', 'PAUSE','LB', 'SB', 'TB', 'WB' ,'Total Handling Time1','Total Handling Time']]
Result1


# In[86]:


#Report = Report[['ID', 'USER NAME', 'Lead generated', 'Gross Revenues', 'Inbound', 'Outbound', 'Total Calls', 'TOTAL LOG IN TIME', 'TALK TIME', 'WAIT', 'PAUSE Time', 'DISPO', 'Idle Time', 'AB HVT Booking', 'HVT BOOKING TIME', 'DEAD TIME', 'BREAK TIME', 'LB- Lunch Break', 'SB- Snack Break', 'TB- Tea break', 'WB- Washroom break', 'ProductTime','TOTAL PROD. TIME', 'Utilization %']]
Result1['AHT1'] = Result1['Total Handling Time1'] / Result1['CALLS']

Result1["AHT1"].fillna(0, inplace = True)



import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')
Result1['AHT'] = pd.to_datetime(Result1['AHT1'] , unit='s').map(fmt)

Result1 = Result1.drop(columns=['Total Handling Time1', 'AHT1'])
Result1.head()


# In[87]:



Result1['Lhour'] = pd.to_datetime(Result1['LB'], format='%H:%M:%S').dt.hour
Result1['Lminutes'] = pd.to_datetime(Result1['LB'], format='%H:%M:%S').dt.minute
Result1['Lsecond'] = pd.to_datetime(Result1['LB'], format='%H:%M:%S').dt.second


Result1['Lhour'] = Result1['Lhour'] * 3600
Result1['Lminutes'] = Result1['Lminutes'] * 60

Result1['Llunch']  = Result1['Lhour'] + Result1['Lminutes'] + Result1['Lsecond']


Result1['Thour'] = pd.to_datetime(Result1['TB'], format='%H:%M:%S').dt.hour
Result1['Tminutes'] = pd.to_datetime(Result1['TB'], format='%H:%M:%S').dt.minute
Result1['Tsecond'] = pd.to_datetime(Result1['TB'], format='%H:%M:%S').dt.second



Result1['Thour']= Result1['Thour'] * 3600
Result1['Tminutes'] = Result1['Tminutes'] * 60

Result1['Ttea']  = Result1['Thour'] +Result1['Tminutes'] + Result1['Tsecond']



Result1['Whour'] = pd.to_datetime(Result1['WB'], format='%H:%M:%S').dt.hour
Result1['Wminutes'] = pd.to_datetime(Result1['WB'], format='%H:%M:%S').dt.minute
Result1['Wsecond'] = pd.to_datetime(Result1['WB'], format='%H:%M:%S').dt.second



Result1['Whour']= Result1['Whour'] * 3600
Result1['Wminutes'] = Result1['Wminutes'] * 60

Result1['Twb']  = Result1['Whour'] +Result1['Wminutes'] + Result1['Wsecond']

#Result1['BREAK TIME1'] = Result1['Llunch'] + Result1['Ttea'] + Result1['Twb']



Result1['Shour'] = pd.to_datetime(Result1['SB'], format='%H:%M:%S').dt.hour
Result1['Sminutes'] = pd.to_datetime(Result1['SB'], format='%H:%M:%S').dt.minute
Result1['Ssecond'] = pd.to_datetime(Result1['SB'], format='%H:%M:%S').dt.second

Result1['Shour']= Result1['Shour'] * 3600
Result1['Sminutes'] = Result1['Sminutes'] * 60

Result1['Ssb']  = Result1['Shour'] +Result1['Sminutes'] + Result1['Ssecond']

Result1['BREAK TIME1'] = Result1['Llunch'] + Result1['Ttea'] + Result1['Twb'] + Result1['Ssb']

import operator
fmt = operator.methodcaller('strftime', '%H:%M:%S')

Result1['BREAK TIME'] = pd.to_datetime(Result1['BREAK TIME1'], unit='s').map(fmt)

Result1 = Result1[['ID', 'Advisor Name', 'CALLS', 'TOTAL LOGIN', 'TALK', 'WAIT', 'BREAK TIME', 'Total Handling Time', 'AHT']]
Result1.head()


# In[88]:


Result1.to_excel(r"C:\Users\3363\Desktop\Akshay\Imagica APR Report.xlsx", index= False)


# In[ ]:




