#!/usr/bin/env python
# coding: utf-8

# In[1]:



import pandas as pd
import numpy as np


# In[2]:


rw = pd.read_excel(r"C:\Users\3363\Desktop\Reports\Care.xlsx")
rw.head()


# In[3]:


rw['Date'] = pd.to_datetime(rw['call_date']).dt.date
rw['Time'] = pd.to_datetime(rw['call_date']).dt.time

#rw.drop('call_date',
#  axis='columns', inplace=True)

rw.head()


# In[4]:


connectivity =  pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\connectivity.xlsx")
connectivity.head()


# In[5]:


Main = pd.merge(rw, connectivity, how='left',
        left_on='status_name', right_on='status_name')
Main.head()


# In[6]:


Main['datehour'] = Main['call_date'].dt.hour


# In[7]:


import datetime


# In[8]:


Main['Date'] = Main['call_date'].dt.strftime('%d %b')


# In[9]:


def test(x):
    if x == 7:
        return "7 AM - 8 AM"
    if x == 8:
        return "8 AM - 9 AM"
    if x == 9:
        return "9 AM - 10 AM"
    if x == 10:
        return "10 AM - 11 AM"
    if x == 11:
        return "11 AM - 12 PM"
    if x == 12:
        return "12 PM - 1 PM"
    if x == 13:
        return "1 PM - 2 PM"
    if x == 14:
        return "2 PM - 3 PM"
    if x == 15:
        return "3 PM - 4 PM"
    if x == 16:
        return "4 PM - 5 PM"
    if x == 17:
        return "5 PM - 6 PM"
    if x == 18:
        return "6 PM - 7 PM"
    if x == 19:
        return "7 PM - 8 PM"
    if x == 20:
        return "8 PM - 9 PM"
    else:
        return "Time exceeds"
    


# In[10]:


Main['Hours'] = Main['datehour'].apply(lambda x: test(x))


# In[11]:


Main = Main[Main.Hours != "Time exceeds"]


# In[12]:


Result =pd.pivot_table(data= Main, index = ['Hours'],columns= ['Dispositions'], values = 'call_date', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Result.head()


# In[13]:


Result1 =pd.pivot_table(data= Main, index = ['Hours'],columns= ['Dispositions'], values = 'length_in_sec', aggfunc = np.sum, fill_value=0, margins = True, margins_name= 'Total')
Result1.head()


# In[14]:


import operator

fmt = operator.methodcaller('strftime', '%H:%M:%S')
Result1['Total Handling Time'] = Result1['Connect'] / 60
Result1['AHT'] = pd.to_datetime(Result1['Total Handling Time'], unit='s').map(fmt)
Result1 = Result1[['AHT']]
Result1


# In[15]:


Time = pd.read_excel(r"C:\Users\3363\Desktop\Reports\Time table.xlsx")
Time.head()


# In[16]:


Report = pd.merge(Result, Time, how='left',
        left_on='Hours', right_on='Hours')


Report = Report.sort_values(by=['Num'])

Report.drop('Num',
  axis='columns', inplace=True)


Report


# In[17]:


Report.columns


# In[18]:


Report = Report.rename(columns={'Total': 'Total Inbound Calls', 'Connect': 'Answered Calls', 'Not Connect': 'Drop calls'})


# In[19]:


Report = Report[['Hours', 'Total Inbound Calls', 'Answered Calls', 'Drop calls']]
Report


# In[20]:


Result1


# In[21]:


Report = pd.merge(Report, Result1, how='left',
        left_on='Hours', right_on='Hours')

Report


# In[22]:


Report.to_excel(r"C:\Users\3363\Desktop\Akshay\Imagica inbound Care Report.xlsx", index= False)


# In[ ]:




