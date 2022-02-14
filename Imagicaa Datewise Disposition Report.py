#!/usr/bin/env python
# coding: utf-8

# In[1]:



import pandas as pd
import numpy as np


# In[2]:


Dump = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\Dec\Inbound.xlsx")
Dump.head()


# In[3]:


Dump.dropna(subset=['status_name', 'phone_number_dialed'], how='all', inplace=True)
#df = df[df['EPS'].notna()]


# In[4]:


Dump['Date'] = pd.to_datetime(Dump['call_date']).dt.date
Dump['Time'] = pd.to_datetime(Dump['call_date']).dt.time

Dump.head()


# In[5]:


Dump['Date'] = pd.to_datetime(Dump['Date'], errors='coerce')

Dump['Date'] = Dump['Date'].dt.strftime('%d-%b')


# In[6]:


Dump["status_name"].replace({'Lead Being Called': 'No Answer', 'Busy Auto': 'Busy'}, inplace=True)
Dump.head()


# In[7]:


connectivity =  pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\connectivity.xlsx")
connectivity.head()


# In[8]:


Result = pd.merge(Dump, connectivity, how='left',
        left_on='status_name', right_on='status_name')
Result.head()


# In[9]:


DL = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\Disposition List.xlsx")
DL.head()


# In[10]:


Result = pd.merge(Result, DL, how='left',
        left_on='status_name', right_on='status_name')

Result.head()


# In[11]:



Result = Result.sort_values("Num", ascending=True)

Result.head()



# In[12]:


Result = Result.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)

Result.head()


# In[13]:


Connoncon =pd.pivot_table(data= Result, index = ['Dispositions'], columns = ['Date'],values = 'phone_number_dialed', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Connoncon


# In[14]:


connect = Result.loc[Result['Dispositions'] == "Connect"]


Notconnect = Result.loc[Result['Dispositions'] == "Not Connect"]


# In[15]:


Con =pd.pivot_table(data= connect, index = ['status_name'], columns = ['Date'],values = 'phone_number_dialed', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
Con


# In[16]:


NotCon =pd.pivot_table(data= Notconnect, index = ['status_name'], columns = ['Date'],values = 'phone_number_dialed', aggfunc = np.size, fill_value=0, margins = True, margins_name= 'Total')
NotCon


# In[17]:




writer = pd.ExcelWriter(r"C:\Users\3363\Desktop\Akshay\Imagica DatewiseEOD.xlsx", engine='xlsxwriter')


# In[18]:


Connoncon.to_excel(writer, sheet_name='Connevtivity')
Con.to_excel(writer, sheet_name='Connect')
NotCon.to_excel(writer, sheet_name='Not Connect')
Result.to_excel(writer, sheet_name='Dump')




writer.save()

