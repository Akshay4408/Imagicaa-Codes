#!/usr/bin/env python
# coding: utf-8

# In[34]:



import pandas as pd
import numpy as np


# In[35]:


Export = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\Nov.xlsx")
Export.head()


# In[36]:


Export.dropna(subset=['status_name', 'phone_number_dialed'], how='all', inplace=True)
#df = df[df['EPS'].notna()]


# In[37]:


Export.head()


# In[38]:


DL = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\Disposition List.xlsx")
DL.head()


# In[39]:


Report = pd.merge(Export, DL, how='left',
        left_on='status_name', right_on='status_name')


# In[40]:



Report = Report.sort_values("Num", ascending=True)

Report.head()



# In[41]:


Report = Report.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[42]:


connectivity =  pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\connectivity.xlsx")
connectivity.head()


# In[43]:


Report = pd.merge(Report, connectivity, how='left',
        left_on='status_name', right_on='status_name')
Report.head()


# In[44]:


Report["status_name"].replace({'Lead Being Called': 'No Answer', 'Busy Auto': 'Busy'}, inplace=True)
Report.head()


# In[45]:


connect = Report.loc[Report['Dispositions'] == "Connect"]


Notconnect = Report.loc[Report['Dispositions'] == "Not Connect"]


# In[46]:


StatusCount = pd.pivot_table(data = connect, index = ['status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )


StatusCount['Avg%'] = StatusCount['phone_number_dialed'] / 539
StatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount['Avg%']], index = StatusCount.index)

StatusCount = StatusCount.rename(columns = {'phone_number_dialed' : 'FTD', 'Avg%': 'FTD Avg%'})



StatusCount


# In[47]:


StatusCount1 = pd.pivot_table(data = Notconnect, index = ['status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )


StatusCount1['Avg%'] = StatusCount1['phone_number_dialed'] / 303
StatusCount1['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in StatusCount1['Avg%']], index = StatusCount1.index)

StatusCount1 = StatusCount1.rename(columns = {'phone_number_dialed' : 'FTD', 'Avg%': 'FTD Avg%'})



StatusCount1


# In[48]:


Dispo = pd.pivot_table(data = Report, index = ['Dispositions'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

Dispo['Avg%'] = Dispo['status_name'] / 842
Dispo['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in Dispo['Avg%']], index = Dispo.index)

Dispo = Dispo.rename(columns = {'status_name' : 'FTD', 'Avg%': 'FTD Avg%'})
Dispo


# In[49]:


Report.head()


# In[50]:


Dump = pd.read_excel(r"C:\Users\3363\Desktop\Imagica\EOD\Export.xlsx")
Dump.head()


# In[51]:


Dump.dropna(subset=['status_name', 'phone_number_dialed'], how='all', inplace=True)
#df = df[df['EPS'].notna()]


# In[52]:


MTDReport = pd.merge(Dump, DL, how='left',
        left_on='status_name', right_on='status_name')

MTDReport.head()


# In[53]:



MTDReport = MTDReport.sort_values("Num", ascending=True)

MTDReport.head()



# In[54]:


MTDReport = MTDReport.drop_duplicates(
  subset = ['phone_number_dialed'],
  keep = 'first').reset_index(drop = True)


# In[55]:


MTDReport = pd.merge(MTDReport, connectivity, how='left',
        left_on='status_name', right_on='status_name')
MTDReport.head()


# In[56]:


MTDReport["status_name"].replace({'Lead Being Called': 'No Answer', 'Busy Auto': 'Busy'}, inplace=True)
MTDReport.head()


# In[57]:


MTDconnect = MTDReport.loc[MTDReport['Dispositions'] == "Connect"]


MTDNotconnect = MTDReport.loc[MTDReport['Dispositions'] == "Not Connect"]


# In[58]:


MTDDispo = pd.pivot_table(data = MTDReport, index = ['Dispositions'], values = 'status_name', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )

MTDDispo['Avg%'] = MTDDispo['status_name'] / 10899
MTDDispo['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in MTDDispo['Avg%']], index = MTDDispo.index)

MTDDispo = MTDDispo.rename(columns = {'status_name' : 'MTD', 'Avg%': 'MTD Avg%'})
MTDDispo


# In[59]:


MTDStatusCount = pd.pivot_table(data = MTDconnect, index = ['status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )


MTDStatusCount['Avg%'] = MTDStatusCount['phone_number_dialed'] / 4073
MTDStatusCount['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in MTDStatusCount['Avg%']], index = MTDStatusCount.index)

MTDStatusCount = MTDStatusCount.rename(columns = {'phone_number_dialed' : 'MTD', 'Avg%': 'MTD Avg%'})



MTDStatusCount


# In[60]:


MTDStatusCount1 = pd.pivot_table(data = MTDNotconnect, index = ['status_name'], values = 'phone_number_dialed', aggfunc = np.size, fill_value= 0,margins = True, margins_name = 'Total' )


MTDStatusCount1['Avg%'] = MTDStatusCount1['phone_number_dialed'] / 6826
MTDStatusCount1['Avg%'] = pd.Series(["{0:.2f}%".format(val * 100) for val in MTDStatusCount1['Avg%']], index = MTDStatusCount1.index)

MTDStatusCount1 = MTDStatusCount1.rename(columns = {'phone_number_dialed' : 'MTD', 'Avg%': 'MTD Avg%'})



MTDStatusCount1


# In[61]:


ConnectivityReport =  pd.merge(MTDDispo, Dispo, how = 'outer', on = 'Dispositions')
ConnectivityReport


# In[62]:


ConnectedReport =  pd.merge( MTDStatusCount,StatusCount, how = 'outer', on = 'status_name')
ConnectedReport


# In[63]:


NotConnectedReport =  pd.merge( MTDStatusCount1,StatusCount1, how = 'outer', on = 'status_name')
NotConnectedReport


# In[64]:


NotConnectedReport.fillna(0, inplace = True)
ConnectedReport.fillna(0, inplace = True)


# In[65]:




writer = pd.ExcelWriter(r"C:\Users\3363\Desktop\Akshay\Imagica EOD.xlsx", engine='xlsxwriter')


# In[66]:


ConnectivityReport.to_excel(writer, sheet_name='Dispo')
ConnectedReport.to_excel(writer, sheet_name='Connect')
NotConnectedReport.to_excel(writer, sheet_name='Not Connect')
MTDReport.to_excel(writer, sheet_name='Dump')




writer.save()


# In[ ]:




