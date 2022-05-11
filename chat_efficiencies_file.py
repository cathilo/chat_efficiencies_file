#!/usr/bin/env python
# coding: utf-8

# In[3]:


# E,I,M,N,R,S,T,U


# In[131]:


import pandas as pd
from pathlib import Path
import win32com.client as win32
win32c = win32.constants


# In[157]:


filepath = "C:/Users/Catherine Liu/Downloads/"
name = "chateff1.xlsx"
name1 = 'chateff_edited_df_.csv'
filename = "C:/Users/Catherine Liu/Downloads/chateff1.xlsx"
sheetname = 'XSELL Chat Efficiencies  - Apr ' # change to the name of the worksheet
updated_date = '4/25/2022'
result_date = '4/23/2022'


# In[136]:


# create Excel object
excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel can be visible or not
excel.Visible = True 
# open Excel Workbook   
wb = excel.Workbooks.Open(filename)


# In[137]:


# insert a new row by shifting the original row down
wb.Sheets(sheetname).Columns("E").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("J").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("O").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("Q").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("V").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("X").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("Z").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("AB").Insert(win32c.xlShiftDown)


# In[138]:


wb.Sheets(sheetname).Columns("A").Insert(win32c.xlShiftDown)
wb.Sheets(sheetname).Columns("A").Insert(win32c.xlShiftDown)


# In[139]:





# In[140]:



wb.Sheets(sheetname).Range("B3").Value = result_date


# In[141]:


lastRow = wb.Sheets(sheetname).UsedRange.Rows.Count 


# In[142]:


#wb.Sheets(sheetname2).Range("B4":lastRow).Copy(Destination = wb.Sheets(sheetname).Range("B3"))
print(lastRow)


# In[143]:


range_ = "B4:B"+str(lastRow)
wb.Sheets(sheetname).Range('B3').Copy(Destination = wb.Sheets(sheetname).Range(range_))


# In[144]:


wb.Sheets(sheetname).Range("A3").Value = updated_date
range_A = "A4:A"+str(lastRow)
wb.Sheets(sheetname).Range('A3').Copy(Destination = wb.Sheets(sheetname).Range(range_A))


# In[145]:


range_G = "G3:G"+str(lastRow)
range_L = "L3:L"+str(lastRow)
range_Q = "Q3:Q"+str(lastRow)
range_S = "S3:S"+str(lastRow)
range_X = "X3:X"+str(lastRow)
range_Z = "Z3:Z"+str(lastRow)
range_AB = "AB3:AB"+str(lastRow)
range_AD = "AD3:AD"+str(lastRow)


wb.Sheets(sheetname).Range('H1').Copy(Destination = wb.Sheets(sheetname).Range(range_G))
wb.Sheets(sheetname).Range('M1').Copy(Destination = wb.Sheets(sheetname).Range(range_L))
wb.Sheets(sheetname).Range('R1').Copy(Destination = wb.Sheets(sheetname).Range(range_Q))
wb.Sheets(sheetname).Range('T1').Copy(Destination = wb.Sheets(sheetname).Range(range_S))
wb.Sheets(sheetname).Range('Y1').Copy(Destination = wb.Sheets(sheetname).Range(range_X))
wb.Sheets(sheetname).Range('AA1').Copy(Destination = wb.Sheets(sheetname).Range(range_Z))
wb.Sheets(sheetname).Range('AC1').Copy(Destination = wb.Sheets(sheetname).Range(range_AB))
wb.Sheets(sheetname).Range('AE1').Copy(Destination = wb.Sheets(sheetname).Range(range_AD))


# In[146]:


wb.Sheets(sheetname).Rows(1).EntireRow.Delete()


# In[147]:


group_1 = 'SMB WRLS Sales- TTEC- Merritt'
group_2 = 'SMB WRLS Sales- TTEC- Odedra'


# In[ ]:





# In[148]:



wb.Close(SaveChanges=1)
excel.Quit()


# In[149]:


df = pd.DataFrame(pd.read_excel(filename))
  


# In[150]:


df = df[df.GROUP!= 'Total']


# In[151]:





# In[152]:



df['GROUP'].fillna(method='ffill')


# In[ ]:





# In[153]:


df = df[df.AGENT!= 'Total']


# In[156]:


df.to_csv(filepath +name1 , encoding='utf-8', header =None)


# In[ ]:




