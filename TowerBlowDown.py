#######===|=<< Extracting Softwater Info from Excel >>=|===######
##############===]=>> Robert P Armstrong <<=[===############
##############=== ===]~~> 2022.01.25 <~~[=== ===############

import numpy as np
from numpy import round
import pandas as pd
import sys
import os
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
import numpy as np
import pandas as pd
import sys
import os
import xlwings as xw
from datetime import date, timedelta
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
import xlsxwriter
import openpyxl
from pathlib import Path


# In[100]:


#os.chdir(r"V:\Facilities Operations\UTILITIES METER READINGS\Utilreadings\UTILITIES 2021")
#BookNames=["JULY2020MR","AUGUST2020MR","SEPTEMBER2020MR","OCTOBER2020MR","NOVEMBER2020MR","DECEMBER2020MR","JANUARY2021MR","FEBRUARY2021MR","MARCH2021MR","APRIL2021MR","MAY2021MR","JUNE2021MR"]
os.chdir(r"V:\Facilities Operations\UTILITIES METER READINGS\Utilreadings\UTILITIES 2020")
BookNames=["JULY2019MR","AUGUST2019MR","SEPTEMBER2019MR","OCTOBER2019MR","NOVEMBER2019MR","DECEMBER2019MR","JANUARY2020MR","FEBRUARY2020MR","MARCH2020MR","APRIL2020MR","MAY2020MR","JUNE2020MR"]
BookNames

n = 12
m = 12

print(BookNames)
rindex=["MRC/STRAMANIS-JAMES HALL","ELECTRICAL ENG","MATH-COMPUTER","CASTLEMAN HALL","POWER PLANT SMALL","CHEMISTRY","BERTELSMEYER HALL","MCNUTT HALL","CIVIL","HAVENER CENTER","TOOMEY HALL","GALE-BULLMAN GEOTHERM"]
mu= [[0] * m] * n
#mu=pd.DataFrame(mu)
bd= [[0] * (m)] * n
#bd=pd.DataFrame(bd)
diff= [[0] * (m)] * n
#diff=pd.DataFrame(diff)

j=0
for i in range (12):
    bookname=BookNames[i]
    bookname=bookname+".xlsx"
    print(bookname)
    meter_reading=pd.read_excel(bookname,sheet_name="BLOWDOWN",header=1)
    mu[j]=meter_reading.iloc[0:12,3]
    bd[j]=meter_reading.iloc[0:11,6]
    diff[j]=meter_reading.iloc[0:11,8]
    j+=1
    
mu=np.array(mu,dtype=float)
bd=np.array(bd,dtype=float)
diff=np.array(diff,dtype=float)

bd

bldgs=["MRC/STRAMANIS-JAMES HALL","ELECTRICAL ENG","MATH-COMPUTER","CASTLEMAN HALL","POWER PLANT SMALL","CHEMISTRY","BERTELSMEYER HALL","MCNUTT HALL","CIVIL","HAVENER CENTER","TOOMEY HALL","GALE-BULLMAN GEOTHERM"]
df_mu=pd.DataFrame(mu.T, columns=BookNames, index=bldgs)
bldgs=["MRC/STRAMANIS-JAMES HALL","ELECTRICAL ENG","MATH-COMPUTER","CASTLEMAN HALL","POWER PLANT SMALL","CHEMISTRY","BERTELSMEYER HALL","MCNUTT HALL","CIVIL","HAVENER CENTER","TOOMEY HALL"]
df_bd=pd.DataFrame(bd.T, columns=BookNames, index=bldgs)
df_diff=pd.DataFrame(diff.T, columns=BookNames, index=bldgs)

df_diff

df_mu[0,15]=df_mu.sum(axis=1)
df_bd[0,15]=df_bd.sum(axis=1)
df_diff[0,15]=df_diff.sum(axis=1)


output_path = Path(r"C:\Users\rpan92\Documents\GitHub\2022_Working_Directory\Tower Makeup")
with pd.ExcelWriter(output_path/"output2.xlsx") as writer:
    df_mu.to_excel(writer, sheet_name='Sheet_1')
    df_bd.to_excel(writer, sheet_name='Sheet_2')
    df_diff.to_excel(writer, sheet_name='Sheet_3')
    
    
#with pd.ExcelWriter(r"C:\Users\rpan92\Documents\GitHub\2022_Working_Directory\Tower Makeup\output2.xlsx") as writer:


# In[ ]:





# In[109]:


#os.chdir(r"V:\Facilities Operations\UTILITIES METER READINGS\Utilreadings\UTILITIES 2021")
#BookNames=["JULY2020MR","AUGUST2020MR","SEPTEMBER2020MR","OCTOBER2020MR","NOVEMBER2020MR","DECEMBER2020MR","JANUARY2021MR","FEBRUARY2021MR","MARCH2021MR","APRIL2021MR","MAY2021MR","JUNE2021MR"]
os.chdir(r"V:\Facilities Operations\UTILITIES METER READINGS\Utilreadings\UTILITIES 2020")
BookNames=["JULY2019MR","AUGUST2019MR","SEPTEMBER2019MR","OCTOBER2019MR","NOVEMBER2019MR","DECEMBER2019MR","JANUARY2020MR","FEBRUARY2020MR","MARCH2020MR","APRIL2020MR","MAY2020MR","JUNE2020MR"]
BookNames

m=4
n=12

softwater=diff= [[0] * m] * n
#softwater=pd.DataFrame(softwater)


# In[110]:


j=0
for i in range (12):
    bookname=BookNames[i]
    bookname=bookname+".xlsx"
    print(bookname)
    soft_skid_reading=pd.read_excel(bookname,sheet_name="UTILITY WORKSHEET",header=None,index_col=0)
    softwater[j]=soft_skid_reading.loc[["MCNUTT GEO SOFT A","MCNUTT GEO SOFT B","MCNUTT GEO SOFT C","MCNUTT GEO SOFT D"],11]
    j+=1  
softwater=np.array(softwater,dtype=float)
df_softwater=pd.DataFrame(softwater.T, columns=BookNames)


# In[105]:





# In[111]:


os.chdir(r"C:\Users\rpan92\Documents\GitHub\2022_Working_Directory\Tower Makeup")
with pd.ExcelWriter('water.xlsx') as writer:  
    df_softwater.to_excel(writer, sheet_name='Sheet_2')


# In[112]:


os.chdir(r"V:\Facilities Operations\UTILITIES METER READINGS\Utilreadings\UTILITIES 2020")
BookNames=["JULY2019MR","AUGUST2019MR","SEPTEMBER2019MR","OCTOBER2019MR","NOVEMBER2019MR","DECEMBER2019MR","JANUARY2020MR","FEBRUARY2020MR","MARCH2020MR","APRIL2020MR","MAY2020MR","JUNE2020MR"]
BookNames

softwater=np.zeros([12,4])
softwater=pd.DataFrame(softwater)

soft_skid_reading=pd.read_excel(bookname,sheet_name="UTILITY WORKSHEET",header=None)

softwater=soft_skid_reading.iloc[[67,70,71,72],11]
print(softwater)


# In[113]:


softwater


# In[114]:


softwater


# In[ ]:




