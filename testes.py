
df_B05 = pd.read_excel('DataProcessing/B05.xlsx','SEARCH')
df_B05_AC = pd.read_excel('DataProcessing/B05.xlsx','AC_IM')
df_NCR_DB = pd.read_excel('DataProcessing/DB_NCR_CQ.xlsx')
df_NCR_PROD = pd.read_excel('DataProcessing/DB_NCR_CQ.xlsx', 'PROD')
df_NCR_EXP = pd.read_excel('DataProcessing/Export_NCR-CQ.xlsx')
df_REGRESS = pd.read_excel('DataProcessing/Export_NCR-REGRESS.xlsx')
df_DQR = pd.read_excel('DataProcessing/SBR-2 Security MilestonesFollow - up2.XLS')




import xlsxwriter
import pandas as pd
import numpy as np
import openpyxl
import re

df_B05 = pd.read_excel('DataProcessing/B05.xlsx','SEARCH')
df_B05_AC = pd.read_excel('DataProcessing/B05.xlsx','AC_IM')
df_NCR_DB = pd.read_excel('DataProcessing/DB_NCR_CQ.xlsx')
df_NCR_PROD = pd.read_excel('DataProcessing/DB_NCR_CQ.xlsx', 'PROD')
df_NCR_EXP = pd.read_excel('DataProcessing/Export_NCR-CQ.xlsx')
df_REGRESS = pd.read_excel('DataProcessing/Export_NCR-REGRESS.xlsx')
df_DQR = pd.read_excel('DataProcessing/SBR-2 Security MilestonesFollow - up2.XLS')


df_B05.fillna('-',inplace=True)
df_NCR_DB.fillna('-',inplace=True)
df_NCR_PROD.fillna('-',inplace=True)
df_DQR.fillna('-',inplace=True)

ncr = []
for a in df_NCR_DB.index:
    for b in df_B05['SLEEV']:
        if b in df_NCR_DB['Descrição'].loc[a]:
            ncr.append([df_NCR_DB['Número NCR'].loc[a], b])
        if b in df_NCR_DB['Deliberação N1'].loc[a]:
           ncr.append([df_NCR_DB['Número NCR'].loc[a], b])
        if b in df_NCR_DB['Deliberação N2'].loc[a]:
           ncr.append([df_NCR_DB['Número NCR'].loc[a], b])
        if b in df_NCR_DB['Deliberação N3'].loc[a]:
           ncr.append([df_NCR_DB['Número NCR'].loc[a], b])

new_PM = pd.merge(df_NCR_PROD, df_B05, left_on=['Produto'], right_on=['PM'], how='inner')
new_PM = new_PM[new_PM['Produto'] != '-'].drop_duplicates()

for a in df_NCR_PROD.index:
    for b in new_PM['Produto']:
        if b in df_NCR_PROD['Produto'].loc[a]:
            ncr.append([df_NCR_PROD['Número NCR'].loc[a], df_NCR_PROD['IM'].loc[a]])

df_read = pd.DataFrame(data=ncr,columns=['NCR','IM'])
df_read = df_read.drop_duplicates()
df_read['STATUS_NCR'] = '-'

for a in df_NCR_DB.index:
    for b in df_read.index:
        if df_read['NCR'].loc[b] == df_NCR_DB['Número NCR'].loc[a]:
            df_read['STATUS_NCR'].loc[b] = df_NCR_DB['Status'].loc[a]

df_DQR = df_DQR[['Item','OriginalJx','Certificate','StatusDQR']]
df_DQR = df_DQR[df_DQR['OriginalJx'] == 'J06']
df_DQR = df_DQR[df_DQR['Certificate'].str.contains('B05')]

'''new_dqr = []
for a in df_DQR.index:
    for b in df_B05_AC.index:
        if df_DQR['Certificate'].loc[a] == df_B05_AC['CERTIFICATE'].loc[b]:
            #print(df_B05_AC['CERTIFICATE'].loc[b], df_DQR['StatusDQR'].loc[a])
            new_dqr.append([df_B05_AC['FUN_IM'].loc[b], df_DQR['StatusDQR'].loc[a]])

dqr = pd.DataFrame(data=new_dqr,columns=['IM','STATUS'])
dqr = dqr.drop_duplicates()'''

new_dqr = pd.merge(df_DQR, df_B05_AC, left_on=['Certificate'], right_on=['CERTIFICATE'], how='inner')
new_dqr = new_dqr.drop_duplicates()

read_end = pd.merge(df_read, new_dqr, left_on=['IM'], right_on=['FUN_IM'], how='left')
read_end.rename(columns={'STATUS':'STATUS_DQR'}, inplace=True)
read_end.fillna('-', inplace=True)
read_end = read_end[read_end['OriginalJx'] != '-']

df_REGRESS.fillna('-', inplace=True)
df_REGRESS = df_REGRESS[['Número NCR','NCR Original','Status']]
df_REGRESS = df_REGRESS[df_REGRESS['NCR Original'] != '-']

result = pd.merge(read_end, df_REGRESS, left_on=['NCR'], right_on=['NCR Original'], how='left')
result.fillna('-', inplace=True)
result.rename(columns={'Número NCR':'REGRESS','Status':'STATUS_REGRESS','StatusDQR':'STATUS_DQR'}, inplace=True)
result = result[['NCR','IM','STATUS_NCR','REGRESS','STATUS_REGRESS','STATUS_DQR']]

df1 =  result[result['IM'] != '-'].groupby(['NCR','IM','STATUS_NCR','REGRESS','STATUS_REGRESS','STATUS_DQR']).count()
df2 =  result[result['IM'] != '-'].groupby(['IM','NCR','STATUS_NCR','REGRESS','STATUS_REGRESS','STATUS_DQR']).count()

writer = pd.ExcelWriter('AC_NCR_SEARCHED.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='BY_NCR')
df2.to_excel(writer, sheet_name='BY_IM')
writer.save()

print('Mal feito, feito!')