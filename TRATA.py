import xlsxwriter
import pandas as pd
import numpy as np
import openpyxl
import re

df_B05 = pd.read_excel('DataProcessing/BASE/B05.xlsx','SEARCH')
df_B05_AC = pd.read_excel('DataProcessing/BASE/B05.xlsx','AC_IM')

df_IM_PM_TUB = pd.read_excel('DataProcessing/BASE/IM_PM.xlsx','TUB')
df_IM_PM_NOMEN = pd.read_excel('DataProcessing/BASE/IM_PM.xlsx','NOMENCLATURE')
df_IM_PM_INSP = pd.read_excel('DataProcessing/BASE/IM_PM.xlsx','INSP')

df_NCR_EXP_IM = pd.read_excel('DataProcessing/Export_NCR-CQ.xlsx','NCRs')
df_NCR_EXP_PM = pd.read_excel('DataProcessing/Export_NCR-CQ.xlsx','Produtos')

df_REGRESS = pd.read_excel('DataProcessing/Export_NCR-REGRESS.xlsx')
df_DQR = pd.read_excel('DataProcessing/SBR-2 Security MilestonesFollow - up2.XLS')
#-------------------------------------------------------------------------------------------------

df_NCR_DB = df_NCR_EXP_IM[['Número NCR','Status','Descrição', 'Deliberação N1', 'Deliberação N2', 'Deliberação N3']]
df_NCR_DB.fillna('-', inplace=True)

df_B05.fillna('-',inplace=True)
df_NCR_DB.fillna('-',inplace=True)
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
#-------------------------------------------------------------------------------------------------

prod = pd.merge(df_IM_PM_NOMEN, df_NCR_EXP_PM, left_on=['PM'], right_on=['Produto'], how='inner')
prod_T = pd.merge(df_IM_PM_NOMEN, prod, on=['PM'], how='right')
prod_T.rename(columns={'IM_x':'IM2','IM_y':'IM3'}, inplace=True)
prod_all = pd.merge(df_IM_PM_INSP, prod_T, on=['PM'], how='right')

df_NCR_PROD = prod_all[['PM','Número NCR','IM','IM2','IM3']].drop_duplicates()
df_NCR_PROD.fillna('-', inplace=True)
df_NCR_PROD.index = pd.Index(np.arange(0,len(df_NCR_PROD)))

for a in df_NCR_PROD.index:
    if df_NCR_PROD['IM'].loc[a] == '-':
        df_NCR_PROD['IM'].loc[a] = df_NCR_PROD['IM2'].loc[a]
    elif df_NCR_PROD['IM2'].loc[a] == '-':
        df_NCR_PROD['IM2'].loc[a] = df_NCR_PROD['IM3'].loc[a]

new_PM = pd.merge(df_NCR_PROD, df_B05, on=['PM'], how='inner')
new_PM = new_PM[new_PM['PM'] != '-'].drop_duplicates()

for a in df_NCR_PROD.index:
    for b in new_PM['PM']:
        if b in df_NCR_PROD['PM'].loc[a]:
            ncr.append([df_NCR_PROD['Número NCR'].loc[a], df_NCR_PROD['IM'].loc[a]])
#-------------------------------------------------------------------------------------------------

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

new_dqr = pd.merge(df_DQR, df_B05_AC, left_on=['Certificate'], right_on=['CERTIFICATE'], how='inner')
new_dqr = new_dqr.drop_duplicates()
#-------------------------------------------------------------------------------------------------

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
#-------------------------------------------------------------------------------------------------

df1 =  result[result['IM'] != '-'].groupby(['NCR','IM','STATUS_NCR','REGRESS','STATUS_REGRESS','STATUS_DQR']).count()
df2 =  result[result['IM'] != '-'].groupby(['IM','NCR','STATUS_NCR','REGRESS','STATUS_REGRESS','STATUS_DQR']).count()

writer = pd.ExcelWriter('AC_NCR_SEARCHED.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name='BY_NCR')
df2.to_excel(writer, sheet_name='BY_IM')
writer.save()

print('Mal feito, feito!')