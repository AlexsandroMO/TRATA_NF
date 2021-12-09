import pandas as pd
import openpyxl

df_ncr_base = pd.read_excel('DATA_SBR3/BASE_NCR-CQ.xlsx')
df_ncr_read = pd.read_excel('DATA_SBR3/Export_NCR-CQ.xlsx')

df_dqr_read = pd.read_excel('DATA_SBR3/SBR-3 Security MilestonesFollow - up.XLS')
df_dqr_base = pd.read_excel('DATA_SBR3/SBR3_MilestonesFollow.xlsx')

df_ncr_read = df_ncr_read[['Número NCR','Status','Tempo de Tramitação','Tempo de Tramitação Mais Recente','Descrição','Deliberação N1', 'Deliberação N2', 'Deliberação N3']]
df_ncr_read.fillna('-', inplace=True)

df_ncr_new = pd.merge(df_ncr_base, df_ncr_read, on=['Número NCR'], how='inner')
df_ncr_new = df_ncr_new.drop('Unnamed: 0', axis=1)
df_ncr_new['CHANGE'] = False
df_ncr_new['OLD'] = ''

for a in df_ncr_new.index:
    if df_ncr_new['Status_x'].loc[a] != df_ncr_new['Status_y'].loc[a]:
        df_ncr_new['OLD'].loc[a] = df_ncr_new['Status_x'].loc[a]
        df_ncr_new['Status_x'].loc[a] = df_ncr_new['Status_y'].loc[a]
        df_ncr_new['CHANGE'].loc[a] = True

df_dqr_read = df_dqr_read[['Item','OriginalJx','ActualJx','NºandDescription','Bigram','Certificate','StatusDQR']]
df_dqr_read.fillna('-')
df_dqr_read['ID'] = ''
df_dqr_base['ID'] = ''

for a in df_dqr_read.index:
    df_dqr_read['ID'].loc[a] = '{}{}{}'.format(df_dqr_read['Item'].loc[a],df_dqr_read['ActualJx'].loc[a],df_dqr_read['Certificate'].loc[a])

for a in df_dqr_base.index:
    df_dqr_base['ID'].loc[a] = '{}{}{}'.format(df_dqr_base['Item'].loc[a],df_dqr_base['ActualJx'].loc[a],df_dqr_base['Certificate'].loc[a])

df_dqr_base.rename(columns={'Item_x':'Item','OriginalJx_x':'OriginalJx','ActualJx_x':'ActualJx','NºandDescription_x':'NºandDescription','Bigram_x':'Bigram','Certificate_x':'Certificate','StatusDQR_x':'StatusDQR','StatusDQR_x':'StatusDQR'}, inplace=True)

df_dqr_base = df_dqr_base.drop('Unnamed: 0', axis=1)
#df_dqr_base = df_dqr_base.drop('StatusDQR_y', axis=1)

df_ncr_base.rename(columns={'Status_x':'Status', 'Tempo de Tramitação_x':'Tempo de Tramitação',
       'Tempo de Tramitação Mais Recente_x':'Tempo de Tramitação Mais Recente', 'Descrição_x':'Descrição', 'Deliberação N1_x':'Deliberação N1',
       'Deliberação N2_x':'Deliberação N2', 'Deliberação N3_x':'Deliberação N3'}, inplace=True)

df_ncr_base = df_ncr_base.drop('Unnamed: 0', axis=1)
#df_ncr_base = df_ncr_base.drop('Status_y', axis=1)

#--------------------------------------------------------------------
df_dqr_new = pd.merge(df_dqr_base, df_dqr_read, on=['ID'], how='inner')
#df_dqr_new = df_dqr_new.drop('Unnamed: 0', axis=1)
df_dqr_new = df_dqr_new[['Item_x','OriginalJx_x','ActualJx_x','NºandDescription_x','Bigram_x','Certificate_x','StatusDQR_x','StatusDQR_y']]
df_dqr_new.fillna('-', inplace=True)
df_dqr_new['CHANGE'] = False
df_dqr_new['OLD'] = ''

for a in df_dqr_new.index:
    if df_dqr_new['StatusDQR_x'].loc[a] != df_dqr_new['StatusDQR_y'].loc[a]:
        df_dqr_new['OLD'].loc[a] = df_dqr_new['StatusDQR_x'].loc[a]
        df_dqr_new['StatusDQR_x'].loc[a] = df_dqr_new['StatusDQR_y'].loc[a]
        df_dqr_new['CHANGE'].loc[a] = True

df_dqr_new.rename(columns={'Item_x':'Item','OriginalJx_x':'OriginalJx','ActualJx_x':'ActualJx','NºandDescription_x':'NºandDescription','Bigram_x':'Bigram','Certificate_x':'Certificate','StatusDQR_x':'StatusDQR','StatusDQR_x':'StatusDQR'}, inplace=True)

df_dqr_new = df_dqr_new.drop('StatusDQR_y', axis=1)
df_dqr_new.fillna('-', inplace=True)

#--------------------------------------------------------------------
df_ncr_new.rename(columns={'Status_x':'Status', 'Tempo de Tramitação_x':'Tempo de Tramitação',
       'Tempo de Tramitação Mais Recente_x':'Tempo de Tramitação Mais Recente', 'Descrição_x':'Descrição', 'Deliberação N1_x':'Deliberação N1',
       'Deliberação N2_x':'Deliberação N2', 'Deliberação N3_x':'Deliberação N3'}, inplace=True)

df_ncr_new = df_ncr_new.drop('Tempo de Tramitação_y', axis=1)
df_ncr_new = df_ncr_new.drop('Tempo de Tramitação Mais Recente_y', axis=1)
df_ncr_new = df_ncr_new.drop('Descrição_y', axis=1)
df_ncr_new = df_ncr_new.drop('Deliberação N1_y', axis=1)
df_ncr_new = df_ncr_new.drop('Deliberação N2_y', axis=1)
df_ncr_new = df_ncr_new.drop('Deliberação N3_y', axis=1)
df_ncr_new = df_ncr_new.drop('Status_y', axis=1)

df_ncr_new.fillna('-', inplace=True)

df_dqr_new.to_excel('DATA_SBR3/SBR3_MilestonesFollow.xlsx')
df_ncr_new.to_excel('DATA_SBR3/BASE_NCR-CQ.xlsx')

print('Mal feito, feito!')