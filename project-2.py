import pandas as pd

df = pd.read_excel('sample1.xlsx', sheet_name='2020-12', skiprows=8, header=None)

df1 = df[[0, 2, 3, 5, 6, 8]].dropna().astype('int').astype('str')

j = '2023'
k = '2023'

for i in [0, 2, 3, 5]:
    j += ',' + df1[i]
    if i > 2: i = i + 3
    k += ',' + df1[i]

df2 = pd.DataFrame()
df2['開始時間'] = j
df2['終了時間'] = k

for x in ['開始時間', '終了時間']:
    df2[x] = pd.to_datetime(df2[x], format='%Y,%m,%d,%H,%M')

df2['使用時間'] = df2['終了時間'] - df2['開始時間']
df2['使用時間（分）'] = df2['使用時間'].map(lambda x: x.total_seconds() / 60.0)

df2.loc[df2['使用時間（分）'] < 0] = df2[['使用時間（分）']] + 1440

df3 = df.loc[:, [12, 13, 15, 23]]
df3 = df3.rename(columns={12: '使用目的', 13: '品種', 15: '解析数', 23: '所属'})

df4 = pd.merge(df2[['使用時間（分）']], df3, left_index=True, right_index=True)
df5 = df4[~(df4['所属'] == '信技')]

product = [['T6XH8', 'T6XM9', 'T6XN5', 'T6XN6', 'T6XP6', 'T6XW5'],
           ['T6XN9', 'T6XY9', 'T6XZ0', 'T6XZ5'],
           ['T6YU6']]

df6 = df4[~df4['使用目的'].isin(['教育', 'その他'])]
df7 = df5[~df5['使用目的'].isin(['教育', 'その他'])]

df8 = df4[df4['使用目的'].isin(['教育', 'その他'])]
df9 = df5[df5['使用目的'].isin(['教育', 'その他'])]

df10 = df6[~df6['品種'].isin(sum(product, []))]
df11 = df6[df6['品種'].isin(product[0])]
df12 = df6[df6['品種'].isin(product[1])]
df13 = df6[df6['品種'].isin(product[2])]

df14 = df7[~df7['品種'].isin(sum(product, []))]
df15 = df7[df7['品種'].isin(product[0])]
df16 = df7[df7['品種'].isin(product[1])]
df17 = df7[df7['品種'].isin(product[2])]

df18 = pd.DataFrame()
df19 = pd.DataFrame()

df18['BiCS4.5以前'] = df10[['使用時間（分）', '解析数']].sum()
df18['BiCS5'] = df11[['使用時間（分）', '解析数']].sum()
df18['BiCS6'] = df12[['使用時間（分）', '解析数']].sum()
df18['BiCS8'] = df13[['使用時間（分）', '解析数']].sum()
df18['その他'] = df8[['使用時間（分）', '解析数']].sum()

df19['BiCS4.5以前'] = df14[['使用時間（分）', '解析数']].sum()
df19['BiCS5'] = df15[['使用時間（分）', '解析数']].sum()
df19['BiCS6'] = df16[['使用時間（分）', '解析数']].sum()
df19['BiCS8'] = df17[['使用時間（分）', '解析数']].sum()
df19['その他'] = df9[['使用時間（分）', '解析数']].sum()

df20 = df4[['所属', '使用時間（分）', '解析数']].groupby('所属').sum().reindex(['解析2', '信技', 'PE技', '品証'], axis=0)

df21 = df5[['所属', '使用時間（分）', '解析数']].groupby('所属').sum().reindex(['解析2', 'PE技', '品証'], axis=0)

df22 = pd.concat([df20, df18.T])
df23 = pd.concat([df21, df19.T])

df22['使用時間（分）'] = df22['使用時間（分）'] / 60.0
df23['使用時間（分）'] = df23['使用時間（分）'] / 60.0

df22 = df22.rename(columns={'使用時間（分）': '時間（h）'})
df23 = df23.rename(columns={'使用時間（分）': '時間（h）'})

writer = pd.ExcelWriter('out1.xlsx', engine='xlsxwriter')

df22.to_excel(writer, sheet_name='LEあり', index=True, header=True)
df23.to_excel(writer, sheet_name='LEなし', index=True, header=True)

writer.save()
writer.close()
