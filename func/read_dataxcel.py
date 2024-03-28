import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt

# df = pd.read_excel(r'D:\PM_PYTHON\Dakdrinh\DATA\DR_THUYVAN.xlsx',sheet_name='Dulieu2017-2020')
# df = df.iloc[1:,:11]
# df = df.drop('Unnamed: 1',axis=1)
# df.rename(columns={'Unnamed: 0':'time'},inplace=True)
# df.dropna(subset=['time'], inplace=True)
# df = df[['time','H(cm)','Q(m3/s)','Q tràn obs','Q xa cống  obs','Q cm thực']]
# df.rename(columns={'Q(m3/s)':'Qluve','Q tràn obs':'Qtranlu','Q xa cống  obs':'Qluxacong','Q cm thực':'Qlucm'},inplace=True)
# # print(df)

df= pd.read_excel(r'DATA\KIEMTRA1.xlsx')

dft= pd.read_excel(r'DATA\KIEMTRA2023.xlsx')
dft = dft.iloc[:,:6]
df =pd.concat([df,dft],axis=0)
df.to_excel('DATA/TONGHOP.xlsx',index=False)
print(df)

# #2021
# df1 = pd.read_excel(r'DATA\mualu2022.xlsm',sheet_name='Nam 2022')
# df1.columns = df1.loc[0]
# df1 = df1.iloc[1:500,:13]
# df1['Ngày giờ'] = pd.to_datetime(df1['Ngày giờ'])
# df1.rename(columns={'Ngày giờ':'time'},inplace=True)
# # plt.plot(df1['time'],df1['Qtđ'])
# # plt.show()

# # print(df1)
# df2 = pd.read_excel(r'DATA\mualu2022.xlsm',sheet_name='lu')
# df2.columns = df2.loc[1]
# df2 = df2.iloc[2:3000,:7]
# df2['time'] = pd.to_datetime(df2['time'])
# df2.rename(columns={'Q(m3/s)':'Qluve','Q tràn obs':'Qtranlu','Q xa cống  obs':'Qluxacong','Q cm thực':'Qlucm'},inplace=True)
# df2['time'] = pd.date_range(datetime(2022,9,1),periods=len(df2['time']),freq='H')
# # plt.plot(df2['time'],df2['H(cm)'])
# # plt.show()

# # print(df2)
# df_mer = df2.merge(df1,how='left',on='time')
# df_mer['H(cm)'].update(df_mer['Htđ(m)'])
# df_mer['Qluve'].update(df_mer['Qtđ'])
# df_mer['Qtranlu'].update(df_mer['Q xả tràn obs'])
# df_mer['Qluxacong'].update(df_mer['Q duytridc_obs'])
# df_mer['Qlucm'].update(df_mer['Qcm-obs'])
# # df_mer = df_mer[df_mer['Htđ(m)'].isnull()==False]
# print(df_mer)
# df_mer.to_excel('DATA/KIEMTRA.xlsx')
# # plt.plot(df_mer['time'],df_mer['H(cm)'])
# # plt.plot(df_mer['time'],df_mer['H(cm)'])
# # plt.plot(df_mer['time'],df_mer['Htđ(m)'])
# # plt.show()


# #2023
# df1 = pd.read_excel(r'DATA/DR_THUYVAN.xlsx',sheet_name='DRHN')
# df1.columns = df1.loc[0]
# df1 = df1.iloc[1:500,:13]
# df1['Ngày giờ'] = pd.to_datetime(df1['Ngày giờ'])
# df1.rename(columns={'Ngày giờ':'time'},inplace=True)
# # plt.plot(df1['time'],df1['Qtđ'])
# # plt.show()

# # print(df1)
# df2 = pd.read_excel(r'DATA/DR_THUYVAN.xlsx',sheet_name='LULU')
# df2.columns = df2.loc[1]
# df2 = df2.iloc[2:3000,:7]
# df2['time'] = pd.to_datetime(df2['time'])
# df2.rename(columns={'Q(m3/s)':'Qluve','Q tràn obs':'Qtranlu','Q xa cống  obs':'Qluxacong','Q cm thực':'Qlucm'},inplace=True)
# df2['time'] = pd.date_range(datetime(2023,9,1),periods=len(df2['time']),freq='H')
# # plt.plot(df2['time'],df2['H(cm)'])
# # plt.show()

# # print(df2)
# df_mer = df2.merge(df1,how='left',on='time')
# df_mer['H(cm)'].update(df_mer['Htđ(m)'])
# df_mer['Qluve'].update(df_mer['Qtđ'])
# df_mer['Qtranlu'].update(df_mer['Q xả tràn obs'])
# df_mer['Qluxacong'].update(df_mer['Q duytridc_obs'])
# df_mer['Qlucm'].update(df_mer['Qcm-obs'])
# # df_mer = df_mer[df_mer['Htđ(m)'].isnull()==False]
# print(df_mer)
# df_mer.to_excel('DATA/KIEMTRA2023.xlsx')
# # # plt.plot(df_mer['time'],df_mer['H(cm)'])
# # # plt.plot(df_mer['time'],df_mer['H(cm)'])
# # # plt.plot(df_mer['time'],df_mer['Htđ(m)'])
# # # plt.show()