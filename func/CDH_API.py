import pandas as pd
from datetime import datetime,timedelta
import numpy as np
from FUNC.Seach_file import tim_file,read_txt,vitridat
def CDH_API_mucnuoc():
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,7)
    bd = kt - timedelta(days=1)
    dfname = pd.read_csv('TS_ID/CDH/CDH_H_QNGA.txt',sep="\s+",header=None)
    data = pd.DataFrame()
    data['time'] = pd.date_range(bd,kt,freq='H')
    
    for tsid in zip(dfname[0],dfname[1]):
        url = 'https://cdh.vnmha.gov.vn/KiWIS/KiWIS?http://slportal.kttv.gov.vn/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&ts_id={}&from={}&to={}'
        df = pd.read_html(url.format(tsid[0],bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d')))
        for i in df:
            i = i.iloc[4:,:].dropna()
            i.rename(columns={0:'time',1:tsid[1]},inplace=True)
            i['time'] = pd.to_datetime(i['time'])
            i['time'] = i['time'].dt.strftime('%Y-%m-%d %H:%M:%S')
            i['time'] = pd.to_datetime(i['time'])
            data = data.merge(i,how='left',on='time')
    data.set_index('time',inplace=True)
    data =data.astype(float)
    id = vitridat() # tim vi tri dat
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    with pd.ExcelWriter(pth,mode='a',engine_kwargs={'keep_vba': True},engine='openpyxl',if_sheet_exists='overlay') as writer:   # ghi vao file co san
        data.to_excel(writer, sheet_name='H',startrow=id -1, startcol=3, header=False, index=False)

def CDH_API_muaoda():
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,7)
    bd = kt - timedelta(hours=25)
    dfname = pd.read_csv('TS_ID/CDH/CDH_MUA_ODA.txt',sep="\s+",header=None)
    dfname['loaitram'] = 'ODA'
    dfname = dfname.sort_values(by=[1])
    data = pd.DataFrame()
    data['time'] = pd.date_range(bd,kt,freq='H')
    for tsid in zip(dfname[0],dfname[1]):
        url = 'https://cdh.vnmha.gov.vn/KiWIS/KiWIS?http://slportal.kttv.gov.vn/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&ts_id={}&from={}&to={}'
        df = pd.read_html(url.format(tsid[0],bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d')))
        for i in df:
            i = i.iloc[4:,:].dropna()
            i.rename(columns={0:'time',1:tsid[1]},inplace=True)
            i['time'] = pd.to_datetime(i['time'])
            i['time'] = i['time'].dt.strftime('%Y-%m-%d %H:%M:%S')
            i['time'] = pd.to_datetime(i['time'])
            data = data.merge(i,how='left',on='time')
    data.set_index('time',inplace=True)
    data =data.astype(float)
    data.fillna(method='ffill', inplace=True) # thay the nhung gia tri trong bang nan
    data = data.diff()
    tgg = bd + timedelta(hours=1)
    data = data.loc[data.index >= tgg]
    id = vitridat() # tim vi tri dat
    print(data)
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    with pd.ExcelWriter(pth,mode='a',engine_kwargs={'keep_vba': True},engine='openpyxl',if_sheet_exists='overlay') as writer:   # ghi vao file co san
        data.to_excel(writer, sheet_name='Mua',startrow=id -1, startcol=0, header=False, index=True)
        
def CDH_API_muavrain():
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,7)
    bd = kt - timedelta(hours=25)
    dfname = pd.read_csv('TS_ID/CDH/CDH_MUAVRAIN.txt',sep="\s+",header=None)
    dfname = dfname.dropna()
    dfname = dfname.sort_values(by=[1])
    data = pd.DataFrame()
    data['time'] = pd.date_range(bd,kt,freq='H')
    for tsid in zip(dfname[0],dfname[1]):
        url = 'https://cdh.vnmha.gov.vn/KiWIS/KiWIS?http://slportal.kttv.gov.vn/KiWIS/KiWIS?service=kisters&type=queryServices&request=getTimeseriesValues&datasource=0&format=html&ts_id={}&from={}&to={}'
        df = pd.read_html(url.format(tsid[0],bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d')))
        for i in df:
            i = i.iloc[4:,:].dropna()
            i.rename(columns={0:'time',1:tsid[1]},inplace=True)
            i['time'] = pd.to_datetime(i['time'])
            i['time'] = i['time'].dt.strftime('%Y-%m-%d %H:%M:%S')
            i['time'] = pd.to_datetime(i['time'])
            data = data.merge(i,how='left',on='time')
    data.set_index('time',inplace=True)
    data =data.astype(float)
    tgg = bd + timedelta(hours=1)
    data = data.loc[data.index >= tgg]
    id = vitridat() # tim vi tri dat
    print(data)
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    with pd.ExcelWriter(pth,mode='a',engine_kwargs={'keep_vba': True},engine='openpyxl',if_sheet_exists='overlay') as writer:   # ghi vao file co san
        data.to_excel(writer, sheet_name='Mua',startrow=id -1, startcol=0, header=False, index=True)


