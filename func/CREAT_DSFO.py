import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
import numpy as np
import pandas as pd
from mikeio import Dfs0, Dataset
from mikeio.eum import ItemInfo, EUMType, EUMUnit
from datetime import datetime, timedelta
from mikecore.DfsFile import DataValueType
import paramiko
from func.Seach_file import read_txt
import pyodbc
import io
from tkinter import messagebox
def docfile(path):
    file = open( path, "r")
    content = file.read()
    file.close()
    return content

def tentram(df,ten):
    df=df[df[0]==int(ten)]
    if df.empty:
        tents = 'no_name'
    else:
        tents= df[1].values[0]
    return tents

def read_muadb_sever_hochua(tg_db,tenho):
    hostname = '203.209.181.171'
    port = 22
    username = 'mpi'
    password = 'mpi@1234'
    # Tạo kết nối SSH
    client = paramiko.SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(hostname, port, username, password)
    # Đường dẫn tới tệp tin .txt trên máy chủ
    tg = datetime.now()
    # tg = (ngay - timedelta(days=1)).strftime('%d%m%Y')
    obs = ['12', '00']
    ngaylayfile = [tg,tg-timedelta(days=1)]
    thoat = False
    for ngay in ngaylayfile:
        if thoat:
            break
        for zz in obs:        
            try:
                if tenho !='A VƯƠNG':
                    remote_file_path = '/home/disk2/KQ_WRF72h/' + ngay.strftime('%d%m%Y') + '/Hinh_{}z_36h/QuangNam_{}z_3k36h.txt'.format(zz,zz)
                else:
                    remote_file_path = '/home/disk2/KQ_WRF72h/' + ngay.strftime('%d%m%Y') + '/{}z/TTB_{}z_36h.txt'.format(zz,zz)
                    thoigianbatdau = datetime(ngay.year,ngay.month,ngay.day,int(zz))
                    # print(thoigianbatdau)
                # print(remote_file_path)
                # Đọc nội dung của tệp tin từ máy chủ
                sftp = client.open_sftp()
                remote_file = sftp.open(remote_file_path)
                file_contents = remote_file.read().decode()
                thoat= True
                break
            except:
                pass
    
    # print(file_contents)
    # Đóng kết nối SSH
    remote_file.close()
    sftp.close()
    client.close()
    df = pd.read_csv(io.StringIO(file_contents), delimiter=r"\s+")
    if tenho !='A VƯƠNG':
        df = df[1:]
        df = df.T
        df.columns = df.iloc[0]
        df = df[1:]
    else:
        # df = df[1:]
        df = df.T
        df.columns = df.iloc[0]
        df = df[1:]
        # print(df)
        # print(df.T)
    if tenho =='SÔNG TRANH 2':
        df =df[['TraBui(ST2)','TraCang(ST2)','TraDon(ST2)','TraGiac(ST2)','TraLeng(ST2)','TraLinh(ST2)','TraMai(ST2)','UBNDNTM(ST2)','Dap(ST2)','TraNam2(ST2)','TraVan(ST2)']]
        # df.columns = ['Trà Bui','Trà Cang','Trà Dơn','Trà Giác','Trà Leng','Trà Linh','Trà Mai','UBNDNTM','Đập chính','Trà Nam','Trà Vân']
    elif tenho =='A VƯƠNG':
        # print(df)
        df = df.reset_index(False)
        # df =df[['TraBui(ST2)','TraCang(ST2)','TraDon(ST2)','TraGiac(ST2)','TraLeng(ST2)','TraLinh(ST2)','TraMai(ST2)','UBNDNTM(ST2)','Dap(ST2)','TraNam2(ST2)','TraVan(ST2)']]
        df = df[['AV1','AV2','AV3','AV4','AV5','AV6','HIEN']]
        # name_viet = ['Đập tràn A Vương','UBND Ab Vương','Đồn biên phòng A Nông','UBND Huyện Tây Giang','UBND Xã Dang','Trạm Xã A Tep','Trạm Xã A Rooi','Trạm UBND Xã Blahee']
        # df.columns = ['Đập tràn A Vương','UBND Ab Vương','Đồn biên phòng A Nông','UBND Huyện Tây Giang','UBND Xã Dang','Trạm Xã A Tep','Hien']
    elif tenho =='SÔNG BUNG 2':
        df =df[['DapSBung2','TrHySBung2','NMSongBung2','GaRiSBung2']]
        # df.columns = ['Đập SB2','TrHy','Chơm','A Xan']
    elif tenho =='SÔNG BUNG 4':
        df =df[['DonBQNGiang','ChaVaNMDHSB4','DapSBung4','ZuoihSBung4','TrHySBung2','LaDeeSBung4','CuakhauNG']]
        # df.columns = ['ĐăkPring','Chalval','Đầu mối','Zuôi','TrHy','LaDee','Đak Ốc']
    
    if tenho !='A VƯƠNG':
        df.dropna(inplace=True)
        df = df.reset_index(drop=True)
        first_line = file_contents[:file_contents.index('Time')-1]
        pp = first_line.split(" ")
        ngaythangnam = pp[-1].replace('\n','')
        gio = pp[-3].replace('h','')
        tgbd  = datetime.strptime(ngaythangnam + " " + gio, "%d-%m-%Y %H") + timedelta(hours=3)
        df.insert(0,'time',pd.date_range(tgbd,periods=(df.shape[0]),freq='3h'))
    else:
        # print(thoigianbatdau)
        df.insert(0,'time',pd.date_range(thoigianbatdau+ timedelta(hours=10),periods=(df.shape[0]),freq='3h')) 
    # print(df)
    df = df[(df['time']>= datetime.now()) & (df['time']<= datetime.now() + timedelta(hours=tg_db))]
    df.iloc[:,1:] = df.iloc[:,1:].astype(float)
    # print(df)
    # df =df.iloc[:,1:].sum()
    # df = df.to_frame().T
    return df


def TTB_API_mua_hochua():
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,now.hour)
    bd = kt - timedelta(days=10)
    data = pd.DataFrame()
    data['time'] = pd.date_range(bd,kt,freq='min')
    tram = pd.read_csv('ts_id/muahochua.txt')
    # print(tram)
    for item in zip(tram.Matram,tram.tentram,tram.TAB):
    # print(item[0],item[2],item[1])
        pth = 'http://113.160.225.84:2018/API_TTB/XEM/solieu.php?matram={}&ten_table={}&sophut=1&tinhtong=0&thoigianbd=%27{}%2000:00:00%27&thoigiankt=%27{}%2023:59:00%27'
        pth = pth.format(item[0],item[2],bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d'))
        df = pd.read_html(pth)
        df[0].rename(columns={"thoi gian":'time','so lieu':item[1]},inplace=True)
        df = df[0].drop('Ma tram',axis=1)
        df['time'] = pd.to_datetime(df['time'])
        data = data.merge(df,how='left',on='time')
    data.set_index('time',inplace=True)
    muagio = data.rolling(60,min_periods=1).sum()
    muagio = muagio[muagio.index.minute == 0]
    epsilon = 1e-10
    muagio = muagio.applymap(lambda x: 0 if abs(x) < epsilon else x)    
    muagio =muagio.astype(float)
    return muagio
def TTB_API_Q_hochua():
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,now.hour)
    bd = kt - timedelta(days=10)
    data = pd.DataFrame()
    data['time'] = pd.date_range(bd,kt + timedelta(days=2),freq='min')
    tram = pd.read_csv('ts_id/Qhochua.txt')
    # print(tram)
    for item in zip(tram.Matram,tram.tentram,tram.TAB):
    # print(item[0],item[2],item[1])
        pth = 'http://113.160.225.84:2018/API_TTB/XEM/solieu.php?matram={}&ten_table={}&sophut=1&tinhtong=0&thoigianbd=%27{}%2000:00:00%27&thoigiankt=%27{}%2023:59:00%27'
        pth = pth.format(item[0],item[2],bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d'))
        df = pd.read_html(pth)
        df[0].rename(columns={"thoi gian":'time','so lieu':item[1]},inplace=True)
        df = df[0].drop('Ma tram',axis=1)
        df['time'] = pd.to_datetime(df['time'])
        data = data.merge(df,how='left',on='time')
    data.set_index('time',inplace=True)
    epsilon = 1e-10
    data = data.applymap(lambda x: 0 if abs(x) < epsilon else x)    
    data =data.astype(float)
    data = data[data.index.minute == 0]
    return data



def muathucdo_dsf():
    # lấy số liệu mua thuc do 10 ngay
    try:
        mua.loc[11] # them vao de loi
        mua = TTB_API_mua_hochua()
        mua = mua.iloc[1:,:]
    except:
        # lay so lieu mua
        pth25 = read_txt('path_tin/DATA_EXCEL.txt') + '/DATA.accdb'
        FileName=(pth25)
        cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + FileName + ';')
        query = "SELECT * FROM mua"
        mua = pd.read_sql(query, cnxn)
        mua = mua.replace('-',0)
        mua = mua[mua['thoigian'] > (datetime.now() - timedelta(days=10))]
        # print(mua)
    #     mua.columns =['time','AV01', 'AV02', 'AV03', 'AV04', 'AV05', 'AV06', 'AV07', 'AV00', 'HIEN',
    #    'SB2DM', 'SB4DK', 'SB4CV', 'SB4DM', 'SB4ZU', 'SB4TR', 'SB2CH', 'SB2AX',
    #    'SB4LE', 'SB4DO']
        mua.rename(columns={'thoigian':'time'},inplace=True)
        mua.set_index('time',inplace= True)
        mua = mua.sort_index()
        mua = mua.astype(float)
    return mua

def creatdfso_R(df):    
    for a in df.columns:
        data =[df[a]]
        items=[(ItemInfo(a,EUMType.Rainfall, data_value_type=DataValueType.Instantaneous))]
        ds = Dataset(data, df.index, items)
        # Ghi thanh file dsfo
        dfs = Dfs0()
        dfs.write(filename='mike/' + a + ".dfs0", data=ds,title="Mua") # ten title

def creatdfso_Q(df):
    for a in df.columns:
        data =[df[a]]
        items=[(ItemInfo(a,EUMType.Discharge, data_value_type=DataValueType.Instantaneous))]
        ds = Dataset(data, df.index, items)
        # Ghi thanh file dsfo
        dfs = Dfs0()
        dfs.write(filename='mike/' + a + ".dfs0", data=ds,title="luuluong") # ten title
    # print(q)

def creat_input_mike():
    mua = muathucdo_dsf()
    mua = mua.reset_index(False)
    # print(mua.columns)
    now = datetime.now()
    bd = datetime(now.year,now.month,now.day,now.hour)
    data_mua =  pd.DataFrame()
    data_mua['time'] = pd.date_range(bd,freq='h',periods=30)
    st = read_muadb_sever_hochua(48,'SÔNG TRANH 2')
    # ['time', 'tracang', 'traleng', 'tranam2', 'trabui', 'tramai', 'tratap','dapsongtranh', 'tralinh', 'tragiac', 'tradon', 'travan', 'trabui2']
    # ['time', 'TraBui(ST2)', 'TraCang(ST2)', 'TraDon(ST2)', 'TraGiac(ST2)','TraLeng(ST2)', 'TraLinh(ST2)', 'TraMai(ST2)', 'UBNDNTM(ST2)','Dap(ST2)', 'TraNam2(ST2)', 'TraVan(ST2)']
    # print(st.columns)
    st.columns = ['time','trabui','tracang','tradon','tragiac','traleng','tralinh','tramai','trabui2','dapsongtranh','tranam2','travan']
    st['tratap'] = st['dapsongtranh']
    st =st[['time', 'tracang', 'traleng', 'tranam2', 'trabui', 'tramai', 'tratap','dapsongtranh', 'tralinh', 'tragiac', 'tradon', 'travan', 'trabui2']]
    # print(st)
    # print(mua)
    # print(data_mua)     
    mua = pd.concat([mua,st],axis=0)
    mua =mua.loc[~mua['time'].duplicated(keep='first')]
    # print(mua)
    mua.set_index('time',inplace=True)
    creatdfso_R(mua)
    # print(mua)
    q = TTB_API_Q_hochua()
    creatdfso_Q(q)
    # print(q)
    messagebox.showinfo('Thông báo','OK')
    