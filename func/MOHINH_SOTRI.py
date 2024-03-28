import os
import io
from datetime import datetime,timedelta
import pandas as pd
import numpy as np
import paramiko
import requests
import urllib.request
# Hàm để ánh xạ và đổi hướng
def convert_direction(x):
    if x >= 0:
        return x  # Giữ nguyên giá trị dương
    else:
        return x + 360  # Cộng 360 để chuyển sang giá trị dương

def map_wind_direction(x):
    return (
        'N' if 348.75 <= x <= 11.25 else
        'NNE' if 11.25 < x <= 33.75 else
        'NE' if 33.75 < x <= 56.25 else
        'ENE' if 56.25 < x <= 78.75 else
        'E' if 78.75 < x <= 101.25 else
        'ESE' if 101.25 < x <= 123.75 else
        'SE' if 123.75 < x <= 146.25 else
        'SSE' if 146.25 < x <= 168.75 else
        'S' if 168.75 < x <= 191.25 else
        'SSW' if 191.25 < x <= 213.75 else
        'SW' if 213.75 < x <= 236.25 else
        'WSW' if 236.25 < x <= 258.75 else
        'W' if 258.75 < x <= 281.25 else
        'WNW' if 281.25 < x <= 303.75 else
        'NW' if 303.75 < x <= 326.25 else
        'NNW'
    )

def read_ecmwf(ngay,tram,tg_db):
    found_path = False  
    for a in range(0,10):
        ngayt = ngay - timedelta(days=a)
        obs = ['12','00']
        for zz in obs:
            pth = r'\\admin-pc\DATA tin\MCMHMFSave\MCMHMFSave\bin\Debug\MCMHMF/' + ngayt.strftime('%Y%m') + '/' + ngayt.strftime('%Y%m%d')  + zz +'/'+ tram
            if os.path.exists(pth):
                print(pth)
                found_path =True
                break
        if found_path:
            break
    with open(pth,'r') as file:
        first_line = file.readline()
    file.close()
    file = open(pth,'r')
    cnten = file.read()
    cnten = cnten[cnten.index('FORECAST'):]
    file.close()
    # print(cnten)
    pp = first_line.split(" ")
    # for l,e in enumerate(pp):
    #     print(e)
    #     print(l)
    gio = int(pp[8].replace('\n',''))
    ngay_tt = int(pp[7].replace('\n',''))
    thang = int(pp[6].replace('\n',''))
    nam = int(pp[5].replace('\n',''))
    
    mohinh = pd.read_csv(io.StringIO(cnten), delimiter=r"\s+",error_bad_lines=False)
    mohinh = mohinh.astype(float)
   
    bd=datetime(nam,thang,ngay_tt,gio) + timedelta(hours=7)
    # print(bd)
    mohinh.insert(0,'time',pd.date_range(bd, periods=len(mohinh['FORECAST']), freq="6H"))
    mohinh['huong'] = np.degrees(np.arctan2(mohinh['USRF(m/s)'], mohinh['VSRF(m/s)']))
    mohinh['huong'] = mohinh['huong'].apply(convert_direction)
    mohinh['huong'] = mohinh['huong'].apply(lambda x: map_wind_direction(x))
    mohinh['wind_speed'] =round(np.sqrt(mohinh['USRF(m/s)']**2 + mohinh['VSRF(m/s)']**2),1)
    ngaytruoc = ngay + timedelta(hours=tg_db)
    mohinh = mohinh[(mohinh['time'] >= ngay) & (mohinh['time']<=ngaytruoc)]
    # print(mohinh)
    muadb = mohinh['RAIN6(mm/6h)'].sum()
    # sxm = mohinh['PoP(%)'].max()
    txmax = mohinh['TSRF(T)'].max()
    txmin = mohinh['TSRF(T)'].min()
    tnmax = mohinh['TTDSRF(T)'].max()
    tnmin = mohinh['TTDSRF(T)'].min()
    huonggio = mohinh['huong'].value_counts().idxmax()
    vmax = mohinh['wind_speed'].max()
    vmin = mohinh['wind_speed'].min()
    doamax = mohinh['RHSRF(%)'].max()
    doamin = mohinh['RHSRF(%)'].min()
    # print(ngay)
    # print(mohinh)
    return muadb,txmax,txmin,tnmax,tnmin,huonggio,vmax,vmin,doamax,doamin


def read_raii(ngay,tg_db):
    found_path = False  
    for m in range(0,10):
        ngayt  = ngay - timedelta(days=m) 
        pth = r'\\admin-pc\DATA tin\MCMHMFSave\MCMHMFSave\bin\Debug\RAII/' + ngayt.strftime('%Y%m') + '/' + ngayt.strftime('%Y%m%d')  +  '/QuangNgai.txt'
        if os.path.exists(pth):
            found_path = True  
            break
        if found_path:
            break
    
    with open(pth,'r') as file:
        first_line = file.readline()
    file.close()
    file = open(pth,'r')
    cnten = file.read()
    file.close()
    # print(cnten)
    pp = first_line.split(" ")
    gio = int(pp[-2].replace('\n',''))
    ngay_tt = int(pp[-3].replace('\n',''))
    thang = int(pp[-4].replace('\n',''))
    nam = int(pp[-5].replace('\n',''))
    
    cnten = cnten[cnten.index('FORECAST'):]
    mohinh = pd.read_csv(io.StringIO(cnten), delimiter=r"\s+",error_bad_lines=False)
    mohinh = mohinh.astype(float)
    
    bd=datetime(nam,thang,ngay_tt,gio) + timedelta(hours=7)
    # print(pd.date_range(bd, periods=40, freq="6H"))
    mohinh.insert(0,'time',pd.date_range(bd, periods=len(mohinh['FORECAST']), freq="6H"))
    # print(mohinh)
    mohinh['huong'] = np.degrees(np.arctan2(mohinh['USRF(m/s)'], mohinh['VSRF(m/s)']))
    mohinh['huong'] = mohinh['huong'].apply(convert_direction)
    mohinh['huong'] = mohinh['huong'].apply(lambda x: map_wind_direction(x))
    mohinh['wind_speed'] =round(np.sqrt(mohinh['USRF(m/s)']**2 + mohinh['VSRF(m/s)']**2),1)
    ngaytruoc = ngay + timedelta(hours=tg_db)
    mohinh = mohinh[(mohinh['time']>=ngay) & (mohinh['time']<=ngaytruoc)]
    # print(mohinh)
    muadb = mohinh['RAIN6(mm/6h)'].sum()
    # sxm = mohinh['PoP(%)'].max()
    txmax = mohinh['TSRF(C)'].max()
    huonggio = mohinh['huong'].value_counts().idxmax()
    vmax = mohinh['wind_speed'].max()
    vmin = mohinh['wind_speed'].min()
    doamax = mohinh['RHSRF(%)'].max()
    doamin = mohinh['RHSRF(%)'].min()
    return muadb,txmax,huonggio,vmax,vmin,doamax,doamin


def read_muadb_sever(ngay,tg_db):
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
    
    for i in range(0,10):
        tg = (ngay - timedelta(days=i)).strftime('%d%m%Y')
        remote_file_path = '/home/disk2/KQ_WRF72h/' + tg + '/Hinh_12z_36h/QUANGNGAI_12z_3k36h.txt'
        try:
            # Kiểm tra tồn tại của tệp tin
            sftp = client.open_sftp()
            sftp.stat(remote_file_path)
            # Đọc nội dung của tệp tin từ máy chủ
            remote_file = sftp.open(remote_file_path)
            file_contents = remote_file.read().decode()
            # Đóng kết nối SSH
            remote_file.close()
            sftp.close()
            client.close()
            # Xử lý dữ liệu đọc được
            df = pd.read_csv(io.StringIO(file_contents), delimiter=r"\s+", error_bad_lines=False)
            break
        except FileNotFoundError:
            print(f"Tệp tin '{remote_file_path}' không tồn tại.")
        except Exception as e:
            print(f"Có lỗi xảy ra khi đọc tệp tin: {e}")
            
            
    # xu ly thoi gian bat dau cua mo hinh
    thoigian = df.columns[-1].split('-')
    tgbd_mh = datetime(int(thoigian[2]),int(thoigian[1]),int(thoigian[0]),int(df.columns[-3].replace('h',''))) + timedelta(hours=3)
    
    df = df[1:]
    df = df.T
    df.columns = df.iloc[0]
    df = df[1:]
    df =df[['AnChi','BaTo','TraBong','ChauO','GaVuc','SonTay','SonGiang','TraKhuc','QuangNgai','MinhLong','SongVe','TraCau','SonHa']]
    df.dropna(inplace=True)
    df = df.reset_index(drop=True)
    df.insert(0,'time',pd.date_range(tgbd_mh,periods=df.shape[1]-1,freq='3H')) # chen them cot thoi gian vao ket qua du bao mo hinh
    df = df[(df['time']>ngay) & (df['time']<= (ngay + timedelta(hours=tg_db))) ]
    # print(df)
    df.iloc[:,1:] = df.iloc[:,1:].astype(float)
    df =df.iloc[:,1:].sum()
    df = df.to_frame().T
    return df




# read_raii(datetime(2023,6,24),72)