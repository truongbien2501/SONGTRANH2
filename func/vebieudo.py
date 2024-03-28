import array as arr
from ctypes.wintypes import SIZE
from datetime import datetime, timedelta
from turtle import st 
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure, legend, xlabel, ylabel
import matplotlib.dates as mdates
import pandas as pd
from tkinter import messagebox
from ftplib import FTP
# from func.Seach_file import tim_file,read_txt,vitridat
from scipy import interpolate
def upload_file(file_path, ftp_url, username, password):
    try:
        # Tách thành phần từ URL FTP
        url_parts = ftp_url.split("/")
        ftp_server = url_parts[2]
        # print(ftp_server)
        remote_path = "/".join(url_parts[3:]) + "/" + file_path.split('/')[-1]
        # print(remote_path)
        # Kết nối đến server FTP
        ftp = FTP(ftp_server)
        ftp.login(username, password)
        
        # Mở file cần tải lên
        with open(file_path, 'rb') as file:
            # Tải file lên FTP
            ftp.storbinary(f'STOR {remote_path}', file)
        print("Tải lên thành công!")
    except Exception as e:
        print("Lỗi khi tải lên file:", str(e))
    finally:
        # Đóng kết nối FTP
        ftp.quit()
        
def read_line(pth):
    with open(pth, 'r') as file:
        lines = [line.strip() for line in file.readlines()]
    return lines

def interpolate_dataframe(df):
    df['time'] = pd.to_datetime(df['time'], format='%d/%m/%Y %H:%M')
    # Chuyển cột 'time' thành dạng số giây kể từ thời điểm ban đầu
    df['time_seconds'] = (df['time'] - df['time'].min()).dt.total_seconds()
    start_time = df['time'].min()
    end_time = df['time'].max()
    hourly_timestamps = pd.date_range(start_time, end_time, freq='H')
    data = pd.DataFrame()
    data['time'] = hourly_timestamps
    for col in df.iloc[:,1:].columns:
        # tao ham noi suy
        f = interpolate.interp1d(df['time_seconds'], df[col], kind='linear', fill_value="extrapolate")
        # Chuyển các timestamp thành số giây kể từ thời điểm ban đầu
        hourly_timestamps_seconds = (hourly_timestamps - df['time'].min()).total_seconds()
        # nội suy value
        interpolated_values = f(hourly_timestamps_seconds)
        # tạo DataFrame
        result_df = pd.DataFrame({'time': hourly_timestamps, col: interpolated_values})
        data = data.merge(result_df,how='left',on='time')
    return data


def solieusongtranh():
    df = pd.read_excel(r'D:\PM_PYTHON\SONGTRANH\BANTIN\H,Q ho2023.xls',sheet_name='Thang11')
    df.columns = df.loc[0]
    df = df.iloc[1:25,:36*4+1]
    # print(df)
    # Chia thành các DataFrame 5 cột
    df_list = [df.iloc[:, i:i+4] for i in range(1, len(df.columns), 4)]
    # print(df_list)

    # for dta in df_list:
    #     print(dta)

    # Ghép các DataFrame lại theo chiều dọc 
    df_concat = pd.concat(df_list, axis=0)
    bd = datetime(2023,10,30)
    df_concat.insert(0,'time',pd.date_range(bd,periods=df_concat.shape[0],freq='H'))
    # df_concat.to_excel('BANTIN/Songtranh9.xlsx')

    now = datetime.now()
    now = datetime(now.year,now.month,now.day,now.hour)
    df_concat = df_concat[(df_concat['time']> now - timedelta(days=10)) & (df_concat['time']< now + timedelta(days=1))]
    # print(df_concat)
    return df_concat.iloc[:,:5]

# solieusongtranh()
def vedothihangngay():
    # pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'TONGHOP.xlsx')
    # now = datetime.now()
    # bd = datetime(int(now.strftime('%Y')),int(now.strftime('%m')),int(now.strftime('%d')),7)
    # kt = bd - timedelta(days=10)
    # df = pd.read_excel(pth,sheet_name='H')
    # df.rename(columns={'Ngày':'time','Trà Khúc':'trakhuc','Sông Vệ':'songve','Trà Bồng\n(Châu Ổ)':'chauo','Trà Câu':'tracau'},inplace=True)
    # df=df[['time','trakhuc','songve','chauo','tracau']]
    # dt_rang = pd.date_range(start=datetime(2022,1,1,1),periods=len(df['time']), freq="H")
    # df['time'] =dt_rang

    df = solieusongtranh()
    df['tongxa'] = df['Qmáy']   +   df['Qxả']
    # print(df)
    # ve bieu do muc nuoc
    df.iloc[:,1:] = df.iloc[:,1:].astype(float) 
    df['H hồ'] = df['H hồ'].astype(float)
    df['Qđến'] = df['Qđến'].astype(float)
    df.iloc[:,1:] = df.iloc[:,1:].interpolate(method='linear')
    fig, ax  = plt.subplots(figsize=(15, 12))

    ax.plot(df['time'],df['H hồ'],color = 'black')

    plt.legend(['Mực nước Sông Tranh'],prop={'size': 20})
    nday = mdates.HourLocator(interval=11)
    ax.xaxis.set_major_locator(nday)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y %Hh'))
    plt.xticks(rotation=60, size=14)
    ax.set_xlabel('Thời Gian',size = 20)
    ax.set_ylabel('Mực nước(m)',size = 20)

    plt.tight_layout(pad=6)
    plt.title('QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VỀ HỒ',size = 25)
    plt.grid(color = 'green', linestyle = '--', linewidth = 0.5)
    plt.savefig('image/chart_H.png')
    # plt.show()
    
    # ve bieu do lưu lượng đến
    df.iloc[:,1:] = df.iloc[:,1:].astype(float) 
    df['H hồ'] = df['H hồ'].astype(float)
    df['Qđến'] = df['Qđến'].astype(float)
    df.iloc[:,1:] = df.iloc[:,1:].interpolate(method='linear')
    fig, ax  = plt.subplots(figsize=(15, 12))

    ax.plot(df['time'],df['Qđến'],color = 'black')
    # ax.plot(df['time'],df['tongxa'],color = 'b')
    
    plt.legend(['Q đến'],prop={'size': 20})
    nday = mdates.HourLocator(interval=11)
    ax.xaxis.set_major_locator(nday)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y %Hh'))
    plt.xticks(rotation=60, size=14)
    ax.set_xlabel('Thời Gian',size = 20)
    ax.set_ylabel('Q(m3/s)',size = 20)
    plt.legend(['Q đến'],prop={'size': 20})
    plt.tight_layout(pad=6)
    plt.title('QUÁ TRÌNH LƯU LƯỢNG NƯỚC',size = 25)
    plt.grid(color = 'green', linestyle = '--', linewidth = 0.5)
    plt.savefig('image/chart_Q.png')
    # plt.show()
    
    # ve bieu do Q xả
    df['tongxa'] = df['tongxa'].interpolate(method='linear')
    fig, ax  = plt.subplots(figsize=(15, 12))

    ax.plot(df['time'],df['tongxa'],color = 'r')
    # ax.plot(df['time'],df['tongxa'],color = 'b')
    
    plt.legend(['Q xả'],prop={'size': 20})
    nday = mdates.HourLocator(interval=11)
    ax.xaxis.set_major_locator(nday)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y %Hh'))
    plt.xticks(rotation=60, size=14)
    ax.set_xlabel('Thời Gian',size = 20)
    ax.set_ylabel('Q(m3/s)',size = 20)
    plt.legend(['Q xả'],prop={'size': 20})
    plt.tight_layout(pad=6)
    plt.title('QUÁ TRÌNH LƯU LƯỢNG NƯỚC RA',size = 25)
    plt.grid(color = 'green', linestyle = '--', linewidth = 0.5)
    plt.savefig('image/chart_Q_xa.png')
    
# vedothihangngay()
def TTB_API_SONGTRANH(matram):
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,now.hour)
    bd = kt - timedelta(days=5)
    # mua
    pth = 'http://113.160.225.84:2018/API_TTB/JSON/solieu.php?matram={}&ten_table={}&sophut=60&tinhtong=0&thoigianbd=%27{}%2000:00:00%27&thoigiankt=%27{}%2023:59:00%27'
    pth = pth.format(matram,'mua_songtranh',bd.strftime('%Y-%m-%d'),kt.strftime('%Y-%m-%d'))
    df = pd.read_json(pth)
    return df
import numpy as np
def vebieudomua():
    trammua = ['Sông Tranh','Trà Bui','Trà Giác','Trà Dơn','Trà Leng','Trà Mai','Trà Cang','Trà Vân','Trà Nam','Trà Linh']
    trammua_eng = ['tramdapst2','TRABUI','tragiac','tradon','traleng','TRAMAI','tracang','travan','tranam2','tralinh']
    # for i in range(1):
    for i in range(len(trammua)):
        df = TTB_API_SONGTRANH(trammua_eng[i])
        print(df)
        fig, ax  = plt.subplots(figsize=(15, 10))
        # print(df['Thoigian_SL'])
        ax.bar(df['Thoigian_SL'], df['Solieu'])
        ax.set_xticks(np.arange(len(df))[::6]) # Chọn các tick cách 6 giá trị
        ax.set_xticklabels(df['Thoigian_SL'][::6])
        plt.xticks(rotation=60, size=14)
        plt.legend(['Mưa(mm)'],prop={'size': 10})
        plt.title('Biểu đồ mưa trạm ' + trammua[i])
        plt.xlabel('Thời gian')
        plt.ylabel('Lượng mưa (mm)')
        plt.tight_layout(pad=6)
        plt.savefig(r'D:\PM_PYTHON\SONGTRANH\image' + '/chart_mua_'+ trammua_eng[i] + '.png')
        # print(read_line(r'D:\PM_PYTHON\SONGTRANH\url_sever\LULU.txt'))
        upload_file(r'D:\PM_PYTHON\SONGTRANH\image' + '/chart_mua_'+ trammua_eng[i] + '.png',read_line(r'D:\PM_PYTHON\SONGTRANH\url_sever\LULU.txt')[2],read_line(r'D:\PM_PYTHON\SONGTRANH\infor\dakdrinh.txt')[0],read_line(r'D:\PM_PYTHON\SONGTRANH\infor\dakdrinh.txt')[1]) # gui ảnh
        # plt.show()
      
# vebieudomua()



def xacdinhbuocngay():
    now = datetime.now()
    for a in range(3,10):
        tttt = now + timedelta(days=a)
        if tttt.strftime('%d')[-1]=='1' or tttt.strftime('%d')[-1]=='6' and ('3' not in tttt.strftime('%d')) :
            ngay = datetime(tttt.year,tttt.month,tttt.day,23)
            break
    return ngay


def vedothituan():
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    now = datetime.now()
    bddb = datetime(now.year,now.month,now.day,23)
    now = now - timedelta(days=1)
    bd = datetime(now.year,now.month,now.day,23)
    kt = bd - timedelta(days=60)

    df = pd.read_excel(pth,sheet_name='H')
    df.rename(columns={'Ngày':'time','Trà Khúc':'trakhuc','Sông Vệ':'songve','Trà Bồng\n(Châu Ổ)':'chauo','Trà Câu':'tracau'},inplace=True)
    df=df[['time','chauo','trakhuc','songve','tracau']]
    dt_rang = pd.date_range(start=datetime(2022,1,1,1),periods=len(df['time']), freq="H")
    df['time'] =dt_rang
    df = df.loc[(df['time'] >= kt) & (df['time'] <= bd)]
    data = df.rolling(24*5,min_periods=1).agg(['mean','max','min'])
    data['time']=df['time']
    data = data[(data['time'].dt.hour==23) & ((data['time'].dt.day==30)|(data['time'].dt.day==5)|(data['time'].dt.day==10)|(data['time'].dt.day==15)|(data['time'].dt.day==20)|(data['time'].dt.day==25))]
    data = data.iloc[1:,:]
    
    # them gia tri du bao vao dataframe
    df = pd.read_excel(pth,sheet_name='TVHV')
    df = df.iloc[2:,5:]
    db = []
    for i in range(4):
        for a in df.iloc[i]:
            db.append(a)
    db.append(xacdinhbuocngay())
    data.loc[len(data['time'])] = db
    # print(data)
    fig, ax  = plt.subplots(2,2,figsize=(20, 12))
    hh = len(data['time'])-1

    ax[0,0].plot(data['time'].head(hh),data['chauo']['mean'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['max'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['min'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].set_xlabel('Thời Gian',size = 11)
    ax[0,0].set_ylabel('Mực nước(cm)',size = 11)
    ax[0,0].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ BỒNG TẠI TRẠM CHÂU Ổ',size=11)
    ax[0,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['mean'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['max'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['min'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].set_xlabel('Thời Gian',size = 11)
    ax[0,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[0,1].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ KHÚC TẠI TRẠM TRÀ KHÚC',size=11)
    ax[0,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,0].plot(data['time'].head(hh),data['songve']['mean'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['max'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['min'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].set_xlabel('Thời Gian',size = 11)
    ax[1,0].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,0].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG VỆ TẠI TRẠM SÔNG VỆ',size=11)
    ax[1,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,1].plot(data['time'].head(hh),data['tracau']['mean'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['max'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['min'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].set_xlabel('Thời Gian',size = 11)
    ax[1,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,1].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ CÂU TẠI TRẠM TRÀ CÂU',size=11)
    ax[1,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    plt.tight_layout(pad=4)
    plt.savefig('image/TVHV_05.png')
    # plt.show()
    messagebox.showinfo('Thông báo','OK!')

# vedothituan()

def vedothithang():
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    now = datetime.now()
    tgt = datetime.now()
    bddb = datetime(now.year,now.month,now.day,1)
    now = now - timedelta(days=1)
    bd = datetime(now.year,now.month,now.day,23)
    kt = datetime(2022,1,1,1)

    df = pd.read_excel(pth,sheet_name='H')
    df.rename(columns={'Ngày':'time','Trà Khúc':'trakhuc','Sông Vệ':'songve','Trà Bồng\n(Châu Ổ)':'chauo','Trà Câu':'tracau'},inplace=True)
    df=df[['time','chauo','trakhuc','songve','tracau']]
    dt_rang = pd.date_range(start=datetime(2022,1,1,1),periods=len(df['time']), freq="H")
    df['time'] =dt_rang
    df = df.loc[(df['time'] >= kt) & (df['time'] <= bd)]
    # print(df)
    
    
    data = df.rolling(24*30,min_periods=1).agg(['mean','max','min'])
    # print(data)
    data['time']=df['time']
    data = data[(data['time'].dt.hour==1) & (data['time'].dt.day==1)]
    # print(data)
    data = data.iloc[1:,:]
    # print(data)
    # them gia tri du bao vao dataframe
    df = pd.read_excel(pth,sheet_name='TVHD')
    
    df = df.iloc[2:,-3:]
    df = df[['Unnamed: 14','Unnamed: 15','Unnamed: 16']]
    # print(df)
    db = []
    for i in range(4):
        for a in df.iloc[i]:
            db.append(a)
    db.append(bddb)
    
    data.loc[len(data['time'])] = db
    # print(data)
    fig, ax  = plt.subplots(2,2,figsize=(20, 12))
    hh = len(data['time'])-1

    ax[0,0].plot(data['time'].head(hh),data['chauo']['mean'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['max'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['min'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].set_xlabel('Thời Gian',size = 11)
    ax[0,0].set_ylabel('Mực nước(cm)',size = 11)
    img_name = 'ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO THÁNG {} SÔNG TRÀ BỒNG TẠI TRẠM CHÂU Ổ'.format(tgt.strftime('%m'))
    ax[0,0].set_title(img_name,size=11)
    ax[0,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['mean'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['max'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['min'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].set_xlabel('Thời Gian',size = 11)
    ax[0,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[0,1].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO THÁNG {} SÔNG TRÀ KHÚC TẠI TRẠM TRÀ KHÚC'.format(tgt.strftime('%m')),size=11)
    ax[0,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,0].plot(data['time'].head(hh),data['songve']['mean'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['max'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['min'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].set_xlabel('Thời Gian',size = 11)
    ax[1,0].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,0].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO THÁNG {} SÔNG VỆ TẠI TRẠM SÔNG VỆ'.format(tgt.strftime('%m')),size=11)
    ax[1,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,1].plot(data['time'].head(hh),data['tracau']['mean'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['max'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['min'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].set_xlabel('Thời Gian',size = 11)
    ax[1,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,1].set_title(('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO THÁNG {} SÔNG TRÀ CÂU TẠI TRẠM TRÀ CÂU').format(tgt.strftime('%m')),size=11)
    ax[1,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    plt.tight_layout(pad=4)
    plt.savefig('image/Dothithang.png')
    messagebox.showinfo('Thông báo','OK!')
    
# plt.show()
# vedothithang()

def vedothituan10():
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    now = datetime.now()
    bddb = datetime(now.year,now.month,now.day,23)
    now = now - timedelta(days=1)
    bd = datetime(now.year,now.month,now.day,23)
    kt = bd - timedelta(days=90)

    df = pd.read_excel(pth,sheet_name='H')
    df.rename(columns={'Ngày':'time','Trà Khúc':'trakhuc','Sông Vệ':'songve','Trà Bồng\n(Châu Ổ)':'chauo','Trà Câu':'tracau'},inplace=True)
    df=df[['time','chauo','trakhuc','songve','tracau']]
    dt_rang = pd.date_range(start=datetime(2022,1,1,1),periods=len(df['time']), freq="H")
    df['time'] =dt_rang
    df = df.loc[(df['time'] >= kt) & (df['time'] <= (bd + timedelta(days=1)))]
    data = df.rolling(24*10,min_periods=1).agg(['mean','max','min'])
    data['time']=df['time']
    data = data[(data['time'].dt.hour==23) & ((data['time'].dt.day==1)|(data['time'].dt.day==11)|(data['time'].dt.day==21))]
    data = data.iloc[1:,:]
    # print(data)
    # them gia tri du bao vao dataframe
    df = pd.read_excel(pth,sheet_name='TVHV10')
    df = df.iloc[2:,11:14]
    # print(df)
    db = []
    for i in range(4):
        for a in df.iloc[i]:
            db.append(a)
    
    db.append(bddb+timedelta(days=10))
    # print(db)
    data.loc[len(data['time'])] = db
    # print(data)
    fig, ax  = plt.subplots(2,2,figsize=(20, 12))
    hh = len(data['time'])-1

    ax[0,0].plot(data['time'].head(hh),data['chauo']['mean'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['max'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].plot(data['time'].head(hh),data['chauo']['min'].head(hh))
    ax[0,0].plot(data['time'].tail(2),data['chauo']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,0].set_xlabel('Thời Gian',size = 11)
    ax[0,0].set_ylabel('Mực nước(cm)',size = 11)
    ax[0,0].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ BỒNG TẠI TRẠM CHÂU Ổ',size=11)
    ax[0,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['mean'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['max'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].plot(data['time'].head(hh),data['trakhuc']['min'].head(hh))
    ax[0,1].plot(data['time'].tail(2),data['trakhuc']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[0,1].set_xlabel('Thời Gian',size = 11)
    ax[0,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[0,1].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ KHÚC TẠI TRẠM TRÀ KHÚC',size=11)
    ax[0,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[0,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,0].plot(data['time'].head(hh),data['songve']['mean'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['max'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].plot(data['time'].head(hh),data['songve']['min'].head(hh))
    ax[1,0].plot(data['time'].tail(2),data['songve']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,0].set_xlabel('Thời Gian',size = 11)
    ax[1,0].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,0].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG VỆ TẠI TRẠM SÔNG VỆ',size=11)
    ax[1,0].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,0].grid(color = 'k', linestyle = '--', linewidth = 0.5)


    ax[1,1].plot(data['time'].head(hh),data['tracau']['mean'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['mean'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['max'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['max'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].plot(data['time'].head(hh),data['tracau']['min'].head(hh))
    ax[1,1].plot(data['time'].tail(2),data['tracau']['min'].tail(2),linestyle = 'dashed',marker = 'o')
    ax[1,1].set_xlabel('Thời Gian',size = 11)
    ax[1,1].set_ylabel('Mực nước(cm)',size = 11)
    ax[1,1].set_title('ĐƯỜNG QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TUẦN SÔNG TRÀ CÂU TẠI TRẠM TRÀ CÂU',size=11)
    ax[1,1].legend(['Trung bình','Trung bình dự báo','Max','Max dự báo','Min','Min dự báo'],prop={'size': 11})
    ax[1,1].grid(color = 'k', linestyle = '--', linewidth = 0.5)

    plt.tight_layout(pad=4)
    plt.savefig('image/Dothituan10.png')
    messagebox.showinfo('Thông báo','OK!')
# vedothituan10()