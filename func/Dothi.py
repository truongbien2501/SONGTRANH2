import array as arr
from ctypes.wintypes import SIZE
from datetime import datetime, timedelta
from turtle import st 
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure, legend, xlabel, ylabel
import matplotlib.dates as mdates
import pandas as pd
from tkinter import messagebox
from FUNC.Seach_file import tim_file,read_txt,vitridat

def vedothihangngay():
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    now = datetime.now()
    bd = datetime(int(now.strftime('%Y')),int(now.strftime('%m')),int(now.strftime('%d')),7)
    kt = bd - timedelta(days=10)
    df = pd.read_excel(pth,sheet_name='H')
    df.rename(columns={'Ngày':'time','Trà Khúc':'trakhuc','Sông Vệ':'songve','Trà Bồng\n(Châu Ổ)':'chauo','Trà Câu':'tracau'},inplace=True)
    df=df[['time','trakhuc','songve','chauo','tracau']]
    dt_rang = pd.date_range(start=datetime(2022,1,1,1),periods=len(df['time']), freq="H")
    df['time'] =dt_rang

# # 10 ngay anh huong trieu
# bd = datetime(int(now.strftime('%Y')),int(now.strftime('%m')),int(now.strftime('%d')),7)
# kt = bd - timedelta(days=11)
# df = df.loc[(df['time'] >= kt) & (df['time'] <= bd) ]
# dfmax = df.rolling(24,min_periods=1).max()
# dfmax['time'] = pd.to_datetime(df['time'])
# dfmax = dfmax.loc[dfmax['time'].dt.hour ==7]

# fig, ax  = plt.subplots(figsize=(20, 12))

# ax.plot(dfmax['time'],dfmax['songve'])
# ax.plot(dfmax['time'],dfmax['tracau'])
# ax.plot(dfmax['time'],dfmax['chauo'])

# nday = mdates.DayLocator(interval=1)
# ax.xaxis.set_major_locator(nday)
# ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y'))
# plt.xticks(rotation=60, size=14)
# plt.legend(['Sông Vệ','Trà Câu','Châu Ổ'],prop={'size': 20})
# plt.tight_layout(pad=6)
# plt.title('QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO',size = 30)
# plt.grid(color = 'green', linestyle = '--', linewidth = 0.5)
# plt.show()
# # print(dfmax)
# # print(df)


    df = df.loc[(df['time'] >= kt) & (df['time'] <= bd) ]
# fig, ax  = plt.subplots(1,1,figsize=(15, 7))
    fig, ax  = plt.subplots(figsize=(20, 12))

    ax.plot(df['time'],df['trakhuc'],color = 'r')
    ax.plot(df['time'],df['songve'],color = 'black')
    ax.plot(df['time'],df['chauo'],color = 'b')
    ax.plot(df['time'],df['tracau'],color = 'm')
# print(df)
    db = pd.DataFrame()
    dfdb =  pd.read_excel(pth,sheet_name='hangngay')


# print(df.tail(1))
    tk = dfdb.iloc[3,11:15]
    sv = dfdb.iloc[5,11:15]
    co = dfdb.iloc[6,11:15]
    tc = dfdb.iloc[7,11:15]

    trkdb =[df['trakhuc'].tail(1).iloc[0]]
    svdb =[df['songve'].tail(1).iloc[0]]
    codb =[df['chauo'].tail(1).iloc[0]]
    tcdb = [df['tracau'].tail(1).iloc[0]]

    for p in tk:
        trkdb.append(p)
    for p in sv:
        svdb.append(p)
    for p in co:
        codb.append(p)
    for p in tc:
       tcdb.append(p)


    tgdb =pd.date_range(start=bd,periods=5, freq="6H")
    db['time'] = tgdb
    db['tk'] = trkdb
    db['sv'] = svdb
    db['co'] = codb
    db['tc'] = tcdb
# print(db)

    plt.plot(db['time'],db['tk'],linestyle = 'dashed',marker = 'o',color = 'r')
    plt.plot(db['time'],db['sv'],linestyle = 'dashed',marker = 'o',color = 'black')
    plt.plot(db['time'],db['co'],linestyle = 'dashed',marker = 'o',color = 'b')
    plt.plot(db['time'],db['tc'],linestyle = 'dashed',marker = 'o',color = 'm')

    plt.legend(['Trà Khúc thực đo','Sông Vệ thực đo','Châu Ổ  thực đo','Trà Câu  thực đo','Trà Khúc dự báo','Sông Vệ dự báo','Châu Ổ dự báo','Trà Câu dự báo'],prop={'size': 20})

    nday = mdates.HourLocator(interval=11)
    ax.xaxis.set_major_locator(nday)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%y %Hh'))
    plt.xticks(rotation=60, size=14)

# ax.set(title='BIỂU ĐỒ MỰC NƯỚC THỰC ĐO 10 NGÀY',fontsize=30)
    ax.set_xlabel('Thời Gian',size = 20)
    ax.set_ylabel('Mực nước(cm)',size = 20)

# plt.xticks(rotation=90, size=6)
    plt.tight_layout(pad=6)
    plt.title('QUÁ TRÌNH MỰC NƯỚC THỰC ĐO VÀ DỰ BÁO TRÊN CÁC SÔNG TRONG TỈNH QUẢNG NGÃI',size = 25)
    plt.grid(color = 'green', linestyle = '--', linewidth = 0.5)
    plt.savefig('image/Dothi10ngay.jpg')
    messagebox.showinfo('Thông báo','OK')
    # plt.show()

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
    print(data)
    data['time']=df['time']
    data = data[(data['time'].dt.hour==1) & (data['time'].dt.day==1)]
    print(data)
    data = data.iloc[1:,:]
    print(data)
    # them gia tri du bao vao dataframe
    df = pd.read_excel(pth,sheet_name='TVHD')
    df = df.iloc[2:,-3:]
    # print(df)
    db = []
    for i in range(4):
        for a in df.iloc[i]:
            db.append(a)
    db.append(bddb)
    
    data.loc[len(data['time'])] = db
    print(data)
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