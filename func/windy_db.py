from selenium import webdriver
import pyautogui
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from datetime import datetime
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import pandas as pd
import pyodbc
from datetime import datetime,timedelta
from tkinter import messagebox
def query_sql(list_import,table_clounms,table_name):#TidVerticalIDVelocityForDetailMeasurement    
    sql = 'INSERT INTO ' + table_name + '('
    gt =  " VALUES ("
    for a in range(len(table_clounms)):
        # print(table_clounms[a])
        # print(list_import[a])
        if str(list_import[a]) != 'nan':
            sql = sql + table_clounms[a]+ ','
            gt = gt + ',\'{}\''.format(list_import[a])
    sql = sql  + ')' + gt + ')'
    sql = sql.replace(',)',')')
    sql = sql.replace('(,','(')
    return sql
import os
def insert_data(df,table_name):
    FileName=(os.getcwd()+ '/dungquat.accdb')
    cnx = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + FileName + ';')
    # Tạo con trỏ
    cursor = cnx.cursor()
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    query = f"SELECT * FROM {table_name} WHERE 1=0"
    cursor.execute(query)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    column_names = [column[0] for column in cursor.description]
    
    # print(column_names)
    for index, row in df.iterrows():
        data = row.values.tolist()
        sql = query_sql(data,column_names,table_name)
        # print(sql)
        try:
            cursor.execute(sql)
            cnx.commit()
        except:
            pass
    cursor.close()
    cnx.close()
    
def luusl_web():
    pth = 'https://thuyvan.hoaphatdungquat.vn/workspaces/1/dashboards/1'
    drive = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    drive.get(pth)
    drive.maximize_window()
    drive.find_element(by=By.NAME,value="username").send_keys('kttvqngai@gmail.com')
    drive.find_element(by=By.NAME,value="password").send_keys('Kttv1234')
    mat_select = drive.find_element(by=By.XPATH, value='/html/body/div/div[2]/div[3]/div/div/form/div/div[4]/div/button/span[1]')
    mat_select.click()
    time.sleep(3)
    pth = 'https://thuyvan.hoaphatdungquat.vn/workspaces/1/dashboards/1'
    drive.get(pth)
    time.sleep(5)
    # lay so lieu tram phao
    table_element = drive.find_element(by=By.CSS_SELECTOR, value="#main-layout > div.content > div > div.dashboard-body > div > div:nth-child(15) > div > div > table")
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)
      

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    
    df_phao = pd.DataFrame(table_data,columns=table_header[1:11])
    df_phao =df_phao.iloc[1:,:]
    df_phao.insert(0,'thoi gian',table_header[11:])


    # lay so lieu bang thong so song
    table_element = drive.find_element(by=By.CSS_SELECTOR, value="#main-layout > div.content > div > div.dashboard-body > div > div:nth-child(12) > div > div > table")
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)
      
    # print(table_header)
    
    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    
    df_tssong = pd.DataFrame(table_data,columns=table_header[1:5])
    df_tssong =df_tssong.iloc[1:,:]
    df_tssong.insert(0,'thoi gian',table_header[5:])

    # lay so lieu ca thong so khac
    table_element = drive.find_element(by=By.CSS_SELECTOR, value="#main-layout > div.content > div > div.dashboard-body > div > div:nth-child(13) > div > div > table")
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)
      
    # print(table_header)
    
    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    
    df_thongso = pd.DataFrame(table_data,columns=table_header[1:5])
    df_thongso =df_thongso.iloc[1:,:]
    df_thongso.insert(0,'thoi gian',table_header[5:])
    drive.quit()
    
    
    df = df_phao.merge(df_thongso,how='left',on= 'thoi gian')
    df = df.merge(df_tssong,how='left',on= 'thoi gian')
    
    # df.to_excel('solieudungquat2311.xlsx')
    
    # df = pd.read_excel('solieudungquat2311.xlsx')
    # df1 = pd.read_excel('solieudungquat2211.xlsx')
    # df2 = pd.read_excel('solieudungquat1911.xlsx')
    # df3 = pd.read_excel('solieudungquat.xlsx')
    # df = pd.concat([df,df1],axis=0)
    # df = pd.concat([df,df2],axis=0)
    # df = pd.concat([df,df3],axis=0)
    # df = df.sort_values(by=['thoi gian'])
    # df =df.loc[~df['thoi gian'].duplicated(keep='first')]
    # # df.iloc[:,1:].to_excel('tonghop.xlsx',index=False)
    # print(df)
    # df = pd.read_excel('tonghop.xlsx')
    # print(df)
    insert_data(df,'tramphao')
    
# luusl_web()  

def laysolieudubao_windy():
    now = datetime.now()
    now = datetime(now.year,now.month,now.day,0)
    data = pd.DataFrame()
    data['time'] = pd.date_range(now,now+timedelta(days=11),freq='h')
    # dfname = pd.read_csv('ts_id/WINDY.txt',sep="\s+",header=None)
    dfname = pd.read_csv(r'ts_id/WINDY.txt',sep=",",header=None)
    dfname.columns = dfname.iloc[0]
    dfname = dfname.iloc[1:,:]
    # print(dfname)
    pth ='https://www.windy.com/login?'
    drive = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    drive.get(pth)
    drive.maximize_window()
    time.sleep(2)
    drive.find_element(by=By.ID,value="email").send_keys('kttvqb')
    drive.find_element(by=By.ID,value="password").send_keys('kttvqb2022')
    drive.find_element(by=By.ID,value="submitLogin").click()
    time.sleep(2)

    for tsid in zip(dfname['lat'],dfname['long'],dfname['Matram']):
        # print(tsid)
        pth = 'https://www.windy.com/{}/{}?{},{},11,i:pressure'.format(tsid[0],tsid[1],tsid[0],tsid[1])
        drive.get(pth)
        time.sleep(2)
        table_element = drive.find_element(by=By.ID, value="detail-data-table")
        drive.find_element(by=By.XPATH,value = '//*[@id="detail-box-desktop"]/div[1]').click()
        time.sleep(1)
        drive.find_element(by=By.XPATH,value ='//*[@id="plugin-detail"]/div[2]/div[2]/div/div[1]/div[1]').click()
        time.sleep(2)
        # print(table_header)
        table_data = []
        for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
            row_data = []
            for cell in row.find_elements(by=By.TAG_NAME,value="td"):
                row_data.append(cell.text)
            table_data.append(row_data)
        # print(table_data)
        table_data = table_data[1:]
        # print(table_data)
        table_header = ['time','nhiet'+tsid[2],'mua'+tsid[2],'gio'+tsid[2],'gio giat'+tsid[2],'huong'+tsid[2]]
        df = pd.DataFrame(table_data)
        df = df.drop(1)
        df = df.T
        df.columns= table_header
        df['time'] = df['time'].astype(int)
        thoigian=[]
        bd = datetime.now().date()
        thoigian.append(bd)
        for a in range(1,len(df['time'])):
            if df['time'].loc[a-1] < df['time'].loc[a]:
                thoigian.append(bd)
            elif df['time'].loc[a-1] > df['time'].loc[a]:
                thoigian.append(bd + timedelta(days=1))
                bd = bd + timedelta(days=1)
        df['date'] = thoigian
        df['date'] = pd.to_datetime(df['date'])
        df['time'] = pd.to_timedelta(df['time'], unit='h')
        # df['date'] =df['date'].astype(str)
        # df['time'] = df['time'].astype(str)
        df.insert(0,'ngay',df['date'] + df['time'])
        df.drop(['time', 'date'], axis=1, inplace=True)
        df.rename(columns={'ngay':'time'},inplace=True)
        # print(df)
        data = data.merge(df,how='left',on='time')
        # df.to_excel('windy.xlsx',index=False)
        # print(df.to_excel('kiemtra.xlsx',index=False))
        # time.sleep(1)
        
        # df = pd.DataFrame(table_data,columns=table_header)

        # df =df.iloc[1:,:]
        # df.insert(0,'thoi gian',table_header[5:])
    drive.quit()
    data.to_excel('DATA/windy_db.xlsx',index=False)
    messagebox.showinfo('Thông báo!', 'OK')
    # return data


# df = laysolieudubao_windy()
# df.to_excel('kiemtra1.xlsx')