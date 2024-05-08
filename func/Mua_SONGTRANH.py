from tkinter import messagebox
from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
import pandas as pd
from datetime import datetime,timedelta
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
import pyautogui
import  numpy as np
import pyodbc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import  mysql.connector
def creat_cxn():
    # Kết nối đến MySQL
    host = '113.160.225.84'
    user = 'qltram'
    password = 'mhq@123456'
    port = 3306
    database = 'datasolieu'
    cnx = mysql.connector.connect(host=host, user=user, password=password, port=port, database=database)
    return cnx

def query_mysql(list_import,table_clounms,table_name):#TidVerticalIDVelocityForDetailMeasurement 
    if  str(list_import[2])=='nan':
        return None
    else:
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


def insert_data_sql(df,table_name):
    df.insert(0,'Matram','5ST')
    df['sldungduoc'] = df[df.columns[2]]
    df['maloi'] = 0
    df['chinhly'] = 0
    df = df.sort_values(by='time')
    # df =df.replace(np.nan,None)
    
    # Tạo kết nối
    cnx = creat_cxn()
    # Tạo con trỏ
    cursor = cnx.cursor(buffered=True)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    query = f"SELECT * FROM {table_name} LIMIT 1"
    cursor.execute(query)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    column_names = [column[0] for column in cursor.description]
    for index, row in df.iterrows():
        
        data = row.values.tolist()
        sql = query_mysql(data,column_names,table_name)
        # print(sql)
        try:
            cursor.execute(sql)
            cnx.commit()
        except:
            pass
    cursor.close()
    cnx.close()

def mucnuocsongtranh():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    pth = 'https://hochuathuydien.evn.com.vn/login.aspx'
    driver.maximize_window()
    driver.get(pth)
    time.sleep(2)
    driver.find_element(by=By.ID,value='ContentPlaceHolder1_txtAdminUser').send_keys('pclbtw')
    driver.find_element(by=By.ID,value='ContentPlaceHolder1_txtAdminPass').send_keys('1234567')
    driver.find_element(by=By.ID,value='ContentPlaceHolder1_btn_sumit').click()
    
    select_element = Select(driver.find_element(By.ID,'ContentPlaceHolder1_ctl00_dropluuvucsong'))
    # Chọn tùy chọn bằng giá trị
    select_element.select_by_value('21')
    
    select_element = Select(driver.find_element(By.ID,'ContentPlaceHolder1_ctl00_dropLake'))
    select_element.select_by_value('32')
    ngaythang = driver.find_element(By.ID,'ctl00_ContentPlaceHolder1_ctl00_txtFromDate_dateInput')
    ngaythang.clear()
    ngaythang.send_keys((datetime.now() -  timedelta(days=1)).strftime('%d/%m/%Y'))
    driver.find_element(By.ID,'ContentPlaceHolder1_ctl00_LinkButton1').click()
    time.sleep(2)
    
    # ban muc nuoc
    table_element = driver.find_element(by=By.XPATH, value= '//*[@id="ContentPlaceHolder1_ctl00_lblText"]/table')
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)
    # print(table_header[:12])
    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
        
    # print(table_header)
    # print(table_data)
    df = pd.DataFrame(table_data[4:])
    print(df)
    # Ghép cột 1 và cột 2 thành một cột datetime
    df.insert(0,'time',pd.to_datetime(df[1] + ' ' + df[0],format='%d-%m-%Y %H:%M'))
    # print(df)
    # Xóa cột 0 và 1 nếu bạn muốn
    df = df.drop(columns=[0, 1,8,9])
    # print(df)
    return df


def muasongtranh():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    now = datetime.now()
    bd = now - timedelta(hours=23)
    pth = 'http://songtranh2.tramthoitiet.vn/#/map/index'
    driver.maximize_window()
    driver.get(pth)
    time.sleep(2)
    pyautogui.hotkey('tab')
    pyautogui.write('songtranh2')
    pyautogui.hotkey('tab')
    pyautogui.write('Songtranh2@2022')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(3)
    date = driver.find_element(by=By.XPATH,value='//*[@id="main-container"]/div/div/div[2]/div[1]/div/div/div[1]/label/div/div/div')
    date.click()
    time.sleep(1)

    for a in range(1,42):
        element = driver.find_element(by=By.CSS_SELECTOR,value='#q-portal--menu--1 > div > div > div.q-date__main.col.column > div.q-date__content.col.relative-position > div > div.q-date__calendar-days-container.relative-position.overflow-hidden > div > div:nth-child({})'.format(str(a)))
        if str(element.text) == str(now.day):
            element.click()
            try:
                element = driver.find_element(by=By.CSS_SELECTOR,value='#q-portal--menu--1 > div > div > div.q-date__main.col.column > div.q-date__content.col.relative-position > div > div.q-date__calendar-days-container.relative-position.overflow-hidden > div > div:nth-child({})'.format(str(a-1)))
                element.click()
            except:
                pass
            break
    element = driver.find_element(by=By.CSS_SELECTOR,value='#q-portal--menu--1 > div > div > div.q-date__main.col.column > div.q-date__actions > div > button.q-btn.q-btn-item.non-selectable.no-outline.q-btn--unelevated.q-btn--rectangle.bg-primary.text-white.q-btn--actionable.q-focusable.q-hoverable > span.q-btn__content.text-center.col.items-center.q-anchor--skip.justify-center.row > span')  
    element.click()
    time.sleep(2)
    #precipitation > div > div.q-item.q-item-type.row.no-wrap > div:nth-child(2) > button > span.q-btn__content.text-center.col.items-center.q-anchor--skip.justify-center.row > i
    # '//*[@id="precipitation"]/div/div[1]/div[2]/button/span[2]/i'
    # element = driver.find_element(by=By.CSS_SELECTOR,value='#precipitation > div > div.q-item.q-item-type.row.no-wrap > div:nth-child(2) > button > span.q-btn__content.text-center.col.items-center.q-anchor--skip.justify-center.row > i')
    # element.click()
    
    # # muc nuoc
    # table_element = driver.find_element(by=By.XPATH, value= "/html/body/div[1]/div/div/div[2]/div/div/div[2]/div[2]/div/div[1]/div/div/div[2]/div/div/table")

    # table_header = []
    # for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
    #     table_header.append(th.text)
        
    # table_header.pop(1)
    # # print(table_header)
    # table_data = []
    # for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
    #     row_data = []
    #     for cell in row.find_elements(by=By.TAG_NAME,value="td"):
    #         row_data.append(cell.text)
    #     table_data.append(row_data)
    # # print(table_data)
    # df_q = pd.DataFrame(table_data,columns=table_header)

    # mưa
    table_element = driver.find_element(by=By.XPATH, value= "/html/body/div[1]/div/div/div[2]/div/div/div[2]/div[2]/div/div[4]/div/div/div[2]/div/div/table")
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)
    # print(table_header[:12])
    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    # print(table_data)
    df_mua = pd.DataFrame(table_data,columns=table_header[:13])
    # df_mua = df_mua.iloc[3:-1,:]
    df_mua.set_index('Thời gian',inplace=True)
    df_mua.sort_index(inplace=True)
    df_mua = df_mua.replace('-',np.nan)
    df_mua = df_mua.apply(pd.to_numeric, errors='coerce')
    # df_mua.to_excel('muasontranh.xlsx')
    # print(df_mua.info())
    mua = df_mua.rolling(6,min_periods=1).sum()
    mua.index = pd.to_datetime(mua.index)
    mua = mua[mua.index.minute==0]
    # mua = mua.sort_index()
    # print(mua)
    # print(df_q)
    driver.quit()
    return mua


# muasongtranh()

def query_sql(list_import,table_clounms,table_name):#TidVerticalIDVelocityForDetailMeasurement  
      
    sql = 'INSERT INTO ' + table_name + '('
    gt =  " VALUES ("
    for a in range(len(table_clounms)):
        # print(table_clounms[a])
        # print(type(list_import[a]))
        if str(list_import[a]) != 'nan' and list_import[a] is not None:
            sql = sql + table_clounms[a]+ ','
            gt = gt + ',\'{}\''.format(list_import[a])
    sql = sql  + ')' + gt + ')'
    sql = sql.replace(',)',')')
    sql = sql.replace('(,','(')
    return sql
import os
def insert_data(df,table_name):
    # print(df)
    FileName=(os.getcwd() + '/DATA/DATA.accdb')
    # print(FileName)
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
        # data[0] = row[0]
        # print(len(column_names))
        # print(data)
        # print(len(data))
        sql = query_sql(data,column_names,table_name)
        # print(sql)
        # cursor.execute(sql)
        # cnx.commit()
        try:
            cursor.execute(sql)
            cnx.commit()
        except:
            pass
    cursor.close()
    cnx.close()
    
def save_solieu_mua():
    df = muasongtranh()
    # print(df)
    df.reset_index(inplace=True)
    insert_data(df,'mua')
    # try:
    #     insert_data_sql(df[['time',4]],'ho_dakdrinh_qve')
    #     insert_data_sql(df[['time',5]],'ho_dakdrinh_qdieutiet')
    #     insert_data_sql(df[['time',2]],'ho_dakdrinh_mucnuoc')
    # except:
    #     pass
    
    messagebox.showinfo('Thông báo','OK!')

def save_solieu_mucnuoc():
    # df = pd.read_excel('DATA/data.xlsx')
    # df =df.sort_values('time')
    # df = df[['time','h','hhadu','qden','qxa','qcm','qxatran']]
    # print(df)
    df = mucnuocsongtranh()
    df.sort_values(by='time')
    insert_data(df,'thuyvan')
    try:
        insert_data_sql(df[['time',4]],'ho_dakdrinh_qve')
        insert_data_sql(df[['time',5]],'ho_dakdrinh_qdieutiet')
        insert_data_sql(df[['time',2]],'ho_dakdrinh_mucnuoc')
        insert_data_sql(df[['time',6]],'ho_dakdrinh_qchaymay')
        insert_data_sql(df[['time',7]],'ho_dakdrinh_qxatran')        
    except:
        pass
    messagebox.showinfo('Thông báo','OK!')

# save_solieu_mua()
# save_solieu_mucnuoc()