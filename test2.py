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
    FileName=(os.getcwd() + '/DATA/QNAM.accdb')
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
        sql = query_sql(data,column_names,table_name)
        # print(sql)
        try:
            cursor.execute(sql)
            cnx.commit()
        except:
            pass
    cursor.close()
    cnx.close()
    
def save_solieu_mua():
    df = pd.read_excel('DATA/20245.xlsx')
    df = df.iloc[1:,:]
    df['Thời gian']  = pd.to_datetime(df['Thời gian'])
    print(df)
    # df.reset_index(inplace=True)
    insert_data(df,'mua')
    # try:
    #     insert_data_sql(df[['time',4]],'ho_dakdrinh_qve')
    #     insert_data_sql(df[['time',5]],'ho_dakdrinh_qdieutiet')
    #     insert_data_sql(df[['time',2]],'ho_dakdrinh_mucnuoc')
    # except:
    #     pass
    
    # messagebox.showinfo('Thông báo','OK!')

save_solieu_mua()


def save_solieu_mucnuoc():
    df = pd.read_excel('DATA/20245.xlsx',sheet_name='Trạm đo lưu lượng')
    df = df.iloc[2:,:]
    print(df)
    df['Thời gian']  = pd.to_datetime(df['Thời gian'])
    df.rename(columns={'Thời gian':'time'},inplace=True)
    df.sort_values(by='time')
    df = df[['time','Trà Tập','Unnamed: 4','Unnamed: 5']]
    df.columns = ['time',]
    # # df.reset_index(inplace=True)
    print(df)
    # insert_data(df,'thuyvan')
    # try:
    #     insert_data_sql(df[['time',4]],'ho_dakdrinh_qve')
    #     insert_data_sql(df[['time',5]],'ho_dakdrinh_qdieutiet')
    #     insert_data_sql(df[['time',2]],'ho_dakdrinh_mucnuoc')
    # except:
    #     pass
    # messagebox.showinfo('Thông báo','OK!')

# save_solieu_mua()
# save_solieu_mucnuoc()