from selenium import webdriver
import time
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pyautogui
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime,timedelta
import numpy as np
from selenium.webdriver.chrome.service import Service as ChromeService
import cv2
import os
from PIL import Image
def vitri_click(im2):
    im1 = chup_man()
    gray = cv2.cvtColor(im1,cv2.COLOR_BGR2GRAY)
    template = cv2.imread(os.getcwd()+ '/image/' + im2,0)
    w, h = template.shape[::-1]
    res = cv2.matchTemplate(gray,template,cv2.TM_CCOEFF_NORMED)
    # print(res)
    loc = np.where(res>=0.8)
    x = 0
    y = 0
    for pt in zip(*loc[::-1]):
        # print(pt)
        # cv2.rectangle(img,pt,(pt[0]+w,pt[1]+h),(0,0,255),3)
        x = int(pt[0]+w/2)
        y = int(pt[1]+h/2)
    return x,y
    
def chup_man():
    screenshot  = pyautogui.screenshot() 
    img_capture = np.array(Image.frombytes('RGB', screenshot.size, screenshot.tobytes()))
    return img_capture

def mua_web_vrain():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    now = datetime.now()
    if now.hour >=20:
        ngay1 = datetime(now.year,now.month,now.day,20)
        ngay_t = now-timedelta(days=2)
        ngay2 = datetime(ngay_t.year,ngay_t.month,ngay_t.day,20)
        dieuhuong = 'right'
    else:
        ngay1  = now-timedelta(days=1)
        ngay1 = datetime(ngay1.year,ngay1.month,ngay1.day,20)
        ngay_t = now-timedelta(days=2)
        ngay2 = datetime(ngay_t.year,ngay_t.month,ngay_t.day,20)
        dieuhuong = 'left'
    pth = 'https://vrain.vn/landing'
    driver.maximize_window()
    driver.get(pth)
    time.sleep(5)
    element = driver.execute_script("return document.querySelector('body > watec-root > watec-landing > div > button > span.mdc-button__label')")
    element.click()
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.write('Muaqn@vinarain')
    pyautogui.hotkey('tab')
    pyautogui.write('123456')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(5)
    # click vào chi tiet
    element = driver.execute_script("return document.querySelector('body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-header > mat-toolbar > div:nth-child(2) > button:nth-child(2) > span.mdc-button__label > div > span')")
    element.click()
    time.sleep(5)

    # lay luong mua ngay hom nay
    table_element = driver.find_element(by=By.CSS_SELECTOR,value= "body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > section.dashboard-container__table-container.ng-star-inserted > table"
    )
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME,value="th"):
        table_header.append(th.text)

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)

    df = pd.DataFrame(table_data,columns=table_header)
    df = df.T
    df.drop(columns=0,inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[3:,:]
    tg_bd = df.index[0]
    ngay1 = datetime(ngay1.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
    # print(ngay1)
    df['time'] = pd.date_range(ngay1,periods=len(df.index),freq='H')
    # print(df)
    # lay luong mua ngay hom qua
    x_tk,y_k= vitri_click('mua_vr.png')
    pyautogui.click(x_tk,y_k)
    # pyautogui.moveTo(300,200)
    # pyautogui.click()
    time.sleep(2)
    # # Đặt giá trị của phần tử mat-select-value-5 thành ngày khác
    divs = driver.find_element(by=By.ID, value='mat-select-4-panel')
    mat_select = divs.find_element(by=By.CSS_SELECTOR, value='#mat-option-68 > span')
    mat_select.click()
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(1)
    pyautogui.hotkey(dieuhuong)
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(4)
    # cick vao xem
    element = driver.execute_script("return document.querySelector('body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > mat-toolbar > div > button > span.mdc-button__label')")
    element.click()
    time.sleep(5)
    table_element = driver.find_element(by=By.CSS_SELECTOR, value=
    "body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > section.dashboard-container__table-container.ng-star-inserted > table"
    )
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
        table_header.append(th.text)

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    driver.quit()
    df1 = pd.DataFrame(table_data,columns=table_header)
    df1 = df1.T
    df1.drop(columns=0,inplace=True)
    df1.columns = df1.iloc[0]
    df1 = df1.iloc[3:,:]
    tg_bd = df1.index[0]
    ngay2 = datetime(ngay2.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
    # print(ngay2)
    df1['time'] = pd.date_range(ngay2,periods=len(df1.index),freq='H')
    df = pd.concat([df,df1])
    df.insert(0, 'new_col_name', df['time'])
    df = df.drop(columns=['time'])  # Loại bỏ cột 'time' nếu cần
    # Đổi tên cột 'new_col_name' thành 'time'
    df = df.rename(columns={'new_col_name': 'time'})
    df = df.sort_values(by='time')
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,7)
    bd = kt - timedelta(days=1)
    df = df[(df['time']>bd) & (df['time']<=kt)]
    df.set_index('time',inplace=True)
    df = df.replace('-',np.nan)
    df =df.astype(float)
    return df
def mucnuoc_web_vrain():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    now = datetime.now()
    ngay1 = datetime(now.year,now.month,now.day,0,10)
    ngay2 = ngay1 - timedelta(days=1)
    ngay2 = datetime(ngay2.year,ngay2.month,ngay2.day,0,10)

    pth = 'https://mucnuoc.vrain.vn/login'
    driver.maximize_window()
    driver.get(pth)
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.write('mnquangbinh')
    pyautogui.hotkey('tab')
    pyautogui.write('123456')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(2)

    element = driver.execute_script("return document.querySelector('body > watec-root > watec-group-details > watec-vwater-header > mat-toolbar > button:nth-child(4) > span.mdc-button__label > div > span')")
    element.click()
    time.sleep(1)

    # ngay 1
    table_element = driver.find_element(by=By.CSS_SELECTOR,value="body > watec-root > watec-group-details > watec-vwater-details-page > div > section.details-page-container__table-container.ng-star-inserted > table")
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME,value="th"):
        table_header.append(th.text)

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)

    df = pd.DataFrame(table_data,columns=table_header)
    df = df.T
    df.drop(columns=0,inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[1:,:]
    tg_bd = df.index[0]
    ngay1 = datetime(ngay1.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
    # print(ngay1)
    df['time'] = pd.date_range(ngay1,periods=len(df.index),freq='10T')
    
    # lay luong muc nuoc ngay hom qua
    x_tk,y_k= vitri_click('H_vr.png')
    pyautogui.click(x_tk,y_k)
    time.sleep(2)
    # # Đặt giá trị của phần tử mat-select-value-5 thành ngày khác
    mat_select = driver.find_element(by=By.ID,value="mat-select-2-panel")
    mat_select.find_element(by=By.CSS_SELECTOR,value='#mat-option-6').click()

    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(1)
    pyautogui.hotkey('left')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(3)
    # cick vao xem
    div = driver.find_element(by=By.CSS_SELECTOR,value='body > watec-root > watec-group-details > watec-vwater-details-page > div > mat-toolbar > div')
    div.click()
    pyautogui.hotkey('enter')
    time.sleep(1)
    pyautogui.hotkey('tab')
    time.sleep(1)
    pyautogui.hotkey('enter')
    time.sleep(3)

    table_element = driver.find_element(by=By.CSS_SELECTOR,value=
    "body > watec-root > watec-group-details > watec-vwater-details-page > div > section.details-page-container__table-container.ng-star-inserted > table"
    )
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME,value="th"):
        table_header.append(th.text)

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)
    driver.quit()
    df1 = pd.DataFrame(table_data,columns=table_header)
    df1 = df1.T
    df1.drop(columns=0,inplace=True)
    df1.columns = df1.iloc[0]
    df1 = df1.iloc[1:,:]
    tg_bd = df1.index[0]
    ngay2 = datetime(ngay2.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
    # print(ngay2)
    df1['time'] = pd.date_range(ngay2,periods=len(df1.index),freq='10T')
    df = pd.concat([df,df1])
    df.insert(0, 'new_col_name', df['time'])
    df = df.drop(columns=['time'])  # Loại bỏ cột 'time' nếu cần
    # Đổi tên cột 'new_col_name' thành 'time'
    df = df.rename(columns={'new_col_name': 'time'})
    df = df.sort_values(by='time')
    df = df[df['time'].dt.minute==0]
    now = datetime.now()
    kt = datetime(now.year,now.month,now.day,7)
    bd = kt - timedelta(days=1)
    df = df[(df['time']>=bd) & (df['time']<=kt)]
    df.set_index('time',inplace=True)
    df = df.replace('-',np.nan)
    df =df.astype(float)
    df = df*100
    # df.to_excel('H.xlsx',index=False)
    return d

def mua_lqsl():
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    now = datetime.now()
    if now.hour >=20:
        ngay1 = datetime(now.year,now.month,now.day,20)
        ngay_t = now-timedelta(days=2)
        ngay2 = datetime(ngay_t.year,ngay_t.month,ngay_t.day,20)
        dieuhuong = 'right'
    else:
        ngay1  = now-timedelta(days=1)
        ngay1 = datetime(ngay1.year,ngay1.month,ngay1.day,20)
        ngay_t = now-timedelta(days=2)
        ngay2 = datetime(ngay_t.year,ngay_t.month,ngay_t.day,20)
        dieuhuong = 'left'
    pth = 'https://vrain.vn/landing'
    driver.maximize_window()
    driver.get(pth)
    time.sleep(5)
    element = driver.execute_script("return document.querySelector('body > watec-root > watec-landing > div > button > span.mdc-button__label')")
    element.click()
    time.sleep(1)
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.write('Muaqn@vinarain')
    pyautogui.hotkey('tab')
    pyautogui.write('123456')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(5)
    # click vào chi tiet
    element = driver.execute_script("return document.querySelector('body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-header > mat-toolbar > div:nth-child(2) > button:nth-child(2) > span.mdc-button__label > div > span')")
    element.click()
    time.sleep(5)

    # lay luong mua ngay hom nay
    table_element = driver.find_element(by=By.CSS_SELECTOR,value= "body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > section.dashboard-container__table-container.ng-star-inserted > table"
    )
    table_header = []
    for th in table_element.find_elements(by=By.TAG_NAME,value="th"):
        table_header.append(th.text)

    table_data = []
    for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
        row_data = []
        for cell in row.find_elements(by=By.TAG_NAME,value="td"):
            row_data.append(cell.text)
        table_data.append(row_data)

    df = pd.DataFrame(table_data,columns=table_header)
    df = df.T
    df.drop(columns=0,inplace=True)
    df.columns = df.iloc[0]
    df = df.iloc[3:,:]
    tg_bd = df.index[0]
    ngay1 = datetime(ngay1.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
    # print(ngay1)
    df['time'] = pd.date_range(ngay1,periods=len(df.index),freq='H')
    if now.hour <= 6:
        # print(df)
        # lay luong mua ngay hom qua
        x_tk,y_k= vitri_click('mua_vr.png')
        pyautogui.click(x_tk,y_k)
        # pyautogui.moveTo(300,200)
        # pyautogui.click()
        time.sleep(2)
        # # Đặt giá trị của phần tử mat-select-value-5 thành ngày khác
        divs = driver.find_element(by=By.ID, value='mat-select-4-panel')
        mat_select = divs.find_element(by=By.CSS_SELECTOR, value='#mat-option-160')
        mat_select.click()
        pyautogui.hotkey('tab')
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(1)
        pyautogui.hotkey(dieuhuong)
        time.sleep(1)
        pyautogui.hotkey('enter')
        time.sleep(3)
        # cick vao xem
        element = driver.execute_script("return document.querySelector('body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > mat-toolbar > div > button > span.mdc-button__label')")
        element.click()
        time.sleep(3)
        table_element = driver.find_element(by=By.CSS_SELECTOR, value=
        "body > watec-root > watec-group > mat-sidenav-container > mat-sidenav-content > watec-vrain-shared-detail > div > section.dashboard-container__table-container.ng-star-inserted > table"
        )
        table_header = []
        for th in table_element.find_elements(by=By.TAG_NAME, value="th"):
            table_header.append(th.text)

        table_data = []
        for row in table_element.find_elements(by=By.TAG_NAME,value="tr"):
            row_data = []
            for cell in row.find_elements(by=By.TAG_NAME,value="td"):
                row_data.append(cell.text)
            table_data.append(row_data)
        driver.quit()
        df1 = pd.DataFrame(table_data,columns=table_header)
        df1 = df1.T
        df1.drop(columns=0,inplace=True)
        df1.columns = df1.iloc[0]
        df1 = df1.iloc[3:,:]
        tg_bd = df1.index[0]
        ngay2 = datetime(ngay2.year,int(tg_bd[-2:]),int(tg_bd[6:8]),int(tg_bd[:2]),int(tg_bd[3:5]))
        # print(ngay2)
        df1['time'] = pd.date_range(ngay2,periods=len(df1.index),freq='H')
        df = pd.concat([df,df1])
    df.insert(0, 'new_col_name', df['time'])
    df = df.drop(columns=['time'])  # Loại bỏ cột 'time' nếu cần
    # Đổi tên cột 'new_col_name' thành 'time'
    df = df.rename(columns={'new_col_name': 'time'})
    df = df.sort_values(by='time')
    now = datetime.now()
    now = datetime(now.year,now.month,now.day,now.hour)
    bd = now - timedelta(hours=6)
    df = df[(df['time']>=bd)]
    df.set_index('time',inplace=True)
    df = df.replace('-',np.nan)
    df =df.astype(float)
    df.reset_index(drop=True)
    # df.to_excel('aaaaaaaaaaaaaaaaaaaaaaaaaa.xlsx')
    return df
# print(mua_web_vrain())
# print(mucnuoc_web_vrain())
# print(mua_lqsl())
