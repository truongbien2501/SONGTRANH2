from selenium import webdriver
import pyautogui
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select


def chon_kieu_tin(path_tin):
    if '_DIEM_' in  path_tin:
        gt ='3'
    elif '_HVHT_' in  path_tin:
        gt ='4'
    elif '_HVHT_' in  path_tin:
        gt ='5'
    elif '_HVHT_' in  path_tin:
        gt ='6'
    elif '_HVHT_' in  path_tin:
        gt ='7'
    elif '_HVHT_' in  path_tin:
        gt ='8'
    elif '_TVHM_' in  path_tin:
        gt ='9'    
    elif '_TVHN_' in  path_tin:
        gt ='10'
    elif '_TVHD_' in  path_tin:
        gt ='11'
    elif '_TVHV_' in  path_tin:
        gt ='12'
    elif '_HVHT_' in  path_tin:
        gt ='13'
    elif '_XTHE_' in  path_tin:
        gt ='14'
    elif '_HVHT_' in  path_tin:
        gt ='15'       
    elif 'NNHN' in  path_tin:
        gt ='16'
    elif '_HVHT_' in  path_tin:
        gt ='17'
    elif '_HVHT_' in  path_tin:
        gt ='18'
    elif '_HVHV_' in  path_tin:
        gt ='19'
    elif '_HVHT_' in  path_tin:
        gt ='20'
    elif '_HVHT_' in  path_tin:
        gt ='21'       
    elif '_HVHT_' in  path_tin:
        gt ='22'
    elif '_HVHT_' in  path_tin:
        gt ='23'
    elif '_XTND_' in  path_tin:
        gt ='24'
    elif '_HVHT_' in  path_tin:
        gt ='25'
    elif '_MLDR_' in  path_tin:
        gt ='26'
    elif '_MLDL_' in  path_tin:
        gt ='27'       
    elif '_KKLR_' in  path_tin:
        gt ='28'
    elif '_NONG_' in  path_tin:
        gt ='29'
    elif '_DONG_' in  path_tin:
        gt ='30'
    elif '_HVNH_' in  path_tin:
        gt ='31'
    elif '_HVHT_' in  path_tin:
        gt ='32'
    elif '_TTNH_' in  path_tin:
        gt ='33'       
    elif '_LULU_' in  path_tin:
        gt ='34'
    elif '_LQSL_' in  path_tin:
        gt ='35'
    elif '_HHAN_' in  path_tin:
        gt ='36'
    elif '_XMAN_' in  path_tin:
        gt ='37'
    elif '_HVHT_' in  path_tin:
        gt ='38'
    elif '_HVNH_' in  path_tin:
        gt ='39'     
    if int(gt) <=20:
        loaitin= '1'
    else:
        loaitin= '2'
    return loaitin,gt
        
def nhap_pass():
    pyautogui.hotkey('TAB')
    pyautogui.write('qnga')
    pyautogui.hotkey('TAB')
    pyautogui.write('qnga')
    pyautogui.hotkey('TAB')
    pyautogui.hotkey('enter')
    time.sleep(2)
    pyautogui.hotkey('enter')
   

def upload_ttdl(path_tin,path_hs):
    drive = webdriver.Chrome(ChromeDriverManager().install())
    # drive = webdriver.Chrome("chromedriver.exe")
    pth = 'http://222.255.11.117:8888/'
    drive.get(pth)
    drive.maximize_window()
    nhap_pass()
    time.sleep(3)
    pth = 'http://222.255.11.117:8888/UploadFile.aspx'
    drive.get(pth)
    time.sleep(3)
    # Chọn phần tử select
    select_element = Select(drive.find_element_by_css_selector("#cphContent_ddl_Kieu"))
    # Chọn tùy chọn bằng giá trị
    # 1: thoi tiet binh thuong
    # 2 thoi tiet nguy hiem
    loaitin, gt = chon_kieu_tin(path_tin)
    select_element.select_by_value(loaitin)
    select_element = Select(drive.find_element_by_css_selector("#cphContent_ddl_Loai"))
    # Chọn tùy chọn bằng giá trị
    select_element.select_by_value(gt)
    time.sleep(2)
    file_input = drive.execute_script('return document.querySelector("#cphContent_FileUpload1")')
    # Gửi đường dẫn tệp tin tới phần tử input file
    file_input.send_keys(path_tin)
    file_input = drive.execute_script('return document.querySelector("#cphContent_FileUpload2")')
    file_input.send_keys(path_hs)
    time.sleep(1)
    drive.find_element_by_id('cphContent_Button1').click()
    time.sleep(2)
    drive.quit()
