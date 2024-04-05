from datetime import datetime, timedelta
from tkinter import *
from tkinter import ttk
from func.Mua_SONGTRANH import save_solieu_mua,save_solieu_mucnuoc
from func.load_data import downloadattmail,write_rain,load_sl_thuyvan,load_sl_5day,mo_excel,downloadattmail_lulu,load_sl_thuyvan_lulu
from Run_tin.DRHN import tin_tvhn,tin_tv_load,nghiemthu_tvhn
from Run_tin.DRHV10 import tin_nenKT_10day,tin_tv10_load
from func.SEND_FILE import gui_drhd,gui_drhv,gui_drhn,gui_lulu
from Run_tin.DRLULU import tin_nenKT_lulu,tin_lulu_load,tin_lulu_load1
from Run_tin import DRHN,DRHV10,DRLULU
from func.Seach_file import tim_file,read_txt
from win32com import client
from func.windy_db import laysolieudubao_windy
def mo_word(pth):
    word = client.Dispatch("Word.Application")
    word.Visible = True
    odoc = word.Documents.Open(pth)
def mo_excel():
    pth = tim_file(read_txt('path_tin/DATA_EXCEL.txt'),'.xlsm')
    excel = client.Dispatch("Excel.Application")
    excel.Visible = True
    book = excel.Workbooks.Open(pth)
    book.Worksheets('H').Select()
    
def mo_tvhn():
    mo_word(tim_file(read_txt('path_tin/DRHN.txt'),'.docx'))
def mo_tvhv():
    mo_word(tim_file(read_txt('path_tin/DRHV.txt'),'.docx'))
def mo_tvhd():
    mo_word(tim_file(read_txt('path_tin/TVHD.txt'),'.docx'))
def mo_lulu():
    mo_word(tim_file(read_txt('path_tin/LULU.txt'),'.docx'))
def mo_lqsl():
    mo_word(tim_file(read_txt('path_tin/LQSL.txt'),'.docx'))
def mo_cblu():
    mo_word(tim_file(read_txt('path_tin/CBLU.txt'),'.docx'))
def mo_hhan():
    mo_word(tim_file(read_txt('path_tin/CBLU.txt'),'.docx'))
def tao_btn(ten_btn,ten_txt_bt,cmd,ox,oy):
    ten_btn = Button(root, text=ten_txt_bt,width=10, command=cmd)# bt dong
    ten_btn.place(x=ox,y=oy)
def tao_textbox(ten_txt,ox,oy):
    ten_txt = Text(root,width=5,height=1)
    ten_txt.place(x=ox,y=oy)

root =Tk()
w = 1000
h = 600

ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

root.geometry('%dx%d+%d+%d' % (w, h, x, y))
root.title("Sông Tranh")


# # Tạo combobox
dbv = []
dbv_kt = ['Nguyễn Mạnh Hà', 'Nguyễn Công Tài', 'Bành Thị Ngọc']
dbv_tv = ['Nguyễn Đình Huấn', 'Vũ Văn Tình', 'Lê Thị Thanh Huyền']
for a in dbv_kt:
    for b in dbv_tv:
        dbv.append(a +', ' + b)

combo_box = ttk.Combobox(root, values=dbv,font=('Times New Roman', 13))
combo_box.current(0)  # Chọn vị trí mặc định là 0 (Trương Văn Biên)
combo_box.config(width=30)
# Đặt vị trí (x, y) cho combobox bằng phương thức place()
combo_box.place(x=310, y=83)  

combo_box1 = ttk.Combobox(root, values=['Trương Tuyến','Nguyễn Đình Huấn'],font=('Times New Roman', 13))
combo_box1.current(0)  # Chọn vị trí mặc định là 0 (Trương Văn Biên)
combo_box1.config(width=20)
# Đặt vị trí (x, y) cho combobox bằng phương thức place()
combo_box1.place(x=720, y=83)  



lbl = Label(root,text="ĐÀI KHÍ TƯỢNG THUỶ VĂN KHU VƯC TRUNG TRUNG BỘ" + '\n' + 'ĐÀI KHÍ TƯỢNG THUỶ VĂN TỈNH QUẢNG NAM',font=('Arial Bold',14)).pack(padx=10,pady=15)
lb1 = Label(root,text='STHN',font=('Arial Bold',14)).place(x=90,y=120)
lb2 = Label(root,text='STHV10',font=('Arial Bold',14)).place(x=90+160,y=120)
lb3 = Label(root,text='STHV05',font=('Arial Bold',14)).place(x=90+320,y=120)
lb4 = Label(root,text='STHD',font=('Arial Bold',14)).place(x=90+320+160,y=120)
lb5 = Label(root,text='LULU',font=('Arial Bold',14)).place(x=90+320+160+130,y=120)
# lb6 = Label(root,text='LQSL',font=('Arial Bold',14)).place(x=860,y=120)
# lb7 = Label(root,text='HHAN',font=('Arial Bold',14)).place(x=860+160,y=120)
lb8 = Label(root,text='Dự báo viên:',font=('Arial Bold',14)).place(x=180,y=80)
lb = Label(root,text='Duyệt tin:',font=('Arial Bold',14)).place(x=620,y=80)
# lb1.place(x=90,y=120)


# tao button DRHN
tao_btn('bt_SQL',"Mưa",save_solieu_mua,80,160) # load so lieu
tao_btn('bt_tvhv',"Dự báo Windy",laysolieudubao_windy,80,160+50) #lam tin
tao_btn('bt_tvhv',"Tin nền",tin_tvhn,80,160+100) #lam tin
tao_btn('bt_tvhv',"H_Web",save_solieu_mucnuoc,40,160+150) #lam tin
tao_btn('bt_tvhv',"H_Mail",downloadattmail,125,160+150) #lam tin
tao_btn('bt_tvhd',"Load_tin_TV",tin_tv_load,80,160+200) #lam tin
tao_btn('bt_upload',"Gửi tin",gui_drhn,80,160+250) # ho so du bao
tao_btn('bt_danhgia',"Nghiệm thu",nghiemthu_tvhn,80,160+300) # danh gia
# tao_btn('bt_hoso',"UP_DATABASE",upload_database,80,160+300) # ho so du bao
tao_btn('hs_uploadt',"Mở tin",mo_tvhn,80,160+350) #lam tin

# # tao button drhv5
tao_btn('bt_dlh',"Tin nền KT",tin_nenKT_10day,245,160) # load ho
# tao_btn('bt_dungtich',"Tin nền KT",tin_nenKT_10day,245,160+50) #dung tich ho
# tao_btn('bt_dungtich',"Mở DATA TV",mo_excel,245,160+100) #dung tich ho
# tao_btn('bt_tvhv_map',"Load_TV", tin_tv10_load,245,160+50+100) #ve map
# tao_btn('bt_tvhv',"Gửi tin",gui_drhv,245,160+50+50+100) #lam tin
# tao_btn('bt_tvhv',"HỒ SƠ",hs_tvhv05,245,160+250) #lam tin
# tao_btn('bt_upload10',"Gửi tin",gui_tvhv,245,160+300) # gui tin


selected_value = combo_box.get()
DRHN.set_selected_value(selected_value)
DRHV10.set_selected_value(selected_value)
def update_selected_value(event):
    selected_value = combo_box.get()
    DRHN.set_selected_value(selected_value)
    DRHV10.set_selected_value(selected_value)

# Gắn sự kiện ComboboxSelected với hàm update_selected_value
combo_box.bind("<<ComboboxSelected>>", update_selected_value)

duyettin = combo_box1.get()
DRHN.set_selected_duyet(duyettin)
DRHV10.set_selected_duyet(duyettin)
def update_selected_duyet(event):
    duyettin = combo_box1.get()
    DRHN.set_selected_duyet(duyettin)
    DRHV10.set_selected_duyet(duyettin)
# Gắn sự kiện ComboboxSelected với hàm update_selected_value
combo_box1.bind("<<ComboboxSelected>>", update_selected_duyet)

root.mainloop()