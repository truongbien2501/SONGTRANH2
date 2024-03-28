import pandas as pd
from datetime import datetime
import  mysql.connector
# df = pd.read_excel(r'D:\PM_PYTHON\SONGTRANH\BANTIN\H,Q ho2023.xls',sheet_name='Thang9')
# df.columns = df.loc[0]
# df = df.iloc[1:25,:33*4+1]
# # print(df)
# # Chia thành các DataFrame 5 cột
# df_list = [df.iloc[:, i:i+4] for i in range(1, len(df.columns), 4)]
# # print(df_list)

# # for dta in df_list:
#
# ';lkb ưtyu#     print(dta)
# # Ghép các DataFrame lại theo chiều dọc 
# df_concat = pd.concat(df_list, axis=0)
# bd = datetime(2023,8,30)
# df_concat.insert(0,'time',pd.date_range(bd,periods=df_concat.shape[0],freq='H'))
# df_concat.to_excel('BANTIN/Songtranh9.xlsx')
# print(df_concat)

# df = pd.read_excel('BANTIN/Songtranh9.xlsx')
# df1 = pd.read_excel('BANTIN/Songtranh10.xlsx')
# df2 = pd.read_excel('BANTIN/Songtranh11.xlsx')
# df = pd.concat([df,df1,df2],axis=0)
# df = df.iloc[:,1:]
# df.to_excel('tonghop.xlsx',index=False)
# print(df)

def creat_cxn():
    # Kết nối đến MySQL
    host = '113.160.225.84'
    user = 'qltram'
    password = 'mhq@123456'
    port = 3306
    database = 'datasolieu'
    cnx = mysql.connector.connect(host=host, user=user, password=password, port=port, database=database)
    return cnx

def query_sql(list_import,table_clounms,table_name):#TidVerticalIDVelocityForDetailMeasurement 
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


def insert_data(df,table_name):
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
        sql = query_sql(data,column_names,table_name)
        print(sql)
        try:
            cursor.execute(sql)
            cnx.commit()
        except:
            pass
    cursor.close()
    cnx.close()

from tkinter import messagebox
from datetime import datetime,timedelta

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

def solieu_web():
    df = pd.read_excel(r'BANTIN\solieu.xlsx')
    df = df.iloc[4:,:]
    print(df.info())
    df['Ngày đo'] = df['Ngày đo'].astype(str)
    df['Giờ đo'] = df['Giờ đo'].astype(str)
    df['time'] = df['Ngày đo'] + ' ' + df['Giờ đo']
    print(df.columns)
    df.rename(columns={'Mực nước thượng lưu':'H hồ','Lưu lượng đến hồ':'Qđến','Tổng lưu lượng xả':'Qtongxa'},inplace=True)
    print(df)
    return df
def updatedatabase():
    # df = pd.read_excel('BANTIN/tonghop.xlsx')
    # df = solieusongtranh()
    df = solieu_web()
    df =df.loc[~df['time'].duplicated(keep='last')]
    # print(df['Qxả'].dtypes)
    # print(df['Qmáy'].dtypes)
    if 'Qtongxa' not in df.columns:
        df['Qxả'] = df['Qxả'].astype(str)
        df['Qxả'] = df['Qxả'].str.replace(',','.')
        df['Qmáy'] = df['Qmáy'].astype(str)
        df['Qmáy'] = df['Qmáy'].str.replace(',','.')
        # print(df)
        df['Qtongxa'] = df['Qmáy'].astype(float) + df['Qxả'].astype(float)
        print(df)
    insert_data(df[['time','Qđến']],'ho_dakdrinh_qve')
    insert_data(df[['time','Qtongxa']],'ho_dakdrinh_qdieutiet')
    insert_data(df[['time','H hồ']],'ho_dakdrinh_mucnuoc')
    messagebox.showinfo('Thông báo!','OK')
updatedatabase()