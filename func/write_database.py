import  mysql.connector
from datetime import datetime,timedelta
import pandas as pd
import numpy as np
import pyodbc
def doc_data():
    now = datetime.now()
    df = pd.read_excel(r'\\admin-pc\DATA tin\Dakdrinh\DATA\DR_THUYVAN.xlsx',sheet_name='DRHN')
    now = datetime(now.year,now.month,now.day,7)
    kt = now - timedelta(hours=24)
    df = df.iloc[1:,:21]
    dt_rang = pd.date_range(start=datetime(now.year,8,31,13), periods=len(df['time']), freq="6H")
    df['time'] = dt_rang
    df = df.loc[df['time'] <= now]
    df= df[['time','Htd','qtd','qxa','qcm','qdt']]
    print(df)
    return df
    

# doc_data()
def creat_cxn():
    # Kết nối đến MySQL
    host = '113.160.225.84'
    user = 'qltram'
    password = 'mhq@123456'
    port = 3306
    database = 'datasolieu'
    cnx = mysql.connector.connect(host=host, user=user, password=password, port=port, database=database)
    return cnx

def add_column(table_name, column_name, data_type, max_length):
    # Tạo kết nối
    cnx = creat_cxn()
    
    # Tạo con trỏ
    cursor = cnx.cursor()
    
    try:
        # Tạo câu lệnh SQL để thêm cột vào bảng
        sql = f"ALTER TABLE {table_name} ADD COLUMN {column_name} {data_type}({max_length})"
        
        # Thực hiện truy vấn SQL
        cursor.execute(sql)
        
        # Xác nhận việc thay đổi cấu trúc bảng
        cnx.commit()
        print(f"Cột '{column_name}' đã được thêm vào bảng '{table_name}'.")
        
    except mysql.connector.Error as err:
        # Xử lý lỗi nếu có
        print(f"Lỗi: {err}")
    
    finally:
        # Đóng con trỏ và kết nối
        cursor.close()
        cnx.close()

def delete_all_rows(table_name):
    # Tạo kết nối
    cnx = creat_cxn()
    
    # Tạo con trỏ
    cursor = cnx.cursor()
    
    try:
        # Tạo câu lệnh SQL để xóa tất cả các hàng từ bảng
        sql = f"DELETE FROM {table_name}"
        
        # Thực hiện truy vấn SQL
        cursor.execute(sql)
        
        # Xác nhận việc thay đổi dữ liệu
        cnx.commit()
        print(f"Tất cả các hàng trong bảng '{table_name}' đã được xóa.")
        
    except mysql.connector.Error as err:
        # Xử lý lỗi nếu có
        print(f"Lỗi: {err}")
    
    finally:
        # Đóng con trỏ và kết nối
        cursor.close()
        cnx.close()

def insert_data(df,table_name):
    # Tạo kết nối
    cnx = creat_cxn()
    # Tạo con trỏ
    cursor = cnx.cursor(buffered=True)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    query = f"SELECT * FROM {table_name} LIMIT 1"
    cursor.execute(query)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    column_names = [column[0] for column in cursor.description]
    # print(column_names)
    # SQL truy vấn để chèn dữ liệu
    sql = f"INSERT INTO {table_name} ({', '.join(column_names)}) VALUES ({', '.join(['%s' for _ in column_names])})"
    # print(sql)
    # values = (datetime.now(), "410", '90.3', '0', '68')
    data = df.values.tolist()
    # print(data)
    try:
        # Thực hiện truy vấn SQL
        # cursor.execute(sql, data)
        cursor.executemany(sql, data)
        
        # Xác nhận việc thay đổi dữ liệu
        cnx.commit()
        print("Dữ liệu đã được chèn thành công.")
        
    except mysql.connector.Error as err:
        # Xử lý lỗi nếu có
        print(f"Lỗi: {err}")
    
    finally:
        # Đóng con trỏ và kết nối
        cursor.close()
        cnx.close()
        
def delete_rows_before_date(table_name, bd,kt):
    # Create a connection
    cnx = creat_cxn()
    
    # Create a cursor
    cursor = cnx.cursor()
    
    try:
        # Create an SQL query to delete rows based on the timestamp column
        sql = f"DELETE FROM {table_name} WHERE time >= {bd} and time <= {kt}"
        
        # Execute the SQL query with the date as a parameter
        cursor.execute(sql)
        
        # Confirm the data changes
    except mysql.connector.Error as err:
        # Handle any errors
        print(f"Error: {err}")
    
    finally:
        # Close the cursor and connection
        cursor.close()
        cnx.close()

def update_data(df,table_name):
    # Tạo kết nối
    cnx = creat_cxn()
    # Tạo con trỏ
    cursor = cnx.cursor(buffered=True)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    query = f"SELECT * FROM {table_name} LIMIT 1"
    cursor.execute(query)
    
    # Lấy danh sách các tên cột từ đối tượng con trỏ
    column_names = [column[0] for column in cursor.description]
    # print(column_names)
    # SQL truy vấn để chèn dữ liệu
    sql = f"UPDATE {table_name} SET ({', '.join(column_names)}) VALUES ({', '.join(['%s' for _ in column_names])})"
    # print(sql)
    # values = (datetime.now(), "410", '90.3', '0', '68')
    data = df.values.tolist()
    # print(data)
    try:
        # Thực hiện truy vấn SQL
        # cursor.execute(sql, data)
        cursor.executemany(sql, data)
        
        # Xác nhận việc thay đổi dữ liệu
        cnx.commit()
        print("Dữ liệu đã được chèn thành công.")
        
    except mysql.connector.Error as err:
        # Xử lý lỗi nếu có
        print(f"Lỗi: {err}")
    
    finally:
        # Đóng con trỏ và kết nối
        cursor.close()
        cnx.close()


FileName = ('DATA/QNAM.accdb')
cnxn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + FileName + ';')
query = "SELECT * FROM thuyvan"
df = pd.read_sql(query, cnxn)
print(df)
df = df[['thoigian','mucnuocthuongluu']]


# df = df[['time','Qtranlu']]
# df = df[['time','Qluxacong']]
# df = df[['time','Qlucm']]


# df.insert(0,'Matram','6DR')
# df['sldungduoc'] = df['xa']
# df['maloi'] = np.nan
# df['chinhly'] = np.nan
# df = df[df['time'] <= datetime(2023,11,5,7)]
# df = df.sort_values(by='time')
# df =df.replace(np.nan,None)
# print(df)
# insert_data(df,'ho_dakdrinh_qve')
# delete_all_rows('ho_dakdrinh_qxa')
# add_column('dakdrinh', 'Qdb', 'VARCHAR',25)
# delete_rows_before_date('dakdrinh',datetime(2023,10,5,0).strftime('%Y-%m-%d'),datetime(2023,10,10,0).strftime('%Y-%m-%d'))