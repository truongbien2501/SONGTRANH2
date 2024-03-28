from ftplib import FTP
from tkinter import messagebox,Tk
def upload_file(file_path, ftp_url, username, password):
    try:
        # Tách thành phần từ URL FTP
        url_parts = ftp_url.split("/")
        ftp_server = url_parts[2]
        remote_path = "/".join(url_parts[3:]) + "/" + file_path.split('\\')[-1]
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