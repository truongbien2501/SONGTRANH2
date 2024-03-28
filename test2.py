import pandas as pd
from datetime import datetime,timedelta
import pyodbc
mua_db = pd.read_excel('DATA/windy_db.xlsx')
mua_db = mua_db[['time','muaTra Doc','muaTrung Luu','muaThuong Luu']]
mua_db.set_index('time',inplace=True)
print(mua_db.sum())