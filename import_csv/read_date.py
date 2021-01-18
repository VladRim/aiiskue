import win32com.client
import pandas as pd
# import openpyxl
# import datetime
import pyodbc
import datetime
# import calendar
import config

# запись данных о поверках в БД по серийному номеру



# Формирование строки SQL
driver = config.DRIVER
server = config.SERVER_ADDR
user = config.UNAME
pw = config.PASS
port = config.PORT
db = config.DATABASE


conn_str = ';'.join([driver, server, port, db, user, pw])

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()


date_N = []
date_Pn = []
sn_N = []
sn_Nt = []
dt_Nt = []
int_L = []

df = pd.read_csv('poverka_tt.csv', encoding='windows-1251', sep=';', error_bad_lines=False)
sn = list(df['N'])
date_P = list(df['DT'])
interval = list(df['Interval'])


for i in range(len(sn)):
    nn = str(sn[i])
    #print(sn[i])
    nn_L = nn.split()
    try:
        dt = date_P[i]
    except:
        dt = date_P[0]
    dt_L = dt.split()
    for j in range(len(nn_L)):
        sn_Nt.append(nn_L[j])
        int_L.append(interval[i])
        try:
            dt_Nt.append(dt_L[j])
        except:
            dt_Nt.append(dt_L[0])


for i in range(len(sn_Nt)):
        dd = int(dt_Nt[i][0:2])
        mm = int(dt_Nt[i][3:5])
        gg_P = int(dt_Nt[i][6:])
        gg_N = gg_P + int(int_L[i])
        dt_N = datetime.datetime(gg_N, mm, dd, 0, 0, 0)
        dt_P = datetime.datetime(gg_P, mm, dd, 0, 0, 0)
        date_N.append(dt_N)
        date_Pn.append(dt_P)
        #print(sn_Nt[i])
        #print(date_Pn[i],date_N[i],i)

for i in range(len(sn_Nt)):
    print(sn_Nt[i])
    print(date_Pn[i], date_N[i], i)
    cursor.execute("UPDATE [SB_ASKUE].[dbo].[MeterInfo] SET DT_QC_Prev=?, DT_QC_Next=?  WHERE SN = ?", date_Pn[i], date_N[i], sn_Nt[i])



#print(date_P)
#print(date_N)

conn.commit()