import win32com.client
import pandas as pd
import datetime
import calendar
import xlwt
from xlwt import *
import pyodbc
import config
import data_base.init_bd_mssql as mssql


def request_bd_id_pp(ID_obj, Type_param):
    cursor.execute("""select ID_PP
                      from PointParams
                      where ID_Point = ? and ID_Param = ?""", ID_obj, Type_param)

    try:
        row = cursor.fetchone()
        return row[0]
    except TypeError:
        print('параметр отсутствует')
# row[0] = 'none'
        return "none"

cursor, conn = mssql.init_mssql()
wb = xlwt.Workbook()
cursor.execute('select PointName, ID_Point from Points where  Point_Type = 144')
P_Name = cursor.fetchall()
for cnt in range(0, len(P_Name) - 1):
    try:
        print(cnt, ' ', P_Name[cnt][1], ' ', P_Name[cnt][0])
        ID_PP = request_bd_id_pp(P_Name[cnt][1], 4)
        cursor.execute('select Val, DT from PointMains where ID_PP = ? order by DT', ID_PP)
        Val_DT = cursor.fetchall()
        for count in range(0, len(Val_DT)):
            DT = Val_DT[count][1]
            DT_next = Val_DT[count + 1][1]
            print(DT, '-', DT_next, " ", Val_DT[count][0])
        #print(Val_DT)
    except:
        print(DT, '-', DT_next, " ", Val_DT[cnt][0])

