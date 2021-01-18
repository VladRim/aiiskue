import win32com.client
import pandas as pd
import datetime
import calendar
import xlwt
from xlwt import *
import pyodbc
import config

Excel = win32com.client.Dispatch("Excel.Application")
# Некоторые констаты
list_params = {2:   "Активная энергия, отдача Значения на интервале, кВт*ч",
               4:   "Активная энергия, прием Значения на интервале, кВт*ч",
               6:   "Реактивная энергия, отдача Значения на интервале, кВар*ч",
               8:   "Реактивная энергия, прием Значения на интервале, кВар*ч",
               23:  "Журнал событий",
               500: "ГВС",
               501: "ХВС"
               }
list_Point_Type = [144,  # Здание
                   10,   # Распределительное устройство
                   7,    # Присоединение
                   21,   # Прибор учета
                   81,   # ТТ
                   85,   # ТН
                   ]

months = {1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}

style_top = XFStyle()
style_top.font.name = 'Arial Cyr'
style_top.font.bold = 2
style_top.alignment.wrap = 1
style_top.alignment.horz = 2
style_top.alignment.vert = 1
style_top.num_format_str = 'M/D/YY'

style_body = XFStyle()
style_top.font.name = 'Arial Cyr'
style_body.num_format_str = '#0.000'

def otch_interval(report_date, tm_zone):
    hour_bg_report = 1  #  часы начала отчетного месяца со сдвигом на московское время
    hour_end_report = 1  #  часы окончания отчетного месяца со сдвигом на московское время
    day_bg_report = 1
    day_end_report = 1


    d = report_date.split(".")
    print(hour_bg_report)
    dt = datetime.datetime(int(d[2]), int(d[1]), int(d[0]))
    year_report = dt.year                    #  отчетный год
    month_report = dt.month - 1              #  отчетный месяц


#  месяц предшествующий отчетному

#  days_report = months[month_report]      #количество дней в отчетном месяце

    if month_report == 0:
        month_report = 12
        month_report_e = 12
        year_report = year_report - 1


    month_report_e = month_report
    year_report_e = year_report

#  days_report = first_day_report.max.day
    first_day_report = dt.replace(day=1, month=month_report)
    days_report = calendar.monthrange(year_report, first_day_report.month)[1]


    year_pastreport = year_report            #  год для месяца следующего за отчетным
    month_pastreport = month_report + 1
    if month_pastreport == 13:
        month_pastreport = 1
        year_pastreport = year_pastreport + 1
    first_day_pastreport = dt.replace(day=1, month=month_pastreport)
#

#  days_prereport = months[month_prereport]    #  количество дней в месяце предшествующем отчетному
    count_halfhours = days_report * 48      #  количество получасовок в отчетном месяце
    count_hours = count_halfhours / 2    #  количнество часов в отчетном месяце


    minuts_bg = 00


    data_begin_report = datetime.datetime(year_report, month_report, day_bg_report, hour_bg_report, 00, 00)
    data_end_report = datetime.datetime(year_pastreport, month_pastreport, day_end_report, hour_end_report, 00, 00)

    return year_report, month_report, hour_bg_report, hour_end_report, data_begin_report, data_end_report, days_report, count_halfhours,dt


def request_ID_ontype(ID_P):
    #index = []
    cursor.execute("""select ID_Point,
                             Point_Type,
                             PointName
                      from   Points
                      where  ID_Parent = ?""", ID_P)
    rows = tuple(cursor.fetchall())
    ln = len(rows)
    for c in range(ln):
        P_Type = rows[c][1]
        if P_Type != 21:
            #if P_Type == 10:
            #str = str + 1
            #print(rows[c][2])
            request_ID_ontype(rows[c][0])

        #if P_Type == 7:
        #    ln = ln + len(rows)
        else:
            print(rows[c][2])
            ID_PP = request_bd_ID_PP(rows[c][0], 4)


            index.append(request_volume(ID_PP, data_begin_report, data_end_report, rows[c][2], 4))


    return rows


def request_bd_ID_PP(ID_obj, Type_param):
    cursor.execute('select ID_PP from PointParams where ID_Point = ? and ID_Param = ?', ID_obj, Type_param)
    try:
        row = cursor.fetchone()
        return row[0]
    except TypeError:
        print('параметр отсутствует')
        #row[0] = 'none'
        return "none"


def request_volume(ID_PP, data_begin_report, data_end_report, Points_prisoed_Name, Type_param):
    if ID_PP != "none":
        #ws.col(0).width = 3000
        #ws.write(j, 0, Points_prisoed_Name)
        #ws.write(j, column, list_params[Type_param])
        cursor.execute('select val from pointNIs_On_Main_Stack where ID_PP = ? and DT = ?', ID_PP, data_begin_report)
        try:
            row = cursor.fetchall()
            vall_beg = row[0][0]
        #ws.write(str, 4, vall[0][0)
        except:
            print("Для адреса нет данных на отчетном периоде")
            vall_beg = "no_data"
        cursor.execute('select val from pointNIs_On_Main_Stack where ID_PP = ? and DT = ? ', ID_PP, data_end_report)
        try:
            row = cursor.fetchall()
            vall_end = row[0][0]
        #ws.write(str, 5, vall[0])
            #j = j + 1
        except:
            print("Для адреса нет данных на отчетном периоде")
            vall_end = "no_data"
    else:
         #ws.col(0).width = 3000
         #ws.write(j, 0, Points_prisoed_Name)
         #ws.write(j, 4, list_params[Type_param])
         vall_beg = vall_end = "no_data"
         #ws.write(str, 4, vall)
         # = j + 1

    #str = str + 1
    return Points_prisoed_Name, vall_beg, vall_end


# Формирование строки SQL
driver = 'DRIVER={SQL Server}'
server = 'SERVER=172.16.1.4'
port = 'PORT=1433'
db = 'DATABASE=SB_ASKUE'
user = 'UID=sa'
pw = 'PWD=Telecor13'
conn_str = ';'.join([driver, server, port, db, user, pw])

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

df = pd.read_csv('list_report.csv', encoding='windows-1251', sep=';')
adr = list(df['Адрес'])
lenth = len(adr)
city = list(df['Город'])
directory = list(df['Каталог'])
report_period = list(df['Отчетный месяц'])
time_zone = list(df['Тайм зона'])
level_deep = list(df['Уровень'])
A_send = list(df['Активная отдача'])
A_received = list(df['Активная прием'])
R_send = list(df['Реактивная отдача'])
R_received = list(df['Реактивная прием'])
Integral_A_received = list(df['Показания'])



for i in range(0, lenth):
    print(adr[i])
    index = []
# определяемся с временнЫми переменными
    year_report, month_report, hour_bg_report, hour_end_report, data_begin_report, data_end_report, days_report, count_halthours,dt = otch_interval(report_period[i], time_zone[i])
    wb = xlwt.Workbook()
    ws = wb.add_sheet(str(month_report) + '-' + str(year_report))
    ws.col(0).width = 9000
    ws.col(3).width = 9000
    ws.col(4).width = 3000
    ws.col(5).width = 3000

# запись временнЫх интервалов в таблицу
    #table_interval_wr(year_report, month_report, days_report)

# запишем в таблицу адрес объекта
    p = 1
    P_Name = adr[i]
    ws.write(1, 0, P_Name)

# Зададим имя файла и папку
    filename = P_Name
    fl = filename.replace("/", "-")
    fl = "Показания" + '\\' + city[i] + ", " + fl + ".xls"


 #запрос к БД
    cursor.execute('select ID_Point, Point_Type, PointName from Points where  PointName = ?',
                   P_Name)  # забираем из БД тип точки учета
    Points_addr = cursor.fetchone()
    ID_P = Points_addr[0]
    P_Type = Points_addr[1]                                                                       # идетификатор точки
    P_Name = str(Points_addr[2])

    print("ID_Point: ", ID_P, "ID_Type: ", P_Type, "Point_Name: ", P_Name)
    P_Name_all = P_Name
    #ws.write(1, 0, P_Name, style_top)

    #str = 1
    request_ID_ontype(ID_P)  # тип подточки
    ws.write(2, 4, data_begin_report, style_top)
    ws.write(2, 5, data_end_report, style_top)
    for j in range(len(index)):
        i = 2
        for ln in index[j]:
            i = i + 1
            #ws.write(j + 3, i, ln, style_body)
            print(ln)
            #ws.write(j, i + 3, ln[0])


    wb.save(fl)
