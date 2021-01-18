import pandas as pd
import pyodbc
import datetime
import calendar
import config
import data_base.init_bd_mssql as mssql
cursor, conn = mssql.init_mssql()


def read_integral(ID_PP, Date_req):
    cursor.execute('select val from pointNIs_On_Main_Stack where ID_PP = ? and DT = ?', ID_PP, Date_req)
    try:
        row = cursor.fetchall()
        vall = row[0][0]
    except:
        print("Для адреса нет данных на отчетном периоде")
        vall = "no_data"
    return vall


def date_db(Date_req, Time_req):
    mm_fl = int(Date_req[3:5])
    mm_bd = mm_fl
    dd_fl = int(Date_req[0:2])
    dd_bd = dd_fl
    gg_fl = int(Date_req[6:])
    gg_bd = gg_fl
    days = calendar.monthrange(gg_fl, mm_fl)
    hh_fl = int(Time_req[0:2])
    min_fl = int(Time_req[3:])
    # костыль
    if hh_fl == 0 and min_fl == 0:
        dd_bd -= 1
    hh_bd = hh_fl + 1
    tm_fl = datetime.datetime(gg_fl, mm_fl, dd_fl, hh_fl, min_fl, 0)
    print("дата, время файл ", tm_fl)
    if hh_bd == 24 or (hh_bd == 1 and min_fl == 0):
        dd_bd += 1
        if dd_bd > days[
            1]:  # если переходим границу месяца, увеличиваем текущий месяц на 1 и начинаем отсчет дней с единицы
            dd_bd = 1
            mm_bd = mm_fl + 1
            if mm_bd == 13:  # если переходим границу года, увеличиваем текущий год на 1 и начинаем отсчет месяцев с единицы
                mm_bd = 1
                gg_bd = gg_fl + 1
    if hh_bd == 24:
        hh_bd = 0

    if min_fl == 30:
        min_bd = 0

    else:
        min_bd = 30
        if hh_bd == 0:
            hh_bd = 23
            dd_bd -= 1
            if dd_bd == 0:
                dd_bd = days[1]
                mm_bd = int(Date_p[i][3:5])
                gg_bd = int(Date_p[i][6:])
        else:
            hh_bd -= 1

    tm_db = datetime.datetime(gg_bd, mm_bd, dd_bd, hh_bd, min_bd, 0)
    return tm_db


df_pu = pd.read_csv(config.FILE_PU, encoding='windows-1251', sep=';', error_bad_lines=False)
print(df_pu)
sum_a = 0
sum_r = 0
integral_a_e = []
integral_r_e = []

for cnt in range(0, len(df_pu.action)):
    if df_pu.action[cnt] == 1:
        sn_pu = str(df_pu.PU[cnt])
        cursor.execute('SELECT ID_MeterInfo FROM MeterInfo WHERE SN = ?', sn_pu)
        ID_Meter = cursor.fetchone()
        cursor.execute('SELECT ID_Point FROM MeterMountHist WHERE ID_MeterInfo = ?', ID_Meter[0])
        ID_Point = cursor.fetchone()
        cursor.execute('SELECT ID_PP FROM PointParams WHERE ID_Point = ? and ID_Param = ?', ID_Point[0], config.ACTIVE)
        ID_PP_activ = cursor.fetchone()
        cursor.execute('SELECT ID_PP FROM PointParams WHERE ID_Point = ? and ID_Param = ?', ID_Point[0],
                       config.REACTIVE)
        ID_PP_reactiv = cursor.fetchone()
        print("ID_PP_activ:", ID_PP_activ[0], " ID_PP_reactiv:", ID_PP_reactiv[0])

        df = pd.read_csv(sn_pu + '.csv', encoding='windows-1251', sep=';', error_bad_lines=False)
        Time_p = list(df['Time'])  # время из файла
        Date_p = list(df['Date'])  # дата из файла

        tm_bd = date_db(Date_p[0], Time_p[0])
        # показания на начало периода
        integral_a_b = read_integral(ID_PP_activ[0], tm_bd), tm_bd
        integral_r_b = read_integral(ID_PP_reactiv[0], tm_bd), tm_bd
        print("показания начало активка ", integral_a_b, "показания начало реактивка ", integral_r_b)

        A_received = list(df['P+'])
        A_received_p = []
        count = 0
        for i in range(0, len(A_received)):
            count += 1
            Val = float(A_received[i].replace(",", "."))
            print(f"{sn_pu},{count}, {Val}")
            A_received_p.append(Val)
            sum_a += Val
        sum_a /= 2
        R_received = list(df['R+'])
        R_received_p = []
        for i in range(0, len(R_received)):
            Val = float(R_received[i].replace(",", "."))
            R_received_p.append(Val)
            sum_r += Val
        sum_r /= 2
        # показания на конец периода
        try:
            integral_a_e = integral_a_b[0] + sum_a
            ntegral_r_e = sum_r
            #integral_r_e = integral_r_b[0] + sum_r
        except TypeError:
            print('Начальные показания отсутствуют: ', integral_a_b[0])

        print("дата ", tm_bd, "показания конец активка ", integral_a_b, "показания конец реактивка ", integral_r_b)
        for i in range(0, len(A_received)):
            P = float(A_received_p[i]) / 2  # делим на постоянную счетчика, если в файле данные из профиля счетчика
            R = float(R_received_p[i]) / 2  # делим на постоянную счетчика, если в файле данные из профиля счетчика
            tm_bd = date_db(Date_p[i], Time_p[i])
            print("дата, время база ", tm_bd, "значение P", P)
            print("дата, время база ", tm_bd, "значение R", R)
            try:
                cursor.execute("INSERT INTO PointMains (ID_PP, DT, Val, State) VALUES(?,?,?,?)", ID_PP_activ[0], tm_bd, P, 0)  # активка на прием
                cursor.execute("INSERT INTO PointMains (ID_PP, DT, Val, State) VALUES(?,?,?,?)", ID_PP_reactiv[0], tm_bd, R, 0)  # реактивка на прием

            # если данные уже есть, заменить на данные из файла
            except pyodbc.IntegrityError:
                cursor.execute("UPDATE PointMains SET Val=? WHERE ID_PP = ? and DT = ?", P, ID_PP_activ[0], tm_bd)  # активка на прием
                #cursor.execute("delete from PointMains WHERE ID_PP = ? and DT = ?", ID_PP_reactiv[0], tm_bd)  # реактивка на прием

if count % 48 == 0:# если кратно 48, то получаем целое количество дней, если количество дней не целое, показания не фиксируем
    hr = 1
    min = 00
    tm_end = datetime.datetime(tm_bd.year, tm_bd.month, tm_bd.day, hr, min, 0)

    # запишем показания в БД
    id_pp_a = ID_PP_activ[0]
    id_pp_r = ID_PP_reactiv[0]

    val_a = integral_a_e
    val_r = integral_r_e

    try:
        cursor.execute("INSERT INTO pointNIs_On_Main_Stack (ID_PP, DT, Val, State, IsExactDate) VALUES(?,?,?,?,?)",
                       id_pp_a, tm_end, val_a, 0, 1)
        cursor.execute("INSERT INTO pointNIs_On_Main_Stack (ID_PP, DT, Val, State, IsExactDate) VALUES(?,?,?,?,?)",
                       id_pp_r, tm_end, val_r, 0, 1)
    except pyodbc.IntegrityError:
        cursor.execute("UPDATE pointNIs_On_Main_Stack SET Val=?  WHERE ID_PP = ? and DT = ?",
                       val_a, id_pp_a, tm_end)
        cursor.execute("UPDATE pointNIs_On_Main_Stack SET Val=?  WHERE ID_PP = ? and DT = ?",
                       val_r, id_pp_r, tm_end)
    except pyodbc.Error:
        print('Начальные показания отсутствуют')

else:
    print("профиль неполный")


print("end")
conn.commit()

cursor.close()
conn.close()
