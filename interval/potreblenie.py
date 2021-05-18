import win32com.client
import pandas as pd
import datetime
import calendar
from datetime import tzinfo
import xlwt
from interval.style_xls import *
from data_base import init_bd_mssql
import dt_interval
import to_TZ_BD

Excel = win32com.client.Dispatch("Excel.Application")


# Некоторые констаты
list_params = {2:   "Активная энергия, отдача Значения на интервале, кВт*ч",
               4:   "Активная энергия, прием Значения на интервале, кВт*ч",
               6:   "Реактивная энергия, отдача Значения на интервале, кВар*ч",
               8:   "Реактивная энергия, прием Значения на интервале, кВар*ч",
               10:  "Сальдо переток, активная энергия, кВт*ч",
               23:  "Журнал событий"
               }
list_Point_Type = [144,  # Здание
                   10,   # Распределительное устройство
                   7,    # Присоединение
                   21    # Прибор учета
                   ]


def otch_interval(report_date, tm_zone_obj):
    utc_t = 0
    tm_zone_bd = 4

    dt_bg = to_TZ_BD.convert_TZ(report_date, tm_zone_obj)
    dt_day = dt_bg.day
    dt_month = dt_bg.month
    dt_year = dt_bg.year
    month_report = dt_bg.month
    month_past_report = month_report + 1
    if month_past_report > 12:
     month_past_report = 1
#  часы начала отчетного месяца со сдвигом на московское время
    hour_bg_report = dt_bg.hour
    if hour_bg_report < 0:

        hour_bg_report += 24
        dt_day -= 1
        dt_month -= 1
        if dt_month == 0:
            dt_month = 12
            dt_year -= 1
        if dt_day == 0:
            dt_day = calendar.monthrange(dt_year, dt_month)[1]
#  часы окончания отчетного месяца со сдвигом на московское время
    day_end_report = calendar.monthrange(dt_year, dt_month)[1]
    hour_end_report = hour_bg_report  # tm_zone_bd - tm_zone_obj
    day_bg_report = dt_day


    #d = report_date.split(".")
    print(hour_bg_report)
    #dt = datetime.datetime(int(d[2]), int(d[1]), int(d[0]))
    dt = report_date
#  отчетный год
    year_report = dt_year
#  отчетный месяц
    month_report = dt_month  # - 1

#  месяц предшествующий отчетному
#  days_report = months[month_report]      # количество дней в отчетном месяце

    if month_report == 0:
        month_report = 12
        month_report_e = 12
        year_report = year_report - 1

    month_report_e = month_report
    year_report_e = year_report

#  days_report = first_day_report.max.day
    #first_day_report = dt.replace(day=1, month=month_report)
    first_day_report = report_date.day
    days_report = calendar.monthrange(year_report, report_date.month)[1]
#  год для месяца следующего за отчетным
    year_pastreport = year_report
    month_pastreport = month_report + 1
    if month_pastreport == 13:
        month_pastreport = 1
        year_pastreport = year_pastreport + 1
    first_day_pastreport = dt.replace(day=1, month=month_pastreport).day

#  количество получасовок в отчетном месяце
    count_halfhours = days_report * 48

#    minuts_bg = 00

    data_begin_report = datetime.datetime(year_report, month_report, day_bg_report, hour_bg_report, 00, 00)
    data_end_report = datetime.datetime(year_pastreport, month_pastreport, first_day_pastreport, hour_end_report, 00, 00)

    return year_report, month_report, hour_bg_report, hour_end_report, data_begin_report, data_end_report, days_report,\
        count_halfhours, dt


# Процедура отрисовки таблички показаний
#def table_integrals(row, ID_P, beg_interval, end_interval):
#    ws.write(row + 2, 0, "Наименование точки учёта", style_body)
#    ws.write(row + 2, 1, "№ счетчика", style_body)
#    ws.write(row + 2, 2, "Расч. коэф.", style_body)
#    ws.write(row + 2, 3, beg_interval.strftime("%d.%m.%Y"), style_body)
#    ws.write(row + 2, 4, end_interval.strftime("%d.%m.%Y"), style_body)


# Процедура записи временных интервалов в таблицу
def table_interval_wr(year_report, month_report, days_report):
    list_ints =[]
    year_report_end = year_report
    month_report_end = month_report
    row_sheet = 2
    d_bg = 1
    d_end = 1
    for k in range(1, days_report + 1):
        for t_hour_bg in range(0, 24):
            t_hour_end = t_hour_bg + 1
            if t_hour_end == 24:
                if d_end == days_report:
                    month_report_end = month_report + 1
                    if month_report_end == 13:
                        month_report_end = 1
                        year_report_end = year_report + 1
                d_end = d_bg + 1
                if d_end > days_report:
                    d_end = 1
                t_hour_end = 0
            interval_bg = datetime.datetime(year_report, month_report, d_bg, t_hour_bg, 00).strftime("%d.%m.%Y"
                                                                                                     " %H:%M:%S")
            interval_end = datetime.datetime(year_report_end, month_report_end, d_end, t_hour_end, 00).strftime(
                "%d.%m.%Y %H:%M:%S")
            interval = interval_bg + ' - ' + interval_end
            list_ints.append(datetime.datetime(year_report, month_report, d_bg, t_hour_bg, 00).strftime("%Y-%m-%d"
                                                                                                     " %H:%M:%S"))
            ws.write(row_sheet, 0, interval, style_body)
            #  print(interval)
            row_sheet = row_sheet + 1
        d_bg = d_bg + 1
    ws.write(row_sheet, 0, "Сумма", style_bottom)
    row_i = row_sheet + 2
    return list_ints


# шапка таблички показаний
    #ws.write(row_sheet + 2, 0, "Наименование точки учёта", style_body)
    #ws.write(row_sheet + 2, 1, "№ счетчика", style_body)
    #ws.write(row_sheet + 2, 2, "Расч. коэф.", style_body)
    #ws.write(row_sheet + 2, 3, data_begin_report.strftime("%d.%m.%Y"), style_body)
    #ws.write(row_sheet + 2, 4, data_end_report.strftime("%d.%m.%Y"), style_body)
    #Name_pu, val_1, val_2 = request_volume_integral()
    #ws.write(row + 3, 0, Name_pu)
    #ws.write(row + 3, 3, val_1)
    #ws.write(row + 3, 4, val_2)

# запрос потомков
def request_id_descendants(ID_P):
    cursor_ms.execute("""select ID_Point,
                             Point_Type,
                             PointName
                      from   Points
                      where  ID_Parent = ?""", ID_P)
    descendants = tuple(cursor_ms.fetchall())               # потомки

    return descendants


def request_bd_id_pp(ID_obj, Type_param):
    cursor_ms.execute("""select ID_PP
                      from PointParams
                      where ID_Point = ? and ID_Param = ?""", ID_obj, Type_param)

    try:
        row = cursor_ms.fetchone()
        return row[0]
    except TypeError:
        print('параметр отсутствует')
# row[0] = 'none'
        return "none"


def request_volume_interval(ID_PP, data_begin_report, data_end_report, column, Points_prisoed_Name, Type_param, count_halfhours, list_ints):
    #dct_ints = {dct_ints: it for it, dct_ints in enumerate(list_ints)}
    #dct_ints = {v: k for k, v in dct_ints .items()}
    j = 1
    summ = 0
    bg_ls = 0
    indx = 1
    date_time = ""
    ws.col(column).width = 3000
    ws.write(j - 1, column, Points_prisoed_Name, style_top)
    ws.write(j, column, list_params[Type_param], style_top)
    if ID_PP != "none":
        try:
            cursor_ms.execute('select val, DT from PointMains where ID_PP = ? and DT >= ? and DT < ?', ID_PP, data_begin_report,
                           data_end_report)
            row = cursor_ms.fetchall()
            len_int = len(row)

            date_time = row[0][1]
            date_time_str = str(date_time)
            #indx = list_ints.index('2018-12-31 01:30:00')
            if date_time.minute == 30:
                vl_h = row[0][0]
                summ += vl_h
                bg_ls = 1
            # noinspection NonAsciiCharacters
            for cnt in range(bg_ls, len(row), 2):
                try:
                    vall = row[cnt][0] + row[cnt + 1][0]
                    date_time = row[cnt][1]

                    indx = list_ints.index(str(date_time))
                    print(indx)
                    indx += 1
                    summ += vall
                    ws.write(indx, column, vall, style_body)
                    print(indx, " ", date_time, "", "ID_PP", ID_PP, vall)
                except IndexError:

                    print(j, "", date_time, "", "Для адреса нет данных на отчетном периоде")
                except ValueError:
                    indx += 1
                    #indx += 1
                    summ += vall
                    ws.write(indx, column, vall, style_body)
                    print(indx, " ", date_time, "", "ID_PP", ID_PP, vall)

        except IndexError:
            print("Нет данных")
    else:
        for i in range(0, count_halfhours, 2):
            vall = 'none'
            ws.write(j, column, vall, style_body)
            j += 1
    ws.write(indx + 1, column, summ, style_bottom)

    wb.save(fl)
    return date_time


df = pd.read_csv('C:\\Users\\admsys.SB-AIISKUE1.000\\aiiskue_py\\csv\\НН Екатеринбург формулы.csv', encoding='windows-1251', sep=';', error_bad_lines=False)
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
Saldo = list(df['Сальдо'])
N_interval = list(df['Отчетный период начало'])
E_interval = list(df['Отчетный период конец'])
# Integral_A_received = list(df['Показания'])


cursor_ms, conn = init_bd_mssql.init_mssql()
now = datetime.datetime.now()
list_date = []
list_date.extend(dt_interval.read_fl())
  # tm_zone_bd - tm_zone_obj
# if hour_bg_report < 0:
#
#        hour_bg_report += 24
#        dt_day -= 1
#        dt_month -= 1
#        if dt_month == 0:
#            dt_month = 12
#            dt_year -= 1
#        if dt_day == 0:
#            dt_day = calendar.monthrange(dt_year, dt_month)[1]


# цикл по всем адресам из таблицы
for i in range(0, lenth):                     # цикл по адресам из файла
    #hour_bg_report = time_zones.get(time_zone[i]) # tm_zone_bd - tm_zone_obj
    wb = xlwt.Workbook()
    index = []
    #list_date = []
    print(adr[i])


    for cnt in range(0, len(list_date)):
        year_report = list_date[cnt].year
        month_report = list_date[cnt].month
        days_report = calendar.monthrange(year_report, month_report)[1]
        ws = wb.add_sheet(str(month_report) + '-' + str(year_report))
        ws.col(0).width = 10500

        # запись временнЫх интервалов в таблицу
        list_ints = table_interval_wr(year_report, month_report, days_report)
        # запишем в таблицу адрес объекта
        colmn = 1
        P_Name = adr[i]
        ws.write(1, 0, P_Name, style_top)

        # определяемся с временнЫми переменными
        year_report, month_report, hour_bg_report, hour_end_report, data_begin_report, data_end_report, days_report,\
            count_halthours, dt = otch_interval(list_date[cnt], time_zone[i])





    # Зададим имя файла и папку
        date_name = now.strftime("%d-%m-%Y")
        filename = P_Name
        fl = filename.replace("/", "-")
        fl = 'C:\\Users\\admsys.SB-AIISKUE1.000\\aiiskue_py\\xls\\' + directory[i] + '\\' + fl + "_" + date_name + ".xls"
    # fl = city[i] + ", " + fl + ".xls"

    # запрос к БД, забираем из БД тип точки учета
        cursor_ms.execute('select ID_Point, Point_Type, PointName from Points where PointName = ?', P_Name)
        Points_addr = cursor_ms.fetchone()
        ID_P = Points_addr[0]
        P_Type = Points_addr[1]
        P_Name = Points_addr[2]
        print("ID_Point: ", ID_P, "ID_Type: ", P_Type, "Point_Name: ", P_Name)
        if level_deep[i] == 144:  # уровень Здание
            if A_send[i] == 1:
                Type_param = 2
                ID_PP = request_bd_id_pp(ID_P, Type_param)
                if ID_PP != "none":
                    request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, P_Name, Type_param, count_halthours, list_ints)
                    colmn += 1
            if A_received[i] == 1:
                Type_param = 4
                ID_PP = request_bd_id_pp(ID_P, Type_param)
                if ID_PP != "none":
                    request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, P_Name, Type_param,
                                            count_halthours, list_ints)
                    colmn += 1
            if R_send[i] == 1:
                Type_param = 6
                ID_PP = request_bd_id_pp(ID_P, Type_param)
                if ID_PP != "none":
                    request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, P_Name, Type_param,
                                            count_halthours, list_ints)
                    colmn += 1
            if R_received[i] == 1:
                Type_param = 8
                ID_PP = request_bd_id_pp(ID_P, Type_param)
                if ID_PP != "none":
                    request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, P_Name, Type_param,
                                            count_halthours, list_ints)
                    colmn += 1

        else:
            Points_Type = request_id_descendants(ID_P)  # тип подточки
            for row in Points_Type:
                print(row[1])  # если тип точки - рапределительное устройство, то:
                if row[1] == 7:  # уровень Ввод/Присоединение
                    ID_P_ru = row[0]
                    P_Type_ru = row[1]
                    P_Name_ru = row[2]

                    print("ID_Point распредустройства:", ID_P_ru, "ID_Type рапредустройства", P_Type_ru,
                          "Point_Name рапредустройства: ", P_Name_ru)
                    Points_prisoed = request_id_descendants(ID_P_ru)
                    for pr in Points_prisoed:
                        ID_Points_prisoed = pr[0]
                        Points_prisoed_Type = pr[1]
                        Points_prisoed_Name = pr[2]

                        print("ID_Point присоединения: ", ID_Points_prisoed, "ID_Type присоединения: ", Points_prisoed_Type,
                              "Point_Name присоедиения: ", Points_prisoed_Name)
                        if A_send[i] == 1:
                            Type_param = 2
                            ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                            if ID_PP != "none":
                                request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name,
                                                        Type_param, count_halthours, list_ints)
                                colmn += 1

                        if A_received[i] == 1:
                            Type_param = 4
                            ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                            request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                    count_halthours, list_ints)
                            colmn += 1

                        if R_send[i] == 1:
                            Type_param = 6
                            ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                            if ID_PP != "none":
                                request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name,
                                                        Type_param, count_halthours, list_ints)
                                colmn += 1

                        if R_received[i] == 1:
                            Type_param = 8
                            ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                            request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                    count_halthours, list_ints)
                            colmn += 1



                else:
                    ID_Points_prisoed = row[0]
                    Points_prisoed_Type = row[1]
                    Points_prisoed_Name = row[2]

                    print("ID_Point присоединения: ", ID_Points_prisoed, "ID_Type присоединения: ", Points_prisoed_Type,
                          "Point_Name присоедиения: ", Points_prisoed_Name)

                    if A_send[i] == 1:
                        Type_param = 2
                        ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                        if ID_PP != "none":
                            request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                    count_halthours, list_ints)
                            colmn += 1
                    if A_received[i] == 1:
                        Type_param = 4
                        ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                        request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                count_halthours, list_ints)
                        colmn += 1
                    if R_send[i] == 1:
                        Type_param = 6
                        ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                        if ID_PP != "none":
                            request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                    count_halthours, list_ints)
                            colmn += 1

                    if R_received[i] == 1:
                        Type_param = 8
                        ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                        request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name, Type_param,
                                                count_halthours, list_ints)
                        colmn += 1

        if Saldo[i] == 1:
            Type_param = 10
            ID_PP = request_bd_id_pp(ID_P, Type_param)
            if ID_PP != "none":
                request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, P_Name, Type_param, count_halthours, list_ints)
                colmn += 1

                    #if Saldo[1] == 1:
                    #    Type_param = 10
                    #    ID_PP = request_bd_id_pp(ID_P, Type_param)
                    #    # ID_PP = request_bd_id_pp(ID_Points_prisoed, Type_param)
                    #    request_volume_interval(ID_PP, data_begin_report, data_end_report, colmn, Points_prisoed_Name,
                    #                           Type_param, count_halthours, list_ints)
                    #    colmn += 1


            #if ID_PP != "none":
            #    cursor_ms.execute('select val from PointMains where ID_PP = ? and DT >= ? and DT <= ?', ID_PP,
            #                   data_begin_report, data_end_report)
            #row = cursor_ms.fetchall()
