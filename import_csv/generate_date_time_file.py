import win32com.client
import pandas as pd
import datetime
import calendar
import xlwt


# функция "приписки" "0"
def to_str(val):
    if val < 10:
        val_str = '0' + str(val)
    else:
        val_str = str(val)
    return val_str


# функция формирования получасовых интервалов на сутки
def days_hh(Actual_day, Actual_month, Actual_year, days_on_month, count):
    for i in range(1, 49):
        h = i // 2
        if h == 24:
            h = 0
            Actual_day += 1
            if Actual_day > days_on_month:
                Actual_day = 1
                Actual_month += 1
                if Actual_month == 13:
                    Actual_month = 1
                    Actual_year += 1
        hh = to_str(h)
        if i % 2 == 0:
            mm = "00"
        else:
            mm = "30"
        str_time = hh + ':' + mm
        Actual_day_str = to_str(Actual_day)
        Actual_month_str = to_str(Actual_month)
        Actual_str = Actual_day_str + "." + Actual_month_str + "." + str(Actual_year)
        print(i, str_time, Actual_str)
        ws.write(count, 2, str_time)
        ws.write(count, 3, Actual_str)
        count += 1

    return Actual_day, Actual_month, Actual_year, count


df = pd.read_csv('date_potreblenie.csv', encoding='windows-1251', sep=';', index_col=[0], error_bad_lines=False)

# дата начала
dt_Bg_lst = list(df['dt_Bg'])
dt_Bg_year = int(dt_Bg_lst[0][6:])
dt_Actual_year = dt_Bg_year
dt_Bg_month = int(dt_Bg_lst[0][3:5])
dt_Actual_month = dt_Bg_month
dt_Bg_day = int(dt_Bg_lst[0][:2])
dt_Actual_day = dt_Bg_day
dt_Bg = datetime.datetime(dt_Bg_year, dt_Bg_month, dt_Bg_day)
days_of_Actual_month = calendar.monthrange(dt_Actual_year, dt_Actual_month)[1]  # количество дней в текущем месяце


# дата окончания
dt_End_lst = list(df['dt_End'])
dt_End_year = int(dt_End_lst[0][6:])
dt_End_month = int(dt_End_lst[0][3:5])
dt_End_day = int(dt_End_lst[0][:2])
dt_End = datetime.datetime(dt_End_year, dt_End_month, dt_End_day)
count_Years = dt_End_year - dt_Bg_year

Excel = win32com.client.Dispatch("Excel.Application")
wb = xlwt.Workbook()
ws = wb.add_sheet(str(dt_Bg_lst[0]) + "-" + str(dt_End_lst[0]))

delta = dt_End - dt_Bg
count_Days = delta.days
count = 1
while count_Days:
    dt_Actual_day, dt_Actual_month, dt_Actual_year, count = days_hh(dt_Actual_day, dt_Actual_month, dt_Actual_year, days_of_Actual_month, count)
    days_of_Actual_month = calendar.monthrange(dt_Actual_year, dt_Actual_month)[1]
    count_Days -= 1
fl = "data_time-" + str(dt_Bg_lst[0]) + "-" + str(dt_End_lst[0]) + ".xls"
wb.save(fl)
