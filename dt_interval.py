import datetime
import calendar
import pandas as pd


def month_interval(dt_b, dt_e):
    list_date = []
    dt_Bg_year = int(dt_b[0][6:])
    dt_Bg_month = int(dt_b[0][3:5])
    dt_Bg_day = int(dt_b[0][:2])
    dt_Bg = datetime.datetime(dt_Bg_year, dt_Bg_month, dt_Bg_day)
    dt_count_m = dt_Bg_month
    dt_count_y = dt_Bg_year
    list_date.append(datetime.datetime(dt_count_y, dt_count_m, 1))

    dt_End_year = int(dt_e[0][6:])
    dt_End_month = int(dt_e[0][3:5])
    dt_End_day = int(dt_e[0][:2])
    dt_End = datetime.datetime(dt_End_year, dt_End_month, dt_End_day)
    years = dt_End_year - dt_Bg_year
    months = dt_End_month - dt_Bg_month
    if months < 0:
        months = 12 - dt_Bg_month + dt_End_month
        years -= 1
    months += years * 12
    if months == 0:
         months = 1
    for month in range(0, months):
        dt_count_m += 1
        if dt_count_m == 13:
            dt_count_m = 1
            dt_count_y += 1
        print(dt_count_m)
        list_date.append(datetime.datetime(dt_count_y, dt_count_m, 1))
    print(list_date)
    return list_date

def read_fl():
    df = pd.read_csv('C:\\Users\\admsys.SB-AIISKUE1.000\\aiiskue_py\\csv\\dt_interval.csv', encoding='windows-1251', sep=';', index_col=[0], error_bad_lines=False)
    dt_Bg_lst = list(df['dt_Bg'])
    dt_End_lst = list(df['dt_End'])
    list_dt = month_interval(dt_Bg_lst, dt_End_lst)
    return list_dt

