import datetime
import calendar


time_zones = {"UTC": 4,
              "UTC+1": 3,
              "UTC+2": 2,
              "UTC+3": 1,
              "UTC+4": 0,
              "UTC+5": -1,
              "UTC+6": -2,
              "UTC+7": -3,
              "UTC+8": -4,
              "UTC+9": -5,
              "UTC+10": -6,
              "UTC+11": -7,
              "UTC+12": -8,
              "UTC-12": -8,
              "UTC-11": -9,
              "UTC-10": -10,
              "UTC-9": -11,
              "UTC-8": -12,
              "UTC-7": +11,
              "UTC-6": +10,
              "UTC-5": +9,
              "UTC-4": +8,
              "UTC-3": +7,
              "UTC-2": +6,
              "UTC-1": +5
            }


def convert_TZ(date, tz):
    year = int(date.year)
    month = int(date.month)
    day = int(date.day)
    hour = int(date.hour)
    # min = int(date.min)
    days_of_month = calendar.monthrange(year, month)

    hour_bd = hour + time_zones.get(tz)
    if hour_bd < 0:
        hour_bd = 24 + hour_bd
        day -= 1
        if day == 0:
            month -= 1
            if month == 0:
                month = 12
                year -= 1
            day = calendar.monthrange(year, month)[1]

    year_bd = year
    month_bd = month
    day_bd = day
    # min_bd = min
    date_bd = datetime.datetime(year_bd, month_bd, day_bd, hour_bd, 00, 00)


    return date_bd



dt = convert_TZ(datetime.datetime(2021, 1, 29, 1, 0, 0), "UTC+3")
print("время БД: ", dt)
