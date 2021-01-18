import data_base.init_bd_mssql as mssql
import data_base.init_bd_mssql as mysql

cursor_ms, conn_ms = mssql.init_mssql()
cursor_my, conn_my = mysql.init_mysql()

cursor_my.execute("SELECT imei, description, serveraddress, con_time, csq, state, phone1 FROM devices")
row = cursor_my.fetchall()
lenth = len(row)

for devices in range(0, lenth):
    try:
        cursor_ms.execute('select Name, URL, TimeTable from USD where URL = ?', row[devices][2])  # забираем из БД тип точки учета
        points_addr = cursor_ms.fetchone()
        print(points_addr[0], points_addr[1], ' ', row[devices][1], row[devices][2])

        querry_my = "UPDATE devices SET description = " + "'" + points_addr[0] + "'" + "where serveraddress = " + "'" + points_addr[1] + "'"
        cursor_my.execute(querry_my)

    except TypeError:
        print('в таблице IRZ IP адреса ', points_addr[1], 'не существует')

conn_my.commit()
cursor_my.close()
conn_my.close()

conn_ms.commit()
cursor_ms.close()
conn_ms.close()
