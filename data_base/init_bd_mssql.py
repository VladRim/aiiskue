import pyodbc
import config


def init_mssql():
    # Формирование строки SQL
    driver = config.DRIVER
    server = config.SERVER_ADDR
    user = config.UNAME
    pw = config.PASS
    port = config.PORT
    db = config.DATABASE
    conn_str = ';'.join([driver, server, port, db, user, pw])

    conn = pyodbc.connect(conn_str)
    cs = conn.cursor()
    return cs, conn
