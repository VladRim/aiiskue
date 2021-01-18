import MySQLdb
import config

# соединяемся с MySQL
def init_mysql():
    server = config.ADDR_MY
    user = config.LOGIN_MY
    pw = config.PASS_MY
    db = config.BD_NAME_MY
    conn_mysql = MySQLdb.connect(server, user, pw, db, charset='utf8', init_command='SET NAMES UTF8')
    cursor = conn_mysql.cursor()

    return cursor, conn_mysql