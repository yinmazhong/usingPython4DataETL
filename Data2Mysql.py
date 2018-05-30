import pymysql


# 用python 连接mysql　so eazy
config = {
    'host': '×××.cn-north-1.amazonaws.com.cn',
    'port': 60916,
    'user': 'dataDev',
    'passwd': '×××××',
    'charset':'utf8mb4',
    'cursorclass':pymysql.cursors.DictCursor
    }
conn = pymysql.connect(**config)
conn.autocommit(1)
cursor = conn.cursor()

try:
    cursor.execute("use online_public;")
    count = cursor._query('SELECT * FROM comfort_sku')
    print('total records:', count)
    one=cursor.fetchone()
    print("one:",one)
    all=cursor.fetchall()
    print("all:",all)


except:
    import traceback
    traceback.print_exc()
    conn.rollback()
finally:
    cursor.close()
    conn.close()