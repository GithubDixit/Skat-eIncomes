#import mysql.connector
#import MySQLdb
import pymysql
#jdbc:db2://10.9.182.11:60000/d1585ind
conn = pymysql.connect(host="local",user="in098254",password="Test1234skat",db='OLD - INDB')
a = conn.cursor()
sql = 'select * from `h1585.FILKONTROLMI`;'
a.execute(sql)
countrow = a.execute(sql)



