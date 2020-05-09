import pymysql

class DB():

    def conn(self):
        connect = pymysql.Connect(
            host='192.168.1.245',
            port=3306,
            user='root',
            passwd='123456',
            db='wbs',
            charset='utf8'
        )
        try:
            connect.ping()
        except:
            print("connect fail")
            self.conn()
        else:
            return connect


    def search(self,sql,value,connect=None):
        if not connect:
            connect = self.conn()
        cur = connect.cursor()
        cur.execute(sql, value)
        result = cur.fetchall()
        return result














