import pymysql
from flask import jsonify


class Login:

    def __init__(self,account,password,safety):
        self.account=account
        self.password=password
        self.safety=safety

        self.host = '127.0.0.1'  # MYSQL服务器地址
        self.port = 3306  # MYSQL服务器端口号
        self.user = 'root'  # 用户名
        self.passwd = "passwd"  # 密码
        self.db = 'market'  # 数据库名称
        self.charset = 'utf8'  # 连接编码
        self.con = pymysql.connect(host=self.host,
                                   port=self.port,
                                   user=self.user,
                                   passwd=self.passwd,
                                   db=self.db,
                                   charset=self.charset)
        self.cursor = self.con.cursor()  # 使用连接对象获得cursor对象
        self.cursor.execute("select username,password,safety from user")
        self.userdata=self.cursor.fetchall()
        self.acoountNo=len(self.userdata)
        self.flag=0#标记账号是否存在
        # self.match()

    #检测账号存在性和账号密码的匹配
    def match(self):
        for i in range(self.acoountNo):
            if(self.account==self.userdata[i][0] and self.userdata[i][2]==self.safety):
                self.flag=1
                if(self.password==self.userdata[i][1]):
                    print("登录成功")
                    return 1
                else:
                    print("账号密码不符")
                    return 2
        if(self.flag==0):
            print("账号不存在")
            return 0


if __name__ == "__main__":
    pass