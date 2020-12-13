import pymysql


class Register:

    def __init__(self, account, passwd1,passwd2,type):
        self.account = account
        self.passwd1 = passwd1
        self.passwd2 = passwd2
        self.type=type

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
        self.userdata = self.cursor.fetchall()
        self.acoountNo = len(self.userdata)
        self.flag = 0  # 标记账号是否存在

    def check(self):
        for i in range(self.acoountNo):
            if(self.account==self.userdata[i][0]):
                self.flag=1
                print("用户已存在")
                return False
        if(self.flag==0):
            return True

    def register(self):
        if(self.check()==False):
            return 0
        elif(self.passwd1!=self.passwd2):
            print("两次输入密码不一致")
            return 2
        else:
            self.cursor.execute("insert into user values(%s,%s,%s)",(self.account,self.passwd1,self.type))
            self.con.commit()
            # print("注册成功")
            return 1


if __name__ == "__main__":
    pass
