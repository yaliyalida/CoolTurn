# app包由app目录和__init__.py脚本来定义构成，并在from app import routes语句中被引用
# app变量被定义为__init__.py脚本中的Flask类的一个实例，以至于它成为app包的属性
'''
routes模块是在底部导入的，而不是在脚本的顶部。
最下面的导入是解决循环导入的问题，你将会看到routes模块中需要导入这个脚本中定义的app变量，因此将routes的导入
放在底部可以避免由于这两个文件之间的相互引用而导致的错误
'''



from flask import Flask
from config import Config
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate

app = Flask(__name__)
app.config.from_object(Config)
db = SQLAlchemy(app)
migrate = Migrate(app, db)


from app import routes,models  # 导入尚未存在的routes模块

