from flask import render_template,flash,redirect,url_for
from app import app
from app.forms import LoginForm

# 装饰器，会修改跟在其后的函数，使用它们将函数注册为某些事件的回调函数
# 在作为参数给出的url和函数之间创建一个关联
# 在这里由两个装饰器，将两个url关联到这个函数
# 这意味着，当Web浏览器请求这两个url中的任何一个时，Flask将调用该函数并将其返回值作为响应传递回浏览器。
@app.route('/')
@app.route('/index')

def index():
    user = {'username':'Miguel'}
    posts = [
        {
            'author': {'username': 'John'},
            'body': 'Beautiful day in Portland!'
        },
        {
            'author': {'username': 'Susan'},
            'body': 'The Avengers movie was so cool!'
        }
    ]
    return render_template('index.html',title='Home',user=user,posts=posts)

# 装饰器中的methods参数，告诉Flask这个视图函数接受GET和POST请求，并覆盖了默认的GET
# 之前的“Method Not Allowed”错误正是由于视图函数还未配置允许POST请求。 通过传入methods参数，你就能告诉Flask哪些请求方法可以被接受。
@app.route('/login',methods=['GET','POST'])
def login():  # 登录视图函数
    form = LoginForm()
    if form.validate_on_submit():  # 执行form校验
    # 当浏览器发起GET请求的时候，它返回False，这样视图函数就会跳过if块中的代码，直接转到视图函数的最后一句来渲染模板
    # 当用户在浏览器点击提交按钮后，浏览器会发送POST请求。form.validate_on_submit()就会获取到所有的数据，运行字段各自的验证器，全部通过之后就会返回True，这表示数据有效。
        # flash()函数是向用户显示消息的有效途径。 许多应用使用这个技术来让用户知道某个动作是否成功。
        flash('Login requested for user{}, remember_me={}'.format(
            form.username.data,form.remember_me.data))
        # 指引浏览器自动重定向到它的参数所关联的URL。当前视图函数使用它将用户重定向到应用的主页。
        return redirect(url_for('index'))
    return render_template('login.html',title='Sign In', form=form)
    # <html>
    #     <head>
    #         <title>Home Page - Microblog</title>
    #     </head>
    #     <body>
    #         <h1>Hello,''' + user['username'] +'''!</h1>
    #     </body>
    # </html>
    # '''

