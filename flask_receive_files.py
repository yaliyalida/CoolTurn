@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':   # 如果是 POST 请求方式
        file = request.files['file']   # 获取上传的文件
        if file and allowed_file(file.filename):   # 如果文件存在并且符合要求则为 true
            filename = secure_filename(file.filename)   # 获取上传文件的文件名
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))   # 保存文件
            return '{} upload successed!'.format(filename)   # 返回保存成功的信息
    # 使用 GET 方式请求页面时或是上传文件失败时返回上传文件的表单页面