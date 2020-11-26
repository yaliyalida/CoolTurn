from flask import Blueprint, request, render_template
import json
from os import path
from werkzeug.utils import secure_filename
from .ExcelToWord.excel_to_table import excel_to_table
from .ExcelToWord.excel_to_text import excel_to_text, get_tplt

from .ProcessExcel.excel_Batch_conversion import *

from .WirteBatchWord.BatchWord import *

from .WordToExcel.WordToExcel2 import *

# 建立蓝图
EXCEL = Blueprint('EXCEL', __name__)


# ExcelToWord
@EXCEL.route('/import/ExcelToWord', methods=['POST', 'GET'])
def ExcelToWord():
    if request.method == 'POST':
        # 上传文件
        f = request.files['file']
        base_path = path.abspath(path.dirname(__file__))
        uploads = 'uploads\\'
        upload_path = path.join(base_path, uploads)
        file_name = upload_path + f.filename
        f.save(file_name)

        """
        excel_path = '学生信息表.xlsx'  # 要处理的excel文件路径
        save_path = ''  # 文件保存路径
        formatpath = 'template.docx'
        template = get_tplt(formatpath)  # 模板字符串
        name_column = '学号'
        """

        excel_path = file_name
        formatpath = request.form.get('formatpath')
        template = get_tplt(formatpath)
        name_column = request.form.get('name_column')
        choice = request.form.get('choice')

        if choice == '1':
            paths = excel_to_text(excel_path, template, name_column)
        elif choice == '2':
            paths = excel_to_table(excel_path,  name_column)

        paths = ['http://127.0.0.1:5000/' + p for p in paths]
        print(paths)

        data = {'url': paths}
        return json.dumps(data)
    else:
        return render_template('main.html')


# Excel_Batch_conversion
@EXCEL.route('/import/Excel_Batch_conversion', methods=['POST', 'GET'])
def Excel_Batch_conversion():
    if request.method == 'POST':
        # 上传文件
        choice = request.form.get('choice')

        if choice == '1':
            dir_path = request.form.get('dir_path')
            to_path = request.form.get('to_path')
            print("将{}文件夹下的表格进行简单合并".format(dir_path))
            paths = excel_simple_connect(dir_path, to_path)
        elif choice == '2':
            dir_path = request.form.get('dir_path')
            to_path = request.form.get('to_path')
            print("将{}文件夹下的表格进行连接合并".format(dir_path))
            paths = excel_combine_connect(dir_path, to_path)
        elif choice == '3':
            f = request.files['file']
            base_path = path.abspath(path.dirname(__file__))
            uploads = 'uploads\\'
            upload_path = path.join(base_path, uploads)
            file_name = upload_path + secure_filename(f.filename)
            f.save(file_name)

            file_path = file_name
            column = request.form.get('column')
            print("将{}表格进行分割".format(file_path))
            paths = excel_split(file_path, column)

        paths = ['http://127.0.0.1:5000/' + p for p in paths]
        print(paths)

        data = {'url': paths}
        return json.dumps(data)
    else:
        return render_template('main.html')


# BatchWord
@EXCEL.route('/import/WriteBatchWord', methods=['POST', 'GET'])
def WriteBatchWord():
    if request.method == 'POST':
        f = request.files['formatpath']
        base_path = path.abspath(path.dirname(__file__))
        uploads = 'uploads\\'
        upload_path = path.join(base_path, uploads)
        file_name = upload_path + f.filename
        f.save(file_name)

        formatpath = file_name
        template = read_text(formatpath)
        print("模板内容为:\n{}".format(template), end='')
        specific_chrt = "{}"  # 采用指定字符{}标记填写位置
        fieldnum = template.count(specific_chrt)  # 需要填写的值的数量
        print("每份word中需要填写的字段数量:{}".format(fieldnum))
        filenum = int(request.form.get('filenum'))


        fields = batch_write(filenum, fieldnum)

        name_rules = int(request.form.get('name_rules'))
        paths = write_words(template, specific_chrt, fields, filenum, name_rules)

        paths = ['http://127.0.0.1:5000/' + p for p in paths]
        print(paths)

        data = {'url': paths}
        return json.dumps(data)
    else:
        return render_template('main.html')


# WordToExcel
@EXCEL.route('/import/WordToExcel', methods=['POST', 'GET'])
def WordToExcel():
    if request.method == 'POST':
        f = request.files['startExcel']
        base_path = path.abspath(path.dirname(__file__))
        uploads = 'uploads\\'
        upload_path = path.join(base_path, uploads)
        file_name = upload_path + f.filename
        f.save(file_name)
        startExcel = file_name

        f = request.files['templete']
        base_path = path.abspath(path.dirname(__file__))
        uploads = 'uploads\\'
        upload_path = path.join(base_path, uploads)
        file_name = upload_path + f.filename
        f.save(file_name)
        templete = file_name

        wordDir = request.form.get('wordDir')
        readTemplate(startExcel, templete)
        docFiles = os.listdir(wordDir.encode('utf-8').decode("utf-8"))
        # 开始数据的行数
        row = 1
        paths = []
        for doc in docFiles:
            # 输出文件名
            try:
                row += 1
                to_path = 'static/export/WordToExcel/'
                name = to_path + doc.split('.')[0] + '.xls'
                writeExcel(wordDir + '\\' + doc.encode('utf-8').decode("utf-8"), row, name)
                paths.append(name)
            except Exception as e:
                print(e)

        paths = ['http://127.0.0.1:5000/' + p for p in paths]
        data = {'url': paths}
        return json.dumps(data)
    else:
        return render_template('main.html')
