from openpyxl import load_workbook
import docx
from docx.oxml.ns import qn
import re
import os


# 转换成word文本
def excel_to_text(excel_path, save_path, template, name_column):
    # 读取excel xlsx文件
    wb = load_workbook(excel_path)
    # 获取所有sheet页名字
    sheet_names = wb.sheetnames
    # 定位到相应sheet页,[0]为sheet页索引
    ws = wb[sheet_names[0]]
    # 获取excel行数
    row_num = ws.max_row
    # 匹配模板字符
    tplt_list = re.findall(r'{\d+}', template)  # 以匹配到的元素列表返回

    # 找到name_column对应的列
    title_column = 0
    for title_ceil in list(ws.rows)[0]:
        title_column += 1
        if name_column == title_ceil.value:
            break

    # 跳过表头
    i = 2

    # 写入word文件
    while i <= row_num:
        tem = template
        for t in tplt_list:
            index = int(t[1:-1])  # 找到要填入的内容对应excel表格的单元格,提取{}中的数值
            value = str(ws.cell(row=i, column=index).value)  # 找到要替换的内容
            tem = tem.replace(t, value)  # 替换

        # 文件名
        name = str(ws.cell(row=i, column=title_column).value)

        # 创建word文档
        document = docx.Document()

        # 设置文档字体
        document.styles['Normal'].font.name = u'宋体'
        document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')

        # 向文档中写入内容
        document.add_paragraph(tem)

        # 输出路径的前面一部分到时候用用户输入的
        docx_path = save_path + name + '.docx'
        try:
            document.save(docx_path)
            print('文件{}已存储在{}'.format(docx_path,os.getcwd()))
        except Exception as err:
            print(err)

        i += 1

def get_tplt(formatpath):
    '''
    获取模板文件字符串,段落之间用换行符链接
    :param formatpath: 模板文件路径
    :return: 模板文件字符内容
    '''
    content = ''
    try:
        fileobj = docx.Document(formatpath)  # 模板文件对象
        for para in fileobj.paragraphs:
            content += para.text+'\n'
    except Exception as err:
        print(err)
    return content

if __name__ == "__main__":
    # 用户指定
    excel_path = '学生信息表.xlsx'  # 要处理的excel文件路径
    save_path = ''  # 文件保存路径
    formatpath = 'template.docx'
    template = get_tplt(formatpath)  # 模板字符串

    # name_column = input("请设置用来命名的列:")  # 设定命名列
    name_column = '学号'

    # 写入word文档
    excel_to_text(excel_path, save_path, template, name_column)
