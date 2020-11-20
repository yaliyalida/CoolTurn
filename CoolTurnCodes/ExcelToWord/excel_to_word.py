from excel_to_table import excel_to_table
from excel_to_text import excel_to_text,get_tplt


if __name__ == "__main__":
    excel_path = '学生信息表.xlsx'  # 要处理的excel文件路径
    save_path = ''  # 文件保存路径
    formatpath = 'template.docx'
    template = get_tplt(formatpath)  # 模板字符串
    name_column = '学号'
    while True:
        try:
            choice = input("功能选择:\n1、以文本形式导出批量word\n2、以表格形式导出批量word\n请选择:(1/2):\n")
            if choice not in ('1','2'):
                raise ValueError('请输入"1"或"2"')
            break
        except Exception as err:
            print(err)
    if choice == '1':
        excel_to_text(excel_path, save_path, template, name_column)
    elif choice == '2':
        excel_to_table(excel_path, save_path, name_column)

