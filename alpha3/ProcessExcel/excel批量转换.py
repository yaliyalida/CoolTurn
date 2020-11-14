import pandas as pd
import os

def excel_simple_connect():
    print("请输入待合并的excel表格所在的目录路径：")
    get_path = input().replace('\\','/')

    # path = 'C:\Users\ASUS/Desktop\存放excell表格_简单合并'
    count = 0
    for file in os.scandir(get_path):
        path_ex = get_path + '/' + file.name
        df = pd.read_excel(path_ex,sheet_name='Sheet1')
        if count ==0:
            df_total = pd.read_excel(path_ex)
            count += 1
            continue
        df_total = pd.concat([df_total,df],axis=0)
    print("请输入简单合并后的excel表格的导出路径：")
    to_path = input().replace('\\','/')
    df_total.to_excel(to_path)
    # df_total.to_excel('C:\Users\ASUS\Desktop\excel_total.xlsx')
    # read = pd.read_excel('C:/Users/ASUS/Desktop/excel_total.xlsx')
    # print(read)

def combine_data(x):
    x.set_index('sno')
    col_name_list = x.columns.values
    data_list = [['NAN']]*len(col_name_list)
    for i in range(x.shape[1]):
        for j in range(x.shape[0]):
            if not pd.isnull(x.iloc[j][i]):
                # print((i,j),x.iloc[j][i])
                data_list[i] = x.iloc[j][i]
                # print(data_list)
                break
    # print(col_name_list,data_list)
    data_list = [[i] for i in data_list]
    # print(data_list)
    dic = dict()
    for i in range(len(col_name_list)):
        dic[col_name_list[i]] = data_list[i]
    df = pd.DataFrame(dic)
    print(df)
    return df

def excel_combine_connect():
    print("请输入待合并的excel表格所在的目录路径：")
    get_path = input().replace('\\','/')
    # path = 'C:\Users\ASUS\Desktop\存放excell表格_连接合并'
    count = 0
    # key_list = input("请输入要进行连接的关键字(关键字之间用逗号隔开):").split(',')
    # key_list = ['sno']
    for file in os.scandir(get_path):
        if count == 0:
            path_ex = get_path + '/' + file.name
            df_total = pd.read_excel(path_ex,sheet_name='Sheet1',keep_default_na=False)
            # df_total = df_total.set_index('sno')
            # print(df_total)
            count += 1
            continue
        path_ex = get_path + '/' + file.name
        df = pd.read_excel(path_ex,sheet_name='Sheet1',keep_default_na=False)
        # df = df.set_index('sno')
        # df_total.set_index('sno')
        df_total = pd.merge(df_total,df,how='outer')
        # print(df_total)
        # count += 1
    print(df_total)
    groups = df_total.groupby(['sno']).apply(combine_data)
    print(groups)
    groups.set_index('sno')
    print("请输入连接合并后的excel表格的导出路径(路径最后为您要保存的文件名：")
    to_path = input().replace('\\','/')
    groups.to_excel(to_path)
    # os.chdir('C:\Users\ASUS\Desktop\')
    # groups.to_excel('C:\Users\ASUS\Desktop\连接合并_总表.xlsx')

def excel_split():
    print("请输入待分割的excel表格的路径：")
    get_path = input().replace('\\','/')
    # xlsx_name = r'C:\Users\ASUS\Desktop\excel一到多\excel一到多总表.xlsx'
    #用来筛选的列名
    print("请输入用于筛选的列名：")
    filter_column_name = input()
    # filter_column_name = 'sno'
    #将该列去重后保存为list
    # df = pd.read_excel(xlsx_name)
    df = pd.read_excel(get_path)
    all_names = df[filter_column_name].unique().tolist()
    #获取所有sheet名
    # df = pd.ExcelFile(xlsx_name)
    df = pd.ExcelFile(get_path)
    sheet_names = df.sheet_names
    print("请输入分割后的多个excel表格需要保存的目录路径：")
    to_path = input().replace('\\','/')
    # os.chdir(r'C:\Users\ASUS\Desktop\excel一到多')
    for one_name in all_names:
        one_excel_name =to_path + '/' + str(one_name) + '.xlsx'
        writer = pd.ExcelWriter(one_excel_name)
        one_name_to_list = []
        one_name_to_list.append(one_name)
        for sheet_name in sheet_names:
            # tmp_df = pd.read_excel(xlsx_name, sheet_name=sheet_name)
            tmp_df = pd.read_excel(get_path, sheet_name=sheet_name)
            tmp_sheet = tmp_df[tmp_df[filter_column_name].isin(one_name_to_list)]
            tmp_sheet.to_excel(excel_writer=writer, sheet_name=sheet_name, encoding="utf-8", index=False)
        writer.save()
        writer.close()
def main():
    print("excel表格简单合并请输入1，excel表格连接合并请输入2，分割excel表格请输入3")
    n = eval(input())
    if n == 1:
        excel_simple_connect()
    elif n == 2:
        excel_combine_connect()
    elif n == 3:
        excel_split()
    else:
        print("输入有误")
