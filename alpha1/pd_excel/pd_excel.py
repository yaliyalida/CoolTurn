import pandas as pd
import os

def get_datas(save_name):
    if not os.path.exists(save_name):  # 如果数据没进行过合并
        entprise_info = 'entprise_info.csv'
        new_base_info = 'new_base_info.csv'
        # 返回为DataFrame对象
        df_entprise = pd.read_csv(entprise_info)
        df_base = pd.read_csv(new_base_info)
        data = pd.merge(df_base, df_entprise, how='inner', on='id')
        data.to_excel(save_name)
        path = os.getcwd()
        print('datas already saved in {}'.format(path))

    datas = pd.read_excel(save_name)
    return datas

save_name = 'base_datas.xlsx'
datas = get_datas(save_name)
print(datas)
