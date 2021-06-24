# -*- coding: utf-8 -*-
import os
import sys
import math
import json
import pandas as pd


# 导出详细的计算结果为Isolation_list_allinfo.xlsx
def OutputInfoExcel(start_mz, end_mz, inclusion_list, isolation_window, cz, file_path_name):
    all_info = {'start_mz': start_mz,
                'end_mz': end_mz,
                'inclusion_list': inclusion_list,
                'isolation_window': isolation_window,
                'cz': cz}
    df_allinfo = pd.DataFrame(all_info)
    df_allinfo.to_excel(file_path_name + 'Isolation_list_allinfo.xlsx', index=False)


# 导出能直接用QExactive系列质谱上机的计算结果为Isolation_list_HFX.csv
def OutputCsv(start_mz, end_mz, inclusion_list, isolation_window, cz, file_path_name):
    positive = list()
    NCE = list()
    NCEtype = list()
    for x in range(0, len(end_mz)):
        positive.append('Positive')
        NCE.append('28')
        NCEtype.append('NCE')
    HFX_list = {'Mass [m/z]': inclusion_list,
                'Formula [M]': None,
                'Formula type':None,
                'Species': None,
                'CS [z]': cz,
                'Polarity': positive,
                'Start [min]': None,
                'End [min]': None,
                '(N)CE': NCE,
                '(N)CE type': NCEtype,
                'MSX ID': None,
                'Comment': None}
    df_HFXlist = pd.DataFrame(HFX_list)
    df_HFXlist.to_csv(file_path_name + 'Isolation_list_HFX.csv', index=False, encoding='utf-8-sig')


def main():
    try:
        # 获取PD自带的输入%NODEARGS%
        a = sys.argv[1:]
        f = open(a[0])
        t = json.load(f)
        table = t['Tables']
        path = table[0]['DataFile']
        # 获取搜库文件存放的路径
        data_path = os.path.dirname(t['ResultFilePath'])
        file_title_name = data_path.split('\\')[-1]
        file_path_name = data_path + '\\' + file_title_name
        f.close()
        # 读取Peptide Group数据表
        df = pd.read_table(path)
        # 默认窗口数40，最大m/z检出限1500
        windows = 40
        max_mz = 1500
        # 读取m/z列的数值,并从小到大排序
        mz = df['mz in Da by Search Engine Sequest HT'].sort_values()
        # 将m/z转至列表进行后续计算
        mz = mz.tolist()
        mod, start_mz, end_mz, inclusion_list, isolation_window, cz = list(), list(), list(), list(), list(), list()
        # 向上取整
        c = math.ceil(len(mz) / windows)
        # 计算窗口特定位置
        for b in range(1, len(mz) + 1):
            mods = b % c
            if mods == 1:
                mod.append(b)
        # 保存窗口起始m/z
        for j in mod:
            k = j - 1
            l = round(mz[k])
            start_mz.append(l)
        # 保存窗口结束m/z
        for d in range(1, windows):
            end_mz.append(start_mz[d])
        end_mz.append(max_mz)
        # 保存cz值
        for i in range(0, len(start_mz)):
            inclusion = (end_mz[i] + start_mz[i]) / 2
            isolation = end_mz[i] - start_mz[i] + 1
            inclusion_list.append(inclusion)
            isolation_window.append(isolation)
            if inclusion <= 500:
                cz.append(3)
            else:
                cz.append(2)
        # 导出计算文件
        OutputInfoExcel(start_mz, end_mz, inclusion_list, isolation_window, cz, file_path_name)
        OutputCsv(start_mz, end_mz, inclusion_list, isolation_window, cz, file_path_name)
    except Exception as info:
            print(info)


if __name__=='__main__':
    main()