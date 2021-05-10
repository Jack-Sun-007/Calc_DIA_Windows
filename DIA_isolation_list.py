# -*- coding: utf-8 -*-
import pandas as pd
import math
import os


#输入要划分的窗口数
def InputWindows():
    while True:
        try:
            windows = input('请输入DIA窗口数：')
            windows = int(windows)
            if windows < 10 or windows > 200:
                print('请输入一个合理的窗口数(一般为40~50)')
                windows = InputWindows()
            break
        except ValueError:
            print('请输入一个合理的数字')
    return windows


#输入m/z最大检出限
def InputMaxMz():
    while True:
        try:
            max_mz = input('请输入m/z最大检出限：')
            max_mz = int(max_mz)
            if max_mz < 1000 or max_mz > 3000:
                print('请输入一个合理的数字(一般为1300或1500)')
                max_mz = InputMaxMz()
            break
        except ValueError:
            print('请输入一个合理的数字')
    return max_mz


# 查找当前目录带Peptide名字的Excel表
def FindPath():
    try:
        p = str(os.getcwd())
        for filename in os.listdir(p):
            if "eptide" in str(filename):
                path = str(filename)
                break
        return path
    except:
        print('请检查当前目录是否含有的Peptide的Excel表')


# 导出详细的计算结果为Isolation_list_allinfo.xlsx
def OutputInfoExcel(start_mz, end_mz, inclusion_list, isolation_window, cz):
    all_info = {'start_mz': start_mz,
                'end_mz': end_mz,
                'inclusion_list': inclusion_list,
                'isolation_window': isolation_window,
                'cz': cz}
    df_allinfo = pd.DataFrame(all_info)
    df_allinfo.to_excel('Isolation_list_allinfo.xlsx', index=False)

# 导出能直接用QExactive系列质谱上机的计算结果为Isolation_list_HFX.csv
def OutputCsv(start_mz, end_mz, inclusion_list, isolation_window, cz):
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
    df_HFXlist.to_csv('Isolation_list_HFX.csv', index=False, encoding='utf-8-sig')

def main():
    try:
        windows = InputWindows()
        max_mz = InputMaxMz()
        print("正在计算，请等十秒钟")
        # 读取Excel的第一个表单
        df = pd.read_excel(FindPath())
        # 读取m/z列的数值,并从小到大排序
        mz = df['m/z [Da] (by Search Engine): Sequest HT'].sort_values()
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
        OutputInfoExcel(start_mz, end_mz, inclusion_list, isolation_window, cz)
        OutputCsv(start_mz, end_mz, inclusion_list, isolation_window, cz)
    except Exception as info:
            print("出错了，错误原因如下，请重新运行")
            print(info)
            os.system("pause")


if __name__=='__main__':
    main()