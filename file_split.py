# Import `os`
import os
import pandas as pd
import xlsxwriter
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Change directory

os.chdir(r"C:\Users\AVShestakov\Расчет новостроек\Юг")

# Загружаем файлы в переменные `file`
file = 'Coverage_BSSI_before_split.xlsx'
file2 = 'Coverage_Task.xlsx'

# Загрузить лист в DataFrame по его имени: df
df_BSSI = pd.read_excel(file, sheet_name='Лист 1')  # BSSI
# df_Task = pd.read_excel(file2, sheet_name='Лист1')  # ТЗ

# для возможности сравнения ТЗ и BSSI должны быть одинаковые названия колонок
df_BSSI = df_BSSI.rename(columns={'ID объекта из файла задания': 'Индекс для BSSi'})

# BSSI делим на 2G и 4G
df_BSSI_2G = df_BSSI.loc[(df_BSSI['Стандарт'] == '2G')]
df_BSSI_4G = df_BSSI.loc[(df_BSSI['Стандарт'] == '4G')]

df_BSSI_2G = df_BSSI_2G['Индекс для BSSi'].str.split(r"[;,]", expand=True)
df_BSSI_2G.reset_index()
# df['new_values'].str.split('\n')[0]
# print(df_BSSI_2G['Индекс для BSSi'].str.split(';', expand=True))

writer2G = pd.ExcelWriter('result_2G_split.xlsx', engine='xlsxwriter')
df_BSSI_2G.to_excel(writer2G, 'Sheet1')
writer2G.close()