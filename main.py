# Import `os`
import os
import pandas as pd
import xlsxwriter
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()

# Change directory

os.chdir(r"C:\Users\AVShestakov\Расчет новостроек\Юг")

# List all files and directories in current directory
os.listdir('.')

# print(f"Список файлов: {os.listdir('.')}")

# Загружаем файлы в переменные `file`
file = 'Coverage_BSSI.xlsx'
file2 = 'Coverage_Task.xlsx'

# Загружаем spreadsheet в объект pandas
# xl = pd.ExcelFile(file)
# xl2 = pd.ExcelFile(file2)

# Печатаем название листов в данном файле
# print(f'Список листов файла {file}: {xl.sheet_names}')
# print(f'Список листов файла {file2}: {xl2.sheet_names}')

# Загрузить лист в DataFrame по его имени: df
df_BSSI = pd.read_excel(file, sheet_name='Лист1')  # BSSI
df_Task = pd.read_excel(file2, sheet_name='Лист1')  # ТЗ

# ТЗ делим на 2G и 4G
df_Task_2G = df_Task[['Индекс для BSSi', '% Покрытия 2G в полигоне', '% Покрытия 2G в полигоне после']]
df_Task_4G = df_Task[['Индекс для BSSi', '% Покрытия 4G в полигоне', '% Покрытия 4G в полигоне после']]

# рассчитаем прирост покрытия в ТЗ
df_Delta_2G = df_Task_2G['% Покрытия 2G в полигоне после'] - df_Task_2G['% Покрытия 2G в полигоне']
df_Delta_4G = df_Task_4G['% Покрытия 4G в полигоне после'] - df_Task_4G['% Покрытия 4G в полигоне']

# вставляем % прироста покрытия в ТЗ
df_Task_2G.insert(loc=len(df_Task_2G.columns), column='% прироста ТЗ', value=df_Delta_2G)
df_Task_4G.insert(loc=len(df_Task_4G.columns), column='% прироста ТЗ', value=df_Delta_4G)

# для возможности сравнения ТЗ и BSSI должны быть одинаковые названия колонок
df_BSSI = df_BSSI.rename(columns={'ID объекта из файла задания': 'Индекс для BSSi'})

# BSSI делим на 2G и 4G
df_BSSI_2G = df_BSSI.loc[(df_BSSI['Стандарт'] == '2G')]
df_BSSI_4G = df_BSSI.loc[(df_BSSI['Стандарт'] == '4G')]

# duplicateRows = df_BSSI_2G[df_BSSI_2G['Индекс для BSSi'].duplicated()]
# print(duplicateRows)

# суммируем прирост покрытия полигонов с несколькими решениями (дубликаты)
df_BSSI_2G = df_BSSI_2G.groupby('Индекс для BSSi')['% Прироста покрытия в полигоне'].sum().reset_index()
df_BSSI_4G = df_BSSI_4G.groupby('Индекс для BSSi')['% Прироста покрытия в полигоне'].sum().reset_index()

# ищем соответствие полигонов, все ли совпадают
# для 2G
result_2G = (df_BSSI_2G.merge(df_Task_2G,
                              on='Индекс для BSSi',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

result_2G['дельта прироста 2G'] = (result_2G['% прироста ТЗ'] - result_2G['% Прироста покрытия в полигоне']).round(2)

# для 4G
result_4G = (df_BSSI_4G.merge(df_Task_4G,
                              on='Индекс для BSSi',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

result_4G['дельта прироста 4G'] = (result_4G['% прироста ТЗ'] - result_4G['% Прироста покрытия в полигоне']).round(2)
#
# m = (df1.merge(df2, how='outer', on=['стандарт','ID объекта из файла задания'],
#               suffixes=['', '_new'], indicator=True))
# m2 = (df2.merge(df1, how='outer', on=['стандарт','ID объекта из файла задания'],
#               suffixes=['', '_new'], indicator=True))
#
# m3=pd.merge(m.query("_merge=='right_only'"), m2.query("_merge=='right_only'"), how ='outer').drop_duplicates(subset=['стандарт','ID объекта из файла задания'])
#
# m3.query("_merge=='right_only'").to_excel('out.xlsx')
# print(df1)
# print('---------------------------')
# print(df2)

# запишем данные в отдельный файл excel
# Указать writer библиотеки
writer2G = pd.ExcelWriter('result_2G.xlsx', engine='xlsxwriter')
writer4G = pd.ExcelWriter('result_4G.xlsx', engine='xlsxwriter')

# Записать ваш DataFrame в файл
result_2G.to_excel(writer2G, 'Sheet1')
result_4G.to_excel(writer4G, 'Sheet1')

writer2G.close()
writer4G.close()
