# Import `os`
import os
import pandas as pd

import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()

# Change directory. Директория, где лежат файлы

os.chdir(r"C:\Users\AVShestakov\Лицензии L900\Этап 2")

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
df_Delta_2G = df_Task_2G['% Покрытия 2G в полигоне после'].round(2) - df_Task_2G['% Покрытия 2G в полигоне'].round(2)
df_Delta_4G = df_Task_4G['% Покрытия 4G в полигоне после'].round(2) - df_Task_4G['% Покрытия 4G в полигоне'].round(2)

# вставляем % прироста покрытия в ТЗ и убираем прирост меньше 1
df_Task_2G.insert(loc=len(df_Task_2G.columns), column='% прироста ТЗ', value=df_Delta_2G)
df_Task_2G = df_Task_2G[(df_Task_2G['% прироста ТЗ'] > 1) | (df_Task_2G['% прироста ТЗ'] < -1)]

df_Task_4G.insert(loc=len(df_Task_4G.columns), column='% прироста ТЗ', value=df_Delta_4G)
df_Task_4G = df_Task_4G[(df_Task_4G['% прироста ТЗ'] > 1) | (df_Task_4G['% прироста ТЗ'] < -1)]

# для возможности сравнения ТЗ и BSSI должны быть одинаковые названия колонок
# (хотя и необязательно, можно left_on= и right_on= применить)
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
# ищем несоответсвие прироста покрытия

# для 2G
result_2G = (df_BSSI_2G.merge(df_Task_2G,
                              on='Индекс для BSSi',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

result_2G['дельта прироста 2G'] = (
    (result_2G['% прироста ТЗ'] - result_2G['% Прироста покрытия в полигоне']).round(2)).astype(int, errors='ignore')


result_2G_no_inBSSI = result_2G[((result_2G['_merge']) == 'right_only')].drop_duplicates(subset=['Индекс для BSSi'])
result_2G_no_inBSSI = result_2G_no_inBSSI['Индекс для BSSi']

result_2G_coverage_mismatch = result_2G[(result_2G['_merge'] == 'both') & (result_2G['дельта прироста 2G'] >= 3) |
                                        (result_2G['_merge'] == 'both') &
                                        (result_2G['дельта прироста 2G'] <= -3)].drop_duplicates(subset=['Индекс для '
                                                                                                         'BSSi'])
result_2G_coverage_mismatch = result_2G_coverage_mismatch[['Индекс для BSSi', 'дельта прироста 2G']]

# для 4G
result_4G = (df_BSSI_4G.merge(df_Task_4G,
                              on='Индекс для BSSi',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

result_4G['дельта прироста 4G'] = (result_4G['% прироста ТЗ'] - result_4G['% Прироста покрытия в полигоне']).round(2)

result_4G_no_inBSSI = result_4G[((result_4G['_merge']) == 'right_only')].drop_duplicates(subset=['Индекс для BSSi'])
result_4G_no_inBSSI = result_4G_no_inBSSI['Индекс для BSSi']

result_4G_coverage_mismatch = result_4G[(result_4G['_merge'] == 'both') & (result_4G['дельта прироста 4G'] >= 3) |
                                        (result_4G['_merge'] == 'both') &
                                        (result_4G['дельта прироста 4G'] <= -3)].drop_duplicates(subset=['Индекс для '
                                                                                                         'BSSi'])
result_4G_coverage_mismatch = result_4G_coverage_mismatch[['Индекс для BSSi', 'дельта прироста 4G']]
#
# m = (df1.merge(df2, how='outer', on=['стандарт','ID объекта из файла задания'],
#               suffixes=['', '_new'], indicator=True))
# m2 = (df2.merge(df1, how='outer', on=['стандарт','ID объекта из файла задания'],
#               suffixes=['', '_new'], indicator=True))
#
# m3=pd.merge(m.query("_merge=='right_only'"), m2.query("_merge=='right_only'"),
# how ='outer').drop_duplicates(subset=['стандарт','ID объекта из файла задания'])
#
# m3.query("_merge=='right_only'").to_excel('out.xlsx')
# print(df1)
# print('---------------------------')
# print(df2)

# запишем данные в отдельный файл excel
# Указать writer библиотеки
writer2G = pd.ExcelWriter('result_2G.xlsx', engine='xlsxwriter')
writer4G = pd.ExcelWriter('result_4G.xlsx', engine='xlsxwriter')
writer_sent2G = pd.ExcelWriter('result_sent2G.xlsx', engine='xlsxwriter')
writer_sent4G = pd.ExcelWriter('result_sent4G.xlsx', engine='xlsxwriter')

# Записать ваш DataFrame в файл
result_2G.to_excel(writer2G, 'Sheet1', index=False)
result_2G_no_inBSSI.to_excel(writer_sent2G, 'отсутствует в BSSI', index=False)
result_2G_coverage_mismatch.to_excel(writer_sent2G, 'дельта покрытия', index=False)
result_4G.to_excel(writer4G, 'Sheet1', index=False)
result_4G_no_inBSSI.to_excel(writer_sent4G, 'отсутствует в BSSI', index=False)
result_4G_coverage_mismatch.to_excel(writer_sent4G, 'дельта покрытия', index=False)

writer2G.close()
writer4G.close()
writer_sent2G.close()
writer_sent4G.close()
