import os
import pandas as pd

import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Change directory. Директория, где лежат файлы

os.chdir(r"C:\Users\AVShestakov\1-2 волна")

file = 'BSSI.xlsx'
file2 = 'Task.xlsx'

df_BSSI = pd.read_excel(file, sheet_name='Sheet1')  # BSSI
df_Task = pd.read_excel(file2, sheet_name='Sheet1')  # ТЗ

# BSSI делим на 2G, 3G, 4G
df_BSSI_4G = df_BSSI[df_BSSI['Стандарт'].str.contains('4G')]
df_BSSI_4G = df_BSSI_4G.drop_duplicates()
# Указываем writer библиотеки
writer_df_BSSI_4G = pd.ExcelWriter('result_4G.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
df_BSSI_4G.to_excel(writer_df_BSSI_4G, 'Sheet1', index=False)
writer_df_BSSI_4G.close()

result = (df_BSSI_4G.merge(df_Task,
                              on='Индекс для BSSI',
                              how='outer',
                              suffixes=['', '_new'],
                              indicator=True))

# Указываем writer библиотеки
writer_result = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')

# Записываем DataFrame в файл
result.to_excel(writer_result, 'Sheet1', index=False)
writer_result.close()