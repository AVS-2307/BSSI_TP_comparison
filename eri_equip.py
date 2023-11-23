# Import `os`
import os
import pandas as pd

import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()

# Change directory. Директория, где лежат файлы

os.chdir(r"C:\Users\AVShestakov\Расчет емкости\Ericsson")

file = 'inventory_Vse filialy.xlsx'
file2 = 'power_Все филиалы.xlsx'

# Загрузить лист в DataFrame по его имени: df
df_inventory = pd.read_excel(file, sheet_name='eri')  # inventory
df_power = pd.read_excel(file2, sheet_name='eri')  # power

# оставляем в inventory только требуемые DU
df_inventory['du'] = df_inventory['du'].replace('', None)
df_inventory['du'] = df_inventory['du'].fillna('ooo')
df_inventory = df_inventory[df_inventory['du'].str.contains('3101|3102|4102|5212|503|5216|6318|6502|6620|6630')]

# суммируем в inventory все возможные RRU
# при этом пустые ячейки заполняем нулями
df_inventory['rru_tot'] = df_inventory['rru_800'].fillna(0).astype(str) + ',' + df_inventory['rru_900'].fillna(
    0).astype(str) \
                          + ',' + df_inventory['rru_1800'].fillna(0).astype(str) + ',' + df_inventory[
                              'rru_2100'].fillna(0).astype(str) \
                          + ',' + df_inventory['rru_2600fdd'].fillna(0).astype(str) + ',' \
                          + df_inventory['rru_2600tdd'].fillna(0).astype(str) + ',' + df_inventory[
                              'rru_1800_2100'].fillna(0).astype(str)

# убираем нули, оставляем только RRU
df_inventory['rru_tot'] = df_inventory['rru_tot'].replace(regex=[r'0,', ',0'], value='')
# подсчитываем кол-во RRU для каждой строки
df_inventory['rru_tot_count'] = df_inventory['rru_tot'].str.split(",").apply(lambda x: len(x))

# writer = pd.ExcelWriter('df_inventory.xlsx', engine='xlsxwriter')
# df_inventory.to_excel(writer, 'eri', index=False)
# writer.close()
# оставляем в power только LTE
df_power['modes'] = df_power.groupby('namebs')['techn'].transform('sum')
df_power.insert(loc=len(df_power.columns), column='mode', value='')


def mode(row):
    if ('GSM' and 'FDD' in row['modes']) or ('GSM' and 'TDD' in row['modes']) or ('UMTS' and 'TDD' in row['modes']) or \
            ('UMTS' and 'FDD' in row['modes']) or ('GSM' and 'UMTS') in row['modes']:
        row['mode'] = 'mixed'
    else:
        row['mode'] = 'single'
    return row['mode']


df_power['mode'] = df_power.apply(mode, axis=1)

# df_power = df_power.loc[((df_power['techn'] == 'FDD') | (df_power['techn'] == 'TDD'))]

# writer = pd.ExcelWriter('df_power.xlsx', engine='xlsxwriter')
# df_power.to_excel(writer, 'eri', index=False)
# writer.close()

# объединяем inventory и power
invpow = (df_inventory.merge(df_power,
                             on='namebs',
                             how='left',
                             suffixes=['', '_new'],
                             indicator=True))

# добавляем колонки в разделы ПЛАНИРУЕМАЯ КОНФИГУРАЦИЯ и ИТОГ
invpow = invpow.reindex(columns=invpow.columns.tolist() + ['cells_1800 MIMO 2X2', 'cells_1800 MIMO 4X4',
                                                           'План BW4G1800', 'cells_2100 MIMO 2X2',
                                                           'cells_2100 MIMO 4X4', 'План BW4G2100', 'cells_2600fdd',
                                                           'План BW4G2600fdd', 'cells_2600tdd', 'План BW4G2600tdd',
                                                           ], fill_value=0)

invpow.drop(['oss', '2.5G_RAN_SFP', '10G_RAN_SFP+', 'SFP_others', 'ret', 'ret_sn', 'date_from', 'vendor',
             'oss_new', 'filial_new', 'purpose', 'dop', 'el_tilt', 'el_tilt_fact', 'mech_tilt',
             'sum_tilt', 'height', 'azimuth', 'ret_new', 'id_motor', 'type_motor', 'conclusion',
             'rru_model', 'pwr_dbm_tx', 'pwr_dbm_total', 'GEOUNIT_ID', 'nrisiteid', 'new_ind',
             '_merge', 'ch_n', 'TX_div', 'pwr_w_tx', 'pwr_w_total', 'pwr_w_resudual',
             'antenna_model'], axis=1, inplace=True)

writer = pd.ExcelWriter('invpow.xlsx', engine='xlsxwriter')
invpow.to_excel(writer, 'eri', index=False)
writer.close()
# заполняем ИТОГ по потребностям в файл eri_equip_plan
