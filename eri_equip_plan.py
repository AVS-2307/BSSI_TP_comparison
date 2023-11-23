# Import `os`
import os
import time
import pandas as pd

import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()

# Change directory. Директория, где лежат файлы

os.chdir(r"C:\Users\AVShestakov\Расчет емкости\Ericsson")

# формируется файл ПЛАНОВОЙ КОНФИГУРАЦИИ, заполняются вручную требуемые ячейки файла invpow.xlsx,

start_time = time.time()
file3 = 'invpow.xlsx'
invpow_planned = pd.read_excel(file3, sheet_name='eri')


# заполняем ИТОГ по потребностям

def add_2600FDD(row):
    if row['cells_2600fdd'] > 0:
        val = row['cells_2600fdd']
    else:
        val = 0
    return val


def add_2600TDD(row):
    if row['cells_2600tdd'] > 0:
        val = row['cells_2600tdd']
    else:
        val = 0
    return val


def add_1800_2x2(row):
    if row['cells_1800 MIMO 2X2'] > 0 and row['TXRX'] == '2T2R' and (('RRUS 12 B3' in row['rru_tot'])
                                                                     or ('2219 B3' in row['rru_tot']) or
                                                                     ('2488 B3' in row['rru_tot']) or
                                                                     ('2242' in row['rru_tot']) or
                                                                     ('2279' in row['rru_tot'])):
        val = 0
    else:
        val = row['cells_1800 MIMO 4X4']
    return val


def add_1800_4x4(row):
    if row['cells_1800 MIMO 4X4'] > 0 and row['TXRX'] == '4T4R' and (('4428 B3' in row['rru_tot']) or
                                                                     ('4429' in row['rru_tot']) or
                                                                     ('4499' in row['rru_tot'])):
        val = 0
    else:
        val = row['cells_1800 MIMO 4X4']
    return val


def add_2100_2x2(row):
    if row['cells_2100 MIMO 2X2'] > 0 and row['TXRX'] == '2T2R' and (('RRUS 12 B1' in row['rru_tot'])
                                                                     or ('RRUS 11 B1' in row['rru_tot']) or
                                                                     ('2219 B1' in row['rru_tot']) or
                                                                     ('2217 B1' in row['rru_tot']) or
                                                                     ('RRUS 13 B1' in row['rru_tot']) or
                                                                     ('RRUS 12mB1' in row['rru_tot']) or
                                                                     ('Radio 2488 B1' in row['rru_tot']) or
                                                                     ('2242' in row['rru_tot']) or
                                                                     ('2279' in row['rru_tot'])):
        val = 0
    else:
        val = row['cells_2100 MIMO 2X2']
    return val


def add_2100_4x4(row):
    if row['cells_2100 MIMO 4X4'] > 0 and row['TXRX'] == '4T4R' and (('4428 B1' in row['rru_tot']) or
                                                                     ('4499' in row['rru_tot'])):
        val = 0
    else:
        val = row['cells_2100 MIMO 4X4']
    return val


# Расчет DU

# Для всех DU вводим расчет cells
# расчет cells в текущей конфигурации
invpow_planned['cells_current'] = invpow_planned['cells_800'] + invpow_planned['cells_900'] \
                                  + invpow_planned['cells_1800'] \
                                  + invpow_planned['cells_2100'] + invpow_planned['cells_2600']

# расчет cells текущая конфига + запланированная
invpow_planned['cells_after_ext'] = invpow_planned['cells_800'] + invpow_planned['cells_900'] \
                                    + invpow_planned['cells_1800'] + invpow_planned['cells_2100'] \
                                    + invpow_planned['cells_2600'] + invpow_planned['cells_1800 MIMO 2X2'] \
                                    + invpow_planned['cells_1800 MIMO 4X4'] + invpow_planned['cells_2100 MIMO 2X2'] \
                                    + invpow_planned['cells_2100 MIMO 4X4'] + invpow_planned['cells_2600fdd'] \
                                    + invpow_planned['cells_2600tdd']


# расчет ant_bw по типу DU
# расчет существующего ant_bw
def ant_bw_current(row):
    ant_bw2 = 0
    ant_bw4 = 0
    if '3101' in row['du'] and row['cells_current'] < 9:
        if row['TXRX'] == '2T2R':
            ant_bw2 = row['dlchbw'] * 2
        elif row['TXRX'] == '4T4R':
            ant_bw4 = row['dlchbw'] * 4
        ant_res = ant_bw2 + ant_bw4
        return ant_res
    elif '4102' in row['du'] and row['cells_current'] < 12:
        if row['TXRX'] == '2T2R':
            ant_bw2 = row['dlchbw'] * 2
        elif row['TXRX'] == '4T4R':
            ant_bw4 = row['dlchbw'] * 4
        ant_res = ant_bw2 + ant_bw4
        return ant_res
    elif '5212' in row['du'] and row['cells_current'] < 12:
        if row['TXRX'] == '2T2R':
            ant_bw2 = row['dlchbw'] * 2
        elif row['TXRX'] == '4T4R':
            ant_bw4 = row['dlchbw'] * 4
        ant_res = ant_bw2 + ant_bw4
        return ant_res


# расчет планируемого + существующего ant_bw
def ant_bw_total(row):
    ant_res_cell = row['План BW4G1800'] * row['cells_1800 MIMO 2X2'] * 2 + \
                   row['План BW4G1800'] * row['cells_1800 MIMO 4X4'] * 4 + \
                   row['План BW4G1800'] * row['cells_2100 MIMO 2X2'] * 2 + \
                   row['План BW4G2100'] * row['cells_2100 MIMO 4X4'] * 4 + \
                   row['План BW4G2600fdd'] * row['cells_2600fdd'] + \
                   row['План BW4G2600tdd'] * row['cells_2600tdd'] + row['antBW']

    if '3101' in row['du'] and row['antBW'] < 240:
        return ant_res_cell
    elif ('4102' in row['du'] or '5212' in row['du']) and row['antBW'] < 480:
        return ant_res_cell


# возможность запуска доп.несущих
def set_new_carriers(row):
    if '3101' in row['du']:
        if (row['cells_after_ext'] <= 9) and (0 < row['antBW_tot'] < 240) and (
                row['rru_tot_count'] <= 9):
            val = 'запуск возможен'
        else:
            val = 'запуск невозможен'
        return val

    elif '4102' in row['du'] or '5212' in row['du']:
        if (row['cells_after_ext'] <= 12) and (0 < row['antBW_tot'] < 480) and (
                row['rru_tot_count'] <= 12):
            val = 'запуск возможен'
        else:
            val = 'запуск невозможен'
        return val


invpow_planned['Потребность RRU LTE-2600FDD'] = invpow_planned.apply(add_2600FDD, axis=1)
invpow_planned['Потребность RRU LTE-2600TDD'] = invpow_planned.apply(add_2600TDD, axis=1)
invpow_planned['Потребность RRU LTE-1800 2Х2'] = invpow_planned.apply(add_1800_2x2, axis=1)
invpow_planned['Потребность RRU LTE-1800 4Х4'] = invpow_planned.apply(add_1800_4x4, axis=1)
invpow_planned['Потребность RRU LTE-2100 2Х2'] = invpow_planned.apply(add_2100_2x2, axis=1)
invpow_planned['Потребность RRU LTE-2100 4Х4'] = invpow_planned.apply(add_2100_4x4, axis=1)
invpow_planned['antBW'] = invpow_planned.apply(ant_bw_current, axis=1)
invpow_planned['antBW_planned'] = invpow_planned.apply(ant_bw_total, axis=1)
invpow_planned['antBW_tot'] = invpow_planned.groupby('namebs')['antBW_planned'].transform('sum')
invpow_planned['cells_planned'] = invpow_planned.apply(set_new_carriers, axis=1)

writer = pd.ExcelWriter('invpow_total.xlsx', engine='xlsxwriter')
invpow_planned.to_excel(writer, 'eri', index=False)
writer.close()

end_time = time.time()
# разница между конечным и начальным временем
elapsed_time = round((end_time - start_time), 2)
print('Elapsed time: ', elapsed_time)
