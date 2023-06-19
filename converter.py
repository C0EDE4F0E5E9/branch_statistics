__author__ = 'Андрей Кравцов'
__version__ = "1.0.2"
__email__ = "kravtsov911@gmail.com"

import platform

import numpy as np
import pandas as pd
from colorama import Fore

from data_processing import get_data_frame, folder_scan
from read_templates import templates

match_table = templates.match_table_branch
column_names_dna = templates.column_names_dna


def convert(path_journal: str):
    """
    Аргументы
    _________
    - path_journal -- путь к журналам в каталоге с журналами головного офиса

    Возвращает
    __________
    - возвращает два файла xlsx, которые содержат сводные данные из журналов головного офиса

    """
    name_list_convert = []
    if platform == 'linux' or platform == 'linux2':
        path_convert = path_journal + '/г.Ессентуки/'
    else:
        path_convert = path_journal + '\\г.Ессентуки\\'
    folder_scan(path_convert, name_list_convert)
    expertise = pd.ExcelWriter(path_convert + 'Журнал экспертиз Ессентуки.xlsx')
    study = pd.ExcelWriter(path_convert + 'Журнал исследований Ессентуки.xlsx')
    consultation = pd.ExcelWriter(path_convert + 'Журнал следственных действий Ессентуки.xlsx')

    for file_name in name_list_convert:
        if 'xlsm' in file_name or 'xlsx' in file_name:
            for key in match_table:
                if key in file_name:
                    direction_name = match_table[key]
                    if direction_name == 'Фоноскопическая':
                        try:
                            print(Fore.YELLOW + 'ПРЕОБРАЗОВАНИЕ ЖУРНАЛА: ' + file_name)
                            df_exp_fono = get_data_frame(file_name, "Фоно_эксп")
                            df_exp_fono.to_excel(expertise, sheet_name='Фоноскопическая', index_label="ind")
                            df_study_fono = get_data_frame(file_name, "Фоно_иссл")
                            df_study_fono.to_excel(study, sheet_name='Фоноскопическая', index_label="ind")
                            df_exp_ling = get_data_frame(file_name, "Лингв_эксп")
                            df_exp_ling.to_excel(expertise, sheet_name='Лингвистическая', index_label="ind")
                            df_study_ling = get_data_frame(file_name, "Лингв_иссл")
                            df_study_ling.to_excel(study, sheet_name='Лингвистическая', index_label="ind")
                        except Exception as err:
                            print(Fore.RED + 'ОШИБКА В ЖУРНАЛЕ: ' + file_name)
                            print(err)
                            continue

                    elif direction_name == 'Экономическая':
                        try:
                            print(Fore.YELLOW + 'ПРЕОБРАЗОВАНИЕ ЖУРНАЛА: ' + file_name)
                            df = get_data_frame(file_name, "экспертизы")
                            df_exp_economy = df[df['ТИП'] == 'ЭКСПЕРТИЗА']
                            df_study_economy = df[df['ТИП'] == 'ИССЛЕДОВАНИЕ']
                            df_exp_economy.loc[df_exp_economy['КАТЕГОРИЯ'] == 'НАЛОГОВАЯ'].to_excel(expertise,
                                                                                                    sheet_name='Налоговая',
                                                                                                    index_label="ind")
                            df_exp_economy.loc[df_exp_economy['КАТЕГОРИЯ'] == 'БУХГАЛТЕРСКАЯ'].to_excel(expertise,
                                                                                                        sheet_name='Бухгалтерская',
                                                                                                        index_label="ind")
                            df_study_economy.loc[df_exp_economy['КАТЕГОРИЯ'] == 'НАЛОГОВАЯ'].to_excel(study,
                                                                                                      sheet_name='Налоговая',
                                                                                                      index_label="ind")
                            df_study_economy.loc[df_exp_economy['КАТЕГОРИЯ'] == 'БУХГАЛТЕРСКАЯ'].to_excel(study,
                                                                                                          sheet_name='Бухгалтерская',
                                                                                                          index_label="ind")
                        except Exception as err:
                            print(Fore.RED + 'ОШИБКА В ЖУРНАЛЕ: ' + file_name)
                            print(err)
                            continue

                    elif direction_name == 'ДНК':
                        try:
                            print(Fore.YELLOW + 'ПРЕОБРАЗОВАНИЕ ЖУРНАЛА: ' + file_name)
                            print(Fore.GREEN + 'ВЫЧИСЛЯЕТСЯ КОЛИЧЕСТВО ОБЪЕКТОВ')
                            df_dna_exp = get_data_frame(file_name, 'экспертизы')
                            df_dna_study = get_data_frame(file_name, 'исследования')
                            column_name = 'НОМЕР_ЭКСПЕРТИЗЫ'
                            df_dna_exp = df_dna_exp.replace(['NAN', np.nan], '0')
                            df_dna_exp = df_dna_exp.loc[df_dna_exp[column_name] != '0']
                            df_dna_exp['ВСЕГО_ОБЪЕКТОВ'] = np.nan
                            row_index = df_dna_exp['ВСЕГО_ОБЪЕКТОВ'].index
                            for i in row_index:
                                value_1 = df_dna_exp.iloc[i]['ОБЪЕКТЫ_КОЛИЧЕСТВО_ВВЕДЕННЫХ_В_ИССЛЕДОВАНИЕ']
                                value_2 = df_dna_exp.iloc[i]['КОЛИЧЕСТВО_ОБРАЗЦОВ_ЛИЦ']
                                try:
                                    value_sum = int(value_1) + int(value_2)
                                    df_dna_exp.loc[i, 'ВСЕГО_ОБЪЕКТОВ'] = str(value_sum)
                                except Exception as err:
                                    print(err)
                            df_dna_exp = df_dna_exp.rename(columns=column_names_dna)
                            df_dna_study = df_dna_study.rename(columns=column_names_dna)
                            df_dna_exp.to_excel(expertise, sheet_name=direction_name, index_label='ind')
                            df_dna_study.to_excel(study, sheet_name=direction_name, index_label='ind')
                        except Exception as err:
                            print(Fore.RED + 'ОШИБКА В ЖУРНАЛЕ: ' + file_name)
                            print(err)
                            continue

                    else:
                        try:
                            df_exp = get_data_frame(file_name, 'экспертизы')
                            df_exp.to_excel(expertise, sheet_name=direction_name, index_label='ind')
                        except Exception as err:
                            print(Fore.RED + 'ОШИБКА В ЖУРНАЛЕ ЭКСПЕРТИЗ: ' + file_name)
                            print(err)

                        try:
                            df_study = get_data_frame(file_name, 'исследования')
                            df_study.to_excel(study, sheet_name=direction_name, index_label='ind')
                        except Exception as err:
                            print(Fore.RED + 'ОШИБКА В ЖУРНАЛЕ ИССЛЕДОВАНИЙ: ' + file_name)
                            print(err)

    for file_name in name_list_convert:
        if "СКТЭ" in file_name:
            print(Fore.YELLOW + 'ПРЕОБРАЗОВАНИЕ ЖУРНАЛА СЛЕДСТВЕННЫХ ДЕЙСТВИЙ: ' + file_name)
            df_cons = get_data_frame(file_name, 'СД')
            df_cons.to_excel(consultation, sheet_name='Компьютерная', index_label='ind')

    expertise.close()
    study.close()
    consultation.close()
