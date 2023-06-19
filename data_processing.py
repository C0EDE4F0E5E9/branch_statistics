__author__ = 'Андрей Кравцов'
__version__ = '1.8.5'
__email__ = 'kravtsov911@gmail.com'

import hashlib
import pathlib
import re

import numpy as np
import pandas as pd
from colorama import Fore
from dateutil.parser import parse
from datetime import datetime, timezone
import error_list as err
from read_templates import templates


def get_key(d: dict, value):
    """Функция возвращает значение ключа в словаре"""
    for k, v in d.items():
        if v.upper() == value.upper():
            return k


def get_direction_id(d: dict, sheet_name: str) -> dict:
    """Функция получает ID направления, исходя из наименования листа в EXCEL"""
    idx = get_key(d, sheet_name)
    return idx


def get_region_id(path_excel_file: str) -> int:
    """Функция получает ID региона, исходя из наименования каталога с журналами"""
    for i in templates.region_id:
        if i in path_excel_file:
            region_id = templates.region_id.get(i)
            return region_id


def id_add(data_frame: object, direction_id='', region_id='') -> object:
    """Добавляет в DataFrame id направления и id региона"""
    if direction_id is None:
        data_frame['ID_РЕГИОНА'] = region_id
    else:
        data_frame['ID_РЕГИОНА'] = region_id
        data_frame['ID_НАПРАВЛЕНИЯ'] = direction_id
    return data_frame


def check_list(name_list: list, len_list: int) -> list:
    """Проверяет список подготовленный для загрузки в базу данных, на заполнение и если список пуст
    добавляет один список со значением -1"""
    if not name_list:
        tmp_ex = []
        for i in range(len_list):
            tmp_ex.append('-1')
        name_list.append(tmp_ex)
    return name_list


def get_data_frame(file_name: str, sheet_name: str) -> object:
    """Открытие листа Excel с преобразованием всех данных в тип str.

    Аргументы
    _________
    - file_name -- имя файла Excel содержащего листы
    - sheet_name -- имя листа в открытом файле Excel

    Возвращает
    __________
    - data_sheet -- возвращает Pandas DataFrame с данными в верхнем регистре

    """
    data = pd.ExcelFile(file_name)
    columns = data.parse(sheet_name).columns
    converters = {column: str for column in columns}
    data_sheet = data.parse(sheet_name, converters=converters)
    data_sheet = column_to_uppercase(data_sheet)

    return data_sheet


def folder_scan(path_: str, name_list: list) -> list:
    """Функция производит рекурсивный обход папок и возвращает список файлов
    с полными путями в указанном каталоге

    Аргументы
    _________
    - path_ -- путь к каталогу для рекурсивного обхода
    - name_list -- список в который будут заполняться сведения о путях

    Возвращает
    __________
    - name_list -- список заполненный путями к файлам, после рекурсии

    """
    try:
        for i in pathlib.Path(path_).glob('**/*'):
            if pathlib.Path(i).is_file():
                name_list.append(str(i))
        return name_list
    except PermissionError as err_0:
        print(err_0)
    except FileNotFoundError as err_1:
        print(err_1)
    except ValueError as err_2:
        print(err_2)


def column_to_uppercase(df: object) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, переводит наименования колонок и значения ячеек в верхний
    регистр, производит замену латинских символов в символы кириллицы и возвращает Pandas DataFrame

    Аргументы
    _________
    - df -- Pandas Data Frame

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
    """
    df.columns = df.columns.str.upper()
    columns_name = df.columns.values.tolist()
    for i in columns_name:
        try:
            df = df.astype({i: str})
            df[i] = df[i].str.upper()
            df[i] = df[i].str.strip()
        except AttributeError:
            continue
    df.columns = df.columns.str.replace(r'[C]', 'С', regex=True)
    df.columns = df.columns.str.replace(r'[E]', 'Е', regex=True)
    df.columns = df.columns.str.replace(r'[A]', 'А', regex=True)
    df.columns = df.columns.str.replace(r'[O]', 'О', regex=True)
    df.columns = df.columns.str.replace(r'[ ]', '_', regex=True)
    df.columns = df.columns.str.replace(r'[,]', '', regex=True)
    return df


def remove_empty_cell(df: object, column_name: str) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, обрезает строки, когда в указанном столбце начинается
    NaN и возвращает Pandas DataFrame.

    Аргументы
    _________
    - df -- Pandas Data Frame
    - column_name -- имя столбца, по которому будет происходить обрезание

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
    """

    values = ['NAN', 'NAT', '00:00:00']
    try:
        index = df.columns.get_loc(column_name)
        df = df.query(f'{df.columns[index]} not in {values}')
        df = df.replace(['NAT', 'NAN', '00:00:00', np.nan], '0')
        df = df.reset_index(drop=True)
        return df
    except ValueError:
        return df


def rename_columns(df: object, rename_column_list: dict, columns_name_map: dict) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, производит переименование колонок в универсальные имена.

    Аргументы
    _________
    - df -- Pandas Data Frame
    - rename_column_list -- должен быть словарем, где ключ это старое имя, а значение новое имя.
    - columns_name_map -- должен быть словарем, который содержит порядок столбцов для
     заполнения таблицы в базе данных

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
    """

    try:
        df.rename(rename_column_list, axis=1, inplace=True)
    except ValueError as val_err:
        print(val_err)
    # добавление недостающих столбцов из списка
    column_name_exist = df.columns.values.tolist()
    for column in columns_name_map.values():
        if column in column_name_exist:
            continue
        else:
            df[column] = np.nan
    # удаление лишних столбцов, которых нет в списке
    for column in column_name_exist:
        if column not in columns_name_map.values():
            df.drop(column, axis=1, inplace=True)
    # производит сортировку столбцов и приводит DataFrame в соответствие с картой столбцов
    sorted(columns_name_map.keys())
    df = df[columns_name_map.values()]
    return df


def rename_territoriality(df: object, name_list: dict) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, переименовывает подразделения по территориальности
    (при необходимости добавить в словарь значения для изменения) и возвращает Pandas DataFrame.
    По умолчанию используется столбец с именем 'ТЕРРИТОРИАЛЬНОСТЬ'.

    Аргументы
    _________
    - df -- Pandas Data Frame
    - name_list -- Словарь, где первое ключ старое имя, значение новое имя/

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
    """
    column_name = 'ТЕРРИТОРИАЛЬНОСТЬ'
    df = df.replace({column_name: name_list})
    return df


def row_process(df: object, column_name: str, name_list: dict) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, преобразует данные выбранном столбце в индекс из списка,

    Аргументы
    _________
    - df -- Pandas Data Frame
    - column_name -- имя столбца в котором нужно изменить значение на индекс
    - list_name -- словарь, который содержит список индексов

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
     """
    column_index = df.columns.get_loc(column_name)
    row_index = df[column_name].index
    for i in row_index:
        data_cell = df.iat[i, column_index]
        try:
            idx = get_key(name_list, data_cell)
            if idx is None:
                print(Fore.RED + err.error_message_7(df.iat[i, 2]))
                df.iat[i, column_index] = -1
            else:
                df.iat[i, column_index] = idx
        except ValueError:
            continue
    return df


def number_selection_cell(df: object, column_name='КОЛИЧЕСТВО_ОБЪЕКТОВ') -> object:
    """Принимает в качестве аргумента Pandas DataFrame, преобразует данные в столбе исключая
    из записей текст и оставляя только целое число

    Аргументы
    _________
    - df -- Pandas Data Frame
    - column_name -- имя столбца в котором нужно сделать выборку (по умолчанию "КОЛИЧЕСТВО_ОБЪЕКТОВ")

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
     """
    column_index = df.columns.get_loc(column_name)
    row_index = df[column_name].index
    for row in row_index:
        data_cell = df.iat[row, column_index]
        try:
            data_cell = str(data_cell)
            tmp = re.findall(r'(\d+)', data_cell)
            a = 0
            for i in tmp:
                a = a + int(i)
            df.iat[row, column_index] = a
            del tmp
        except ValueError:
            df.iat[row, column_index] = 0
    return df


def datetime_transformation(df: object, column_name: str) -> object:
    """Функция преобразует значение даты в ячейке в EPOCH-TIME UTC +0.

    Аргументы
    _________
    - df -- Pandas Data Frame
    - column_name -- имя столбца в котором содержится дата

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
     """
    df_rows = len(df.index)
    column_index = df.columns.get_loc(column_name)
    for i in range(df_rows):
        cell_value = df.iat[i, column_index]
        if column_name == 'ДАТА_ПОСТУПЛЕНИЯ' and cell_value == '0':
            print(Fore.RED + err.error_message_3(df.iat[i, 2]))
        elif cell_value == '0':
            continue
        else:
            try:
                dt = parse(cell_value)
                if dt.date() == datetime.now().date():
                    print(Fore.RED + err.error_message_1(df.iat[i, 2]))
                else:
                    timestamp = dt.replace(tzinfo=timezone.utc).timestamp()
                    df.iat[i, column_index] = int(timestamp)
            except ValueError:
                print(Fore.RED + err.error_message_2(df.iat[i, 2]))
    return df


def del_old_data(df: object, start_year=None) -> object:
    """Функция удаляет материалы выполненные позже отчетного года.

    Аргументы
    _________
    - df -- Pandas Data Frame
    - start_year -- начало отчетного года в формате YYYY-mm-dd

    Возвращает
    __________
    - df -- преобразованный Pandas Data Frame
    """

    column_name = 'ID_НАПРАВЛЕНИЯ'
    start_year = parse(start_year)
    start_year = int(start_year.replace(tzinfo=timezone.utc).timestamp())
    column_index = df[column_name].index
    delete_index = []
    check_col_name = df.columns
    if 'НОМЕР' in check_col_name:
        for i in column_index:
            date_receipt = int(df.iat[i, df.columns.get_loc('ДАТА_ПОСТУПЛЕНИЯ')])
            date_delivery = int(df.iat[i, df.columns.get_loc('ДАТА_СДАЧИ')])
            if date_delivery == 0 or date_delivery == '0':
                continue
            elif date_receipt > date_delivery:
                print(Fore.RED + err.error_message_4(df.iat[i, df.columns.get_loc('НОМЕР')]))
                # это условие для удаления строк которые выполнены до закрытия года в файле settings.ini
                if date_delivery < start_year:
                    delete_index.append(i)
                    print(Fore.BLUE + err.error_message_5(df.iat[i, df.columns.get_loc('НОМЕР')]))
            #  устанавливается дата на день позже закрытия года формате в Epoch timestamp GMT
            elif date_delivery < start_year:
                delete_index.append(i)
            elif date_delivery > start_year:
                continue
    if 'ВИД_ДЕЙСТВИЯ' in check_col_name:
        for i in column_index:
            date_action = int(df.iat[i, df.columns.get_loc('ДАТА_ДЕЙСТВИЯ')])
            if date_action < start_year:
                delete_index.append(i)
    if 'ДАТА_ПРОВЕДЕНИЯ' in check_col_name:
        for i in column_index:
            date_action = int(df.iat[i, df.columns.get_loc('ДАТА_ПРОВЕДЕНИЯ')])
            if date_action < start_year:
                delete_index.append(i)

    for i in delete_index:
        df = df.drop(index=[i])
    del_rec = str(len(delete_index))
    if delete_index:
        print(Fore.BLUE + err.error_message_6(str(del_rec)))
    df = df.reset_index(drop=True)
    return df


def hash_entry(df: object) -> object:
    """Функция добавляет значения HASH для выявления изменений в журналах"""
    column_index = df['НОМЕР'].index
    hash1_index = df.columns.get_loc('HASH_1')
    hash2_index = df.columns.get_loc('HASH_2')
    number_case_index = df.columns.get_loc('НОМЕР')
    date_receipt_index = df.columns.get_loc('ДАТА_ПОСТУПЛЕНИЯ')
    date_of_delivery_index = df.columns.get_loc('ДАТА_СДАЧИ')
    for i in column_index:
        str_number = df.iat[i, number_case_index]
        str_date_receipt = df.iat[i, date_receipt_index]
        str_date_delivery = df.iat[i, date_of_delivery_index]
        hash_1_str = (str(str_number) + str(str_date_receipt))
        hash_1_str = hash_1_str.replace(' ', '')

        # значение has_1 используется для проверки общего количества материалов
        hash_1 = hashlib.md5(hash_1_str.encode('utf-8')).hexdigest()
        df.iat[i, hash1_index] = hash_1
        if str_date_delivery != '0' and str_date_delivery != 0:
            hash_2_str = (str(str_number) + str(str_date_receipt) + str(str_date_delivery))
            hash_2_str = hash_2_str.replace(' ', '')

            # значение hash_2 используется для выявления изменений о датах поступления и сдачи материалов
            hash_2 = hashlib.md5(hash_2_str.encode('utf-8')).hexdigest()
            df.iat[i, hash2_index] = hash_2
    return df


def check_duration_verbatim(df: object) -> object:
    """Принимает в качестве аргумента Pandas DataFrame, преобразует данные в столбе
    "СУММАРНАЯ_ДЛИТЕЛЬНОСТЬ_ДОСЛОВКИ" строку в число с плавающей точкой (float)"""
    column_index = df.columns.get_loc('СУММАРНАЯ_ДЛИТЕЛЬНОСТЬ_ДОСЛОВКИ')
    row_index = df['СУММАРНАЯ_ДЛИТЕЛЬНОСТЬ_ДОСЛОВКИ'].index
    for row in row_index:
        data_cell = df.iat[row, column_index]
        for i in templates.format_date:
            try:
                data_cell = float(data_cell)
                df.iat[row, column_index] = data_cell
            except ValueError:
                const_date = datetime(1900, 1, 1)
                data = datetime.strptime(data_cell, i)
                data = (data - const_date).total_seconds() / 60
                df.iat[row, column_index] = data
    return df


def get_result_list(df: object, result_list: list) -> list:
    """Функция обрабатывает список списков и создает один список"""
    temp_list_0 = [df.values.tolist()]
    temp_list_1 = [item for sublist in temp_list_0 for item in sublist]
    for row in temp_list_1:
        result_list.append(row)
    if df.empty:
        print(Fore.GREEN + 'ЗАГРУЖНО ЗАПИСЕЙ-> ', str(0))
    else:
        print(Fore.GREEN + 'ЗАГРУЖНО ЗАПИСЕЙ-> ', str(len(df)))
    print(Fore.GREEN + '_______________________________________________________')
    temp_list_0.clear()
    temp_list_1.clear()
    return result_list


def data_frame_from_bd(name_list: list, columns_map: list) -> object:
    """Функция создает 'Pandas DataFrame' из списка полученного после обработки журналов EXCEL"""
    df = pd.DataFrame(name_list)
    df.rename(columns=columns_map, errors="raise", inplace=True)
    return df


def get_data_se(name_list: list, result_list: list) -> list:
    """Функция преобразует данные из журналов экспертиз и исследований в список для загрузки в БД.

    Аргументы
    _________
    - name_list -- список путей к журналу экспертиз

    Возвращает
    __________
    - result_list -- подготовленный список для загрузки в БД
    """
    for i in name_list:
        print(Fore.GREEN + 'ЗАГРУЗКА ЖУРНАЛА ', i.upper())
        region_id = get_region_id(i)
        df = pd.ExcelFile(i)
        for sheet in df.sheet_names:
            for ii in templates.regex_direction:
                if ii.findall(sheet):
                    name = str(sheet)
                    print(Fore.GREEN + 'ЗАГРУЗКА НАПРАВЛЕНИЯ -> ', name.upper())
                    df = get_data_frame(i, sheet)
                    df = rename_columns(df, templates.column_names_se, templates.columns_names_map_se)
                    df = rename_territoriality(df, templates.territory)
                    df = remove_empty_cell(df, 'НОМЕР')
                    df = datetime_transformation(df, 'ДАТА_ПОСТУПЛЕНИЯ')
                    df = datetime_transformation(df, 'ДАТА_СДАЧИ')
                    df = number_selection_cell(df, 'КОЛИЧЕСТВО_ОБЪЕКТОВ')
                    df = del_old_data(df, templates.start_year)
                    df = row_process(df, 'СТАТУС', templates.index)
                    df = row_process(df, 'ТИП_ПОСТАНОВЛЕНИЯ', templates.index_type_exp)

                    # раскомментировать после добавления сложности в журналы
                    #df = row_process(df, 'СЛОЖНОСТЬ', templates.index_complexity)

                    direction_id = get_direction_id(templates.sheet_name, name)
                    df = id_add(df, direction_id, region_id)
                    df = check_duration_verbatim(df)
                    df = hash_entry(df)
                    result_list = get_result_list(df, result_list)
                else:
                    continue
    return result_list


def get_data_inv(name_list: list, result_list: list) -> list:
    """Функция преобразует данные из журналов следственных действий в список для загрузки в БД.

    Аргументы
    _________
    - name_list -- список путей к журналу экспертиз

    Возвращает
    __________
    - result_list -- подготовленный список для загрузки в БД
    """
    for i in name_list:
        print(Fore.GREEN + 'ЗАГРУЗКА ЖУРНАЛА ', i.upper())
        region_id = get_region_id(i)
        df = pd.ExcelFile(i)
        for sheet in df.sheet_names:
            for ii in templates.regex_direction:
                if ii.findall(sheet):
                    name = str(sheet)
                    print(Fore.GREEN + 'ЗАГРУЗКА НАПРАВЛЕНИЯ -> ', name.upper())
                    df = get_data_frame(i, sheet)
                    df = rename_columns(df, templates.column_names_inv, templates.columns_names_map_inv)
                    df = rename_territoriality(df, templates.territory)
                    df = remove_empty_cell(df, 'ДАТА_ДЕЙСТВИЯ')
                    df = datetime_transformation(df, 'ДАТА_ДЕЙСТВИЯ')
                    df = number_selection_cell(df, 'КОЛИЧЕСТВО_ОБЪЕКТОВ')
                    df = row_process(df, 'ВИД_ДЕЙСТВИЯ', templates.index)
                    df = del_old_data(df, templates.start_year)
                    direction_id = get_direction_id(templates.sheet_name, name)
                    df = id_add(df, direction_id, region_id)
                    result_list = get_result_list(df, result_list)
                else:
                    continue
    return result_list


def get_data_cons(name_list: list, result_list: list) -> list:
    """Функция преобразует данные из журналов консультаций в список для загрузки в БД.

    Аргументы
    _________
    - name_list -- список путей к журналу экспертиз

    Возвращает
    __________
    - result_list -- подготовленный список для загрузки в БД
    """
    for i in name_list:
        print(Fore.GREEN + 'ЗАГРУЗКА ЖУРНАЛА ', i.upper())
        region_id = get_region_id(i)
        df = pd.ExcelFile(i)
        for sheet in df.sheet_names:
            if 'Иная_деятельность' in sheet:
                name = str(sheet)
                print(Fore.GREEN + 'ЗАГРУЗКА НАПРАВЛЕНИЯ -> ', name.upper())
                df = get_data_frame(i, sheet)
                df = rename_columns(df, templates.column_names_cons, templates.columns_names_map_consult)
                df = rename_territoriality(df, templates.territory)
                df = remove_empty_cell(df, 'ДАТА_ПРОВЕДЕНИЯ')
                df = datetime_transformation(df, 'ДАТА_ПРОВЕДЕНИЯ')
                df = row_process(df, 'ТИП', templates.index)
                df = row_process(df, 'ID_НАПРАВЛЕНИЯ', templates.sheet_name)
                df = id_add(df, None, region_id)
                df = del_old_data(df, templates.start_year)
                result_list = get_result_list(df, result_list)
            else:
                continue

    return result_list
