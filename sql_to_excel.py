__author__ = 'Андрей Кравцов'
__version__ = "1.0"
__email__ = "kravtsov911@gmail.com"

import sqlite3, pathlib, shutil
import pandas as pd
from data_processing import folder_scan
from read_templates import templates


def create_actual_db():
    """Создает новую базу данных, в которую будут сводиться данные в отчетный период"""
    template_bd = pathlib.Path(templates.db_path)
    db_name = (f'{templates.end_date}.db'.replace('-', '.'))
    result_db = pathlib.Path('result_database\\' + db_name)
    shutil.copy(str(template_bd), str(result_db))
    return result_db


def create_data_frame_from_bd(name_list: list, columns_map: dict) -> object:
    """Функция создает 'Pandas DataFrame' из списка полученного после обработки журналов EXCEL"""
    df = pd.DataFrame(name_list)
    df = df.rename(columns=columns_map)
    return df


def convert_sql_to_excel():
    """Первый параметр База данных SQLite3, второй параметр дата начала периода, третий параметр дата конца периода"""
    sql_list = []
    folder_scan(templates.SQL_path, sql_list)
    sql_list.sort()
    result_ex = creat_data_frame_from_bd(templates.result_E_list, templates.columns_names_map_se)
    result_st = creat_data_frame_from_bd(templates.result_S_list, templates.columns_names_map_se)
    result_inv = creat_data_frame_from_bd(templates.result_I_list, templates.columns_names_map_inv)
    result_consult = creat_data_frame_from_bd(templates.result_C_list, templates.columns_names_map_consult)

    with pd.ExcelWriter(
            'result_queries\\' + 'Статистика c ' + templates.start_date + ' по ' + templates.end_date + '.xlsx',
            engine='openpyxl') as ex_res:
        try:
            connection = sqlite3.connect(templates.db_name)
            cursor = connection.cursor()
            print('СОЕДИНЕНИЕ С SQLITE УСТАНОВЛЕНО')
            connection.execute('DROP TABLE IF EXISTS tb_expertise')
            connection.execute('DROP TABLE IF EXISTS tb_study')
            connection.execute('DROP TABLE IF EXISTS tb_investigation')
            connection.execute('DROP TABLE IF EXISTS tb_consultation')
            result_ex.to_sql('tb_expertise', connection, if_exists='append')
            result_st.to_sql('tb_study', connection, if_exists='append')
            result_inv.to_sql('tb_investigation', connection, if_exists='append')
            result_consult.to_sql('tb_consultation', connection, if_exists='append')
            for sql in sql_list:
                file_name = pathlib.Path(sql).stem
                if '!' in sql:
                    continue
                with open(sql, 'r', encoding='utf-8') as f:
                    query = f.read()
                    query = query.replace('$first_date', templates.start_date)
                    query = query.replace('$second_date', templates.end_date)
                    query = query.replace('$start_date', templates.start_year)
                cursor.execute(query)
                result_query = cursor.fetchall()
                column_names = []
                for i in cursor.description:
                    column_names.append(i[0])
                df = pd.DataFrame(result_query, columns=column_names)
                df.to_excel(ex_res, sheet_name=file_name)

        except sqlite3.Error as error:
            print('ОШИБКА ПРИ ПОДКЛЮЧЕНИИ К SQLITE', error)
            print('НЕ ВЫПОЛНЕН ЗАПРОС', str(sql))

        finally:
            if connection:
                connection.close()
                print('СОЕДИНЕНИЕ С SQLITE ЗАКРЫТО')
