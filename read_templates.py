__author__ = 'Андрей Кравцов'
__version__ = "1.0.1"
__email__ = "kravtsov911@gmail.com"

import csv
import json
import re
import sys
from pathlib import Path

from colorama import init, Fore


class Settings:
    """Класс содержит даты в качестве атрибутов.

    Attributes
    __________

    - start_year -- Начало отчетного года
    - start_date -- Начало отчетного периода
    - end_date -- Конец отчетного периода
    - templates_path -- Путь к каталогу с шаблонами
    - SQL_path -- Каталог с шаблонами SQL запросов
    - db_path -- Каталог с шаблоном базы данных SQLite 3
    - reference_files -- Список шаблонов

    """

    def __init__(self, settings='settings.json'):
        self.settings = settings
        self.start_year = str
        self.start_date = str
        self.end_date = str
        self.templates_path = str
        self.SQL_path = str
        self.db_path = str
        self.reference_files = []

        try:
            with open(settings, 'r') as f:
                config = json.load(f)
                setattr(self, 'start_year', config['start_year'])
                setattr(self, 'start_date', config['start_date'])
                setattr(self, 'end_date', config['end_date'])
                setattr(self, 'templates_path', config['templates_path'])
                setattr(self, 'reference_files', config['reference_files'])
                setattr(self, 'SQL_path', config['SQL_path'])
                setattr(self, 'db_path', config['db_path'])
        except FileNotFoundError:
            print(Fore.RED + 'ОТСУТСТВУЕТ ФАЙЛ КОНФИГУРАЦИИ "SETTINGS.INI" \n'
                             'НАЖМИТЕ ЛЮБУЮ КНОПКУ ДЛЯ ВЫХОДА ИЗ ПРОГРАММЫ')
            input()
            sys.exit()


class ReadTemplates(Settings):
    """Класс принимает путь к templates и заполняет аттрибуты.

        Attributes
        __________

        - column_names_cons: Словарь изменения имен столбцов для журналов консультаций.

        - column_names_es: Словарь изменения имен столбцов для журналов экспертиз и исследований.

        - column_names_inv: Словарь изменения имен столбцов для журналов следственных действий.

        - columns_name_map_consult: Словарь наименований столбцов для журналов консультаций.

        - columns_name_map_inv: Словарь наименований столбцов для журналов следственных действий.

        - columns_name_map_se: Словарь наименований столбцов для экспертиз и исследований.

        - column_names_dna: Словарь для переименования ДНК журнала.

        - index: Словарь индексов, которые указываются в столбцах
                 для журналов экспертиз и исследований "статус экспертизы","статус исследования"
                 для журналов следственных действий "вид деятельности"
                 для журналов консультаций "тип деятельности".

         - index_complexity: Словарь индексов категорий сложности экспертиз.

         - index_type_exp: Словарь индексов, которые указываются в столбцах "тип экспертизы", "тип исследований".

         - match_table_branch: Словарь имен журналов в Ессентуках.

         - region_id: Словарь регионов.

         - sheet_name: Словарь направлений для получения идентификатора направления, при необходимости добавляются
                       названия новых направлений.

         - territory: Словарь приведения к одному стандарту записей в столбце Территориальность.

         - regex_direction: Список направлений и листов в файлах EXCEL,
                            при необходимости добавляются названия новых направлений.
         - format_date: Список форматов времени, для преобразования длительности фонограммы.

         - start_year -- Начало отчетного года

         - start_date -- Начало отчетного периода

         - end_date -- Конец отчетного периода

         - templates_path -- Путь к каталогу с шаблонами

         - reference_files -- Список шаблонов

    """
    column_names_cons = {}
    column_names_se = {}
    column_names_inv = {}
    columns_names_map_consult = {}
    columns_names_map_inv = {}
    columns_names_map_se = {}
    column_names_dna = {}
    index = {}
    index_complexity = {}
    index_type_exp = {}
    match_table_branch = {}
    region_id = {}
    sheet_name = {}
    territory = {}
    regex_direction = []
    format_date = []

    def __init__(self):
        super().__init__()
        if not Path(str(self.templates_path)).is_dir():
            print(Fore.RED + 'ОТСУТСТВУЕТ КАТАЛОГ ШАБЛОНОВ "TEMPLATES" \n'
                             'НАЖМИТЕ ЛЮБУЮ КНОПКУ ДЛЯ ВЫХОДА ИЗ ПРОГРАММЫ')
            input()
            sys.exit()
        else:
            file_list = sorted(Path(str(self.templates_path)).glob('*.csv'))
            check_list = []
            for i in file_list:
                check_list.append(i.name)
            for i in self.reference_files:
                if i not in check_list:
                    print(Fore.RED + 'ШАБЛОН:' + i + ' ОТСУТСТВУЕТ. ДОБАВЬТЕ ШАБЛОН В КАТАЛОГ "TEMPLATES" \n'
                                                     'НАЖМИТЕ ЛЮБУЮ КНОПКУ ДЛЯ ВЫХОДА ИЗ ПРОГРАММЫ')
                    input()
                    sys.exit()
            for i in file_list:
                if '#' in i.name:
                    continue
                elif i.name not in self.reference_files:
                    print(Fore.RED + 'ШАБЛОН:' + i.name + ' ОТСУТСТВУЕТ. ДОБАВЬТЕ ШАБЛОН В КАТАЛОГ "TEMPLATES" \n'
                                                          'НАЖМИТЕ ЛЮБУЮ КНОПКУ ДЛЯ ВЫХОДА ИЗ ПРОГРАММЫ')
                    input()
                    sys.exit()
                else:
                    continue
            for i in file_list:
                with open(i, 'r', encoding='utf-8', newline='') as file:
                    reader = csv.reader(file, delimiter=';')
                    if 'dict_' in i.resolve().stem:
                        if 'map' in i.resolve().stem:
                            result = {int(rows[0]): rows[1] for rows in reader}
                        else:
                            result = {rows[0]: rows[1] for rows in reader}
                        sorted(result.items())
                        setattr(self, i.resolve().stem.replace('dict_', ''), result)
                    elif 'list_' in i.resolve().stem:
                        result = []
                        for sublist in list(reader):
                            for item in sublist:
                                result.append(re.compile(item))
                        setattr(self, i.resolve().stem.replace('list_', ''), result)
                    else:
                        continue


templates = ReadTemplates()
