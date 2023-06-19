from sys import exit

from colorama import init, Fore

import converter
import data_processing
import sql_to_excel
from read_templates import templates

print('ВВЕДИТЕ ПУТЬ К ЖУРНАЛАМ: ', end='')
general_folder = input()

print('ДАТА НАЧАЛА ОТЧЕТНОГО ГОДА ->', templates.start_year)
print('ДАТА НАЧАЛА ОТЧЕТНОГО ПЕРИОДА ->', templates.start_date)
print('ДАТА КОНЦА ОТЧЕТНОГО ПЕРИОДА ->', templates.end_date)
print('ПАРАМЕТРЫ ИЗМЕНЯЮТСЯ В ФАЙЛЕ С ИМЕНЕМ -> settings.json')
print('ПРОДОЛЖИТЬ ВЫПОЛНЕНИЕ С ВЫБРАННЫМИ ПАРАМЕТРАМИ ПЕРИОДОВ? 1 = ДА, 2 = НЕТ ', end='')

while True:
    result = input()
    try:
        result = int(result)
        if result == 1:
            break
        elif result == 2:
            exit()
        else:
            print('ВВЕДЕН НЕ СУЩЕСТВУЮЩИЙ ПАРАМЕТР -> ', result)
    except ValueError:
        print('ВВЕДЕНЫ НЕ ВЕРНЫЕ ПАРАМЕТРЫ -> ', result)

temp_files_folder = []  # список для сохранения путей файлов
expertise = []  # список журналов экспертиз
study = []  # список журналов исследований
consultations = []  # список журналов консультаций
investigative_actions = []  # список журналов следственных действий
result_E_list = []  # данные по экспертизам готовые для загрузки в базу
result_S_list = []  # данные по исследованиям готовые для загрузки в базу
result_I_list = []  # данные по следственным действиям готовые для загрузки в базу
result_C_list = []  # данные по консультациям готовые для загрузки в базу
templates.db_name = sql_to_excel.create_actual_db()  # генерируется новая база данных, на основе отчетного периода

converter.convert(general_folder)
data_processing.folder_scan(general_folder, temp_files_folder)
for i in temp_files_folder:
    if '~$' in i:
        continue
    elif 'свыше 2-х' in i:
        continue
    elif 'экспертиз' in i:
        expertise.append(i)
    elif 'исследований' in i:
        study.append(i)
    elif 'консультаций' in i:
        consultations.append(i)
    elif 'следственных' in i:
        investigative_actions.append(i)
    else:
        continue

temp_files_folder.clear()
init(autoreset=True)
print(Fore.GREEN + 'ОБРАБОТКА ЖУРНАЛОВ ЭКСПЕРТИЗ')
data_processing.get_data_se(expertise, result_E_list)
data_processing.check_list(result_E_list, len(templates.columns_names_map_se))
templates.result_E_list = result_E_list

init(autoreset=True)
print(Fore.GREEN + 'ОБРАБОТКА ЖУРНАЛОВ ИССЛЕДОВАНИЙ')
data_processing.get_data_se(study, result_S_list)
data_processing.check_list(result_S_list, len(templates.columns_names_map_se))
templates.result_S_list = result_S_list

init(autoreset=True)
print(Fore.GREEN + 'ОБРАБОТКА ЖУРНАЛОВ СЛЕДСТВЕННЫХ ДЕЙСТВИЙ')
data_processing.get_data_inv(investigative_actions, result_I_list)
data_processing.check_list(result_I_list, len(templates.columns_names_map_inv))
templates.result_I_list = result_I_list

init(autoreset=True)
print(Fore.GREEN + 'ОБРАБОТКА ЖУРНАЛОВ КОНСУЛЬТАЦИЙ')
data_processing.get_data_cons(consultations, result_C_list)
data_processing.check_list(result_C_list, len(templates.columns_names_map_consult))
templates.result_C_list = result_C_list

sql_to_excel.convert_sql_to_excel()

print('DONE')
