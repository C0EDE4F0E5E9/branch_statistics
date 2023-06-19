def error_message_1(text: str):
    msg1 = f'ОШИБКА (В ЯЧЕЙКЕ УСТАНОВЛЕНА ТЕКУЩАЯ ДАТА) --> {text}'
    return msg1


def error_message_2(text: str):
    msg2 = f'ОШИБКА (НЕИЗВЕСТНЫЙ ФОРМАТ ДАТЫ) --> {text}'
    return msg2


def error_message_3(text: str):
    msg3 = f'ОШИБКА (ЯЧЕЙКА "ДАТА_ПОСТУПЛЕНИЯ" НЕ ЗАПОЛНЕНА) --> {text}'
    return msg3


def error_message_4(text: str):
    msg4 = f'ОШИБКА (МАТЕРИАЛ ВЫПОЛНЕН РАНЬШЕ ЧЕМ ПОСТУПИЛ) --> {text}'
    return msg4


def error_message_5(text: str):
    msg5 = f'ПРЕДУПРЕЖДЕНИЕ (МАТЕРИАЛ ВЫПОЛНЕН ДО НАЧАЛА ОТЧЕТНОГО ГОДА, ЗАПИСЬ БУДЕТ УДАЛЕНА) --> {text}'
    return msg5


def error_message_6(text: str):
    msg6 = f'ПРЕДУПРЕЖДЕНИЕ (МАТЕРИАЛЫ НЕ ЗАГУРЖЕНЫ ТАК КАК ВЫПОЛНЕНЫ ДО НАЧАЛА ОТЧЕТНОГО ГОДА, ЗАПИСЕЙ БУДЕТ ' \
           f'УДАЛЕНО -> {text}) '
    return msg6


def error_message_7(text: str):
    msg7 = f'ОШИБКА (ИНДЕКС НЕ ПРИСВОЕН, УСТАНОВЛЕНО ЗНАЧЕНИЕ -1, ОШИБКА В ОДНОМ ИЗ СТОЛБЦОВ СТОЛБЦОВ "СТАТУС", ' \
           f'"ТИП_ПОСТАНОВЛЕНИЯ" ИЛИ "СЛОЖНОСТЬ") --> {text}'
    return msg7
