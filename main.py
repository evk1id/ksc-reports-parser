"""
    Скрипт, который преобразует отчет 'Status devices.xls' (отчет KSC)
    в несколько CSV-файлов, разбитых по проблемам
"""

import csv
import sys
import pandas as pd
import re

def parse_excel_to_dict_list(filepath, sheet_name='Details'):
    """

        Данная функция читает xls-файл, возвращая данные в виде словаря

    :param filepath: путь к отчету KSC для чтения
    :param sheet_name: наименование листа xls-файла
    :return: список, содержащий сведения из отчета
    """
    df = pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
    dict_list = df.to_dict(orient='records')
    return dict_list

if __name__ == '__main__':
    # Инициализация
    try:
        filename = sys.argv[1]
    except IndexError:
        exit("USAGE: python3 main.py <file>")

    # Получение данных из отчета KSC
    data = parse_excel_to_dict_list(filename)

    with open('Прочее.csv', 'w', newline='') as f1, \
         open('Базы устарели.csv', 'w', newline='') as f2, \
         open('Давно не выполнялся поиск вредоносного ПО.csv', 'w', newline='') as f3, \
         open('Обнаружены активные угрозы.csv', 'w', newline='') as f4, \
         open('Защита выключена.csv', 'w', newline='') as f5, \
         open('Обнаружено много вирусов.csv', 'w', newline='') as f6, \
         open('Программа безопасности не установлена.csv', 'w', newline='') as f7, \
         open('Устройство давно не подключалось к Серверу администрирования.csv', 'w', newline='') as f8, \
         open('Устройство стало неуправляемым.csv', 'w', newline='') as f9:

        # Заголовки столбцов
        fieldnames = ['Статус', 'Виртуальный Сервер администрирования', 'Группа', 'Устройство', 'Последнее подключение к Серверу администрирования', 'Причина', 'Статус устройства определен программой', 'IP-адрес', 'Последнее появление в сети', 'Windows-домен', 'NetBIOS-имя', 'DNS-имя', 'DNS-домен', 'Операционная система', 'Дата выпуска антивирусной базы', 'Последняя полная проверка']
        files = [f1, f2, f3, f4, f5, f6, f7, f8, f9]
        writers = [csv.DictWriter(f, fieldnames=fieldnames) for f in files]

        # Записываем заголовки в файл
        for w in writers:
            w.writeheader()

        # Фильтруем данные по "Причине"
        for row in data:
            flag = False  # Для того, чтобы уже записанные строчки не записать еще в файл 'Прочее.csv'
            if re.search(r'\bБазы\s+устарели[\.!?]?\b', row.get('Причина', [])):
                writers[1].writerow(row)
                flag = True
            if  re.search(r'\bДавно\s+не\s+выполнялся\s+поиск\s+вредоносного\s+ПО[\.!?]?\b', row.get('Причина', [])):
                writers[2].writerow(row)
                flag = True
            if re.search(r'\bОбнаружены\s+активные\s+угрозы[\.!?]?\b', row.get('Причина', [])):
                writers[3].writerow(row)
                flag = True
            if re.search(r'\bЗащита\s+выключена[\.!?]?\b', row.get('Причина', [])):
                writers[4].writerow(row)
                flag = True
            if re.search(r'\bОбнаружено\s+много\s+вирусов[\.!?]?\b', row.get('Причина', [])):
                writers[5].writerow(row)
                flag = True
            if re.search(r'\bПрограмма\s+безопасности\s+не\s+установлена[\.!?]?\b', row.get('Причина', [])):
                writers[6].writerow(row)
                flag = True
            if re.search(r'\bУстройство\s+давно\s+не\s+подключалось\s+к\s+Серверу\s+администрирования[\.!?]?\b', row.get('Причина', [])):
                writers[7].writerow(row)
                flag = True
            if re.search(r'\bУстройство\s+стало\s+неуправляемым[\.!?]?\b', row.get('Причина', [])):
                writers[8].writerow(row)
                flag = True
            if not flag:
                writers[0].writerow(row)