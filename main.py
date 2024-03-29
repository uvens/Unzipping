import os
import zipfile

import openpyxl
import pandas as pd
import numpy as np
import re
from zipfile import BadZipfile
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook




def get_data():
    data = {
        'file_name': None,
        'station': None,
        'buyer': None,
        'seller': None

    }
    return data

import pandas as pd

# def dataframe_to_dict(df,file_name):
#     dictionary = {}
#     for index, row in df.iterrows():
#         for column in df.columns:
#             cell_value = row[column]
#             cell_coords = f'{index}:{column}'
#             dictionary[cell_coords] = cell_value
#     dictionary['file_name'] = file_name
#     return dictionary

# def dataframe_to_dict(df, file_name):
#     dictionary = {}
#     for cell in df.columns:
#         for index, value in df[cell].items():
#             cell_coords = f'{index}:{cell}'
#             dictionary[cell_coords] = value
#     dictionary['file_name'] = file_name
#     return dictionary

def dataframe_to_dict(df, file_name):
    dictionary = {}
    # df = df.apply(lambda x: pd.Series(x.dropna().values))
    df = df.dropna(how='all')
    for i, (index, row) in enumerate(df.iterrows(), start=1):
        for j, cell_value in enumerate(row, start=1):
            cell_coords = f'{i}:{j}'
            if cell_value is None:
                continue
            dictionary[cell_coords] = cell_value
    dictionary['file_name'] = file_name
    return dictionary


import pandas as pd

def write_to_excel(dataframe, file_path):
    if not os.path.exists(file_path):
        dataframe.to_csv(file_path, index=False)
        print(f"Файл {file_path} создан и данные записаны в него.")
    else:
        dataframe.to_csv(file_path, mode='a', index=False, header=False)
        print(f"Данные добавлены в файл CSV: {file_path}")


def dict_to_excel(dictionary, file_name='output.csv', sheet_name='Data'):
    # Создание DataFrame из словаря
    df = pd.DataFrame(list(dictionary.items()), columns=['Coordinates', 'Value'])

    try:
    # Создание новых столбцов для разделенных данных
        df[['Row', 'Column']] = df['Coordinates'].str.split(':', expand=True)

        df.drop(columns=['Coordinates'], inplace=True)

        # Дублирование имени файла и добавление в последний столбец
        df['File Name'] = dictionary.pop('file_name')
        df = df[df['Row'] != 'file_name']
    except ValueError:
        l = ''
    # Проверка наличия файла Excel
    write_to_excel(df,file_name)




def process_directory(directory,data):
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.xlsx') or file.endswith('.xls'):
                # Если найден Excel файл, читаем его содержимое
                print(f"Найден файл Excel: {file_path}")
                df = pd.read_excel(file_path)
                # Обработка содержимого файла df (DataFrame) по вашим потребностям
                # В этом примере сохранение в CSV файл, как и делается в process_zip_archive_in_memory
                file_name = re.sub(r'[\\/*?:"<>|]', '_', file_path)
                # write_invoice_in_data(df)
                dict_to_excel(dataframe_to_dict(df, file_name))
                try:
                    root_file = '/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files'
                    check_file = os.path.join(root_file, f'{file_name}.xlsx')
                    if not os.path.isfile(check_file):
                        df.to_excel(f'/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files/{file_name}.xlsx',
                                    index=False)
                    else:
                        continue
                except BadZipfile as ex:
                    print(f'{ex} Error with file = {file_name}')
                    continue
            elif file.endswith('.zip'):
                # Если найден архив, рекурсивно обрабатываем его содержимое
                print(f"Найден архив: {file_path}")
                process_zip_archive(file_path,data)
            else:
                # Добавьте здесь дополнительные условия для других типов файлов или действия, если нужно
                print(f"Найден файл: {file_path}")
                # Ваша логика обработки других типов файлов или пропуск (pass), если необходимо

        for directory in dirs:
            dir_path = os.path.join(root, directory)
            # Обрабатываем вложенные директории рекурсивно
            process_directory(dir_path,data)


def process_zip_archive(zip_file,data):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith('.xlsx') or file.endswith('.xls'):
                # Если в архиве найден Excel файл, читаем его содержимое
                print(f"Найден файл Excel в архиве: {file}")
                with zip_ref.open(file) as excel_file:
                    df = pd.read_excel(excel_file)
                    # Обработка содержимого файла df (DataFrame) по вашим потребностям
                    # В этом примере сохранение в CSV файл, как и делается в process_zip_archive_in_memory
                    file_name = re.sub(r'[\\/*?:"<>|]', '_', excel_file.name)
                    dict_to_excel(dataframe_to_dict(df, file_name))
                    try:
                        root_file = '/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files'
                        check_file = os.path.join(root_file, f'{file_name}.xlsx')
                        # write_invoice_in_data(df)
                        if not os.path.isfile(check_file):
                            df.to_excel(
                                f'/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files/{file_name}.xlsx',
                                index=False)
                        else:
                            continue
                    except BadZipfile as ex:
                        print(f'{ex} Error with file = {file_name}')
                        continue
            elif file.endswith('.zip'):
                # Если в архиве найден вложенный архив, рекурсивно обрабатываем его содержимое
                print(f"Найден вложенный архив в архиве: {file}")
                with zip_ref.open(file) as inner_zip_file:
                    process_zip_archive_in_memory(inner_zip_file,data)


def process_zip_archive_in_memory(zip_content,data):
    with zipfile.ZipFile(zip_content) as in_memory_zip:
        for file in in_memory_zip.namelist():
            if file.endswith('.xlsx') or file.endswith('.xls'):
                # Если в памяти архива найден Excel файл, читаем его содержимое
                print(f"Найден файл Excel во вложенном архиве: {file}")
                with in_memory_zip.open(file) as excel_file:
                    df = pd.read_excel(excel_file)
                    # Обработка содержимого файла df (DataFrame) по вашим потребностям
                    # В этом примере сохранение в CSV файл, как и делается в process_zip_archive_in_memory
                    file_name = re.sub(r'[\\/*?:"<>|]', '_', excel_file.name)
                    dict_to_excel(dataframe_to_dict(df, file_name))
                    try:
                        root_file = '/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files'
                        check_file = os.path.join(root_file, f'{file_name}.xlsx')
                        # write_invoice_in_data(df)
                        if not os.path.isfile(check_file):
                            df.to_excel(
                                f'/home/uventus/Works/Разархивация/Архивы_инвойсов/Excel_Files/{file_name}.xlsx',
                                index=False)
                        else:
                            continue
                    except BadZipfile as ex:
                        print(f'{ex} Error with file = {file_name}')
                        continue
            elif file.endswith('.zip'):
                # Если в памяти архива найден вложенный архив, рекурсивно обрабатываем его содержимое
                print(f"Найден вложенный архив во вложенном архиве: {file}")
                with in_memory_zip.open(file) as inner_zip_content:
                    process_zip_archive_in_memory(inner_zip_content,data)


# Пример использования:
data = []
input_directory = '/home/uventus/Works/test_microservice'
process_directory(input_directory,data)
# df = pd.read_excel('/home/uventus/Works/Разархивация/Архивы_инвойсов/inv-spe Tongfu Orange Sia 07.08_TKRU4202180.xlsx')
# dataframe_to_dict(df,'/home/uventus/Works/Разархивация/Архивы_инвойсов/inv-spe Tongfu Orange Sia 07.08_TKRU4202180.xlsx')