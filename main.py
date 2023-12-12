import os
import shutil
import zipfile
import glob

def extract_zip(file_path, extract_to):
    """
    Разархивировать файл ZIP.

    Args:
    - file_path: путь к файлу ZIP
    - extract_to: путь для извлечения содержимого архива
    """
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

def process_directory(directory, output_folder):
    """
    Обработать содержимое директории.

    Args:
    - directory: путь к директории
    - output_folder: путь для сохранения файлов Excel
    """
    for i,item in enumerate(os.listdir(directory),1):
        item_path = os.path.join(directory, item)
        if os.path.isdir(item_path):
            # Если элемент - это директория
            process_directory(item_path, output_folder)  # Рекурсивно обрабатываем содержимое директории
        elif os.path.isfile(item_path):
            # Если элемент - это файл
            if item.endswith('.zip'):
                # Если файл - это ZIP-архив, разархивируем его
                extracted_folder = os.path.join(output_folder, str(i))  # Папка для извлечения содержимого архива
                os.makedirs(extracted_folder, exist_ok=True)
                extract_zip(item_path, extracted_folder)
                process_directory(extracted_folder, output_folder)  # Обрабатываем содержимое разархивированной папки
            elif item.endswith('.xls') or item.endswith('.xlsx'):
                # Если файл Excel, копируем его в папку output_folder
                output_path = os.path.join(output_folder, item)
                shutil.copyfile(item_path, output_path)
                print(f"Файл Excel скопирован в {output_path}")

# Пример использования:

directory_path = '/home/uventus/Works/Разархивация/Ноябрь'
output_directory = os.path.join(directory_path, 'Excel_Files')  # Создание папки для сохранения файлов Excel
os.makedirs(output_directory, exist_ok=True)

process_directory(directory_path, output_directory)
