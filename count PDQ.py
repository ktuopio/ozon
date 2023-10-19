import os
import pandas as pd

# Указываем путь к папке с файлами
folder_path = r'C:\'

# Получаем список всех файлов в указанной папке
file_list = os.listdir(folder_path)

# Проходим по каждому файлу
for file_name in file_list:
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)

        # Загружаем файл Excel в DataFrame
        df = pd.read_excel(file_path, engine='openpyxl')

        # Создаем столбец PDQ и вычисляем значения
        df['PDQ'] = df.iloc[:, 5:].count(axis=1) / (df.shape[1] - 5)

        # Сохраняем обновленный DataFrame обратно в файл
        df.to_excel(file_path, index=False, engine='openpyxl')

print("Завершено")
