import pandas as pd
import os
import xlrd

path_to_folder = 'C:\\Users\\IMatveev\\Desktop\\ЛКЕ\\index_fedstat (1)'

files = [f for f in os.listdir(path_to_folder) if f.endswith('.xls')]

all_data = []

for file in files:
    file_path = os.path.join(path_to_folder, file)
    xls = pd.ExcelFile(file_path)
    sheet = xls.parse('Паспорт')

    headers = sheet.iloc[1:15, 0].tolist()
    values = sheet.iloc[1:15, 1].tolist()

    # Если это первый файл, добавляем заголовки в итоговый список
    if not all_data:
        all_data.append(['Имя файла'] + headers)

    # Добавляем значения из файла в итоговый список
    all_data.append([file] + values)

# Создаем DataFrame из списка
df_all = pd.DataFrame(all_data[1:], columns=all_data[0])

# Сохраняем в файл
df_all.to_excel('итоговый_файл.xlsx', index=False)