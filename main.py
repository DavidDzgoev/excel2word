from docxtpl import DocxTemplate
import pandas as pd
import os

# проверяем, есть ли папка result
if os.path.exists('result'):
    # если она есть, программа выдаст ошибку
    raise Exception('Удалите папку result')

else:
    # создаем папку, если её нет
    os.mkdir('result')

# считываем наши данные, называя столбцы
data = pd.read_excel('input.xlsx', header=None, names=['фамилия', 'имя'])

# для каждой строки в таблицы выполняем следующие действия
for i in range(len(data)):
    # считываем имя и фамилию из таблицы
    first_name = data.iloc[i]['имя']
    last_name = data.iloc[i]['фамилия']

    doc = DocxTemplate("template.docx")  # берём шаблон ворд файла
    context = {'фамилия': last_name,  # вместо переменной фамилия, вставляем значение из таблицы
               'имя': first_name}  # вместо переменной имя, вставляем значение из таблицы

    doc.render(context)
    doc.save(f"result/{last_name}_{first_name[:1]}_{i}.docx")  # сохраняем новый файл с названием типа Иванов_И_228
