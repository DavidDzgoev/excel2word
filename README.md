### Description
Небольшой скрипт, который заполняет шаблон ворд файла данными.

**Перед началом стоит установить все зависимости:**

`pip install -r requirements.txt`

**Затем:**

1. Создадим шаблон ворд файла. Это должен быть обычный ворд файл, однако части текста, которые необходимо
   будет заменить в процессе работы скрипта, мы выделяем в двойные кавычки: `{{ имя }}` и `{{ фамилия }}`. Пример можно посмотреть в файле
   `template.docx`. Шаблон тоже обязательно должен так называться.
   
2. Создадим `input.xlsx` из наших данных. Обязательно **xlsx**. 
   В нем должна быть четкая таблица с нашими данными: первый столбец фамилии, второй имена. 
   Можно посмотреть пример как из `db.xls` я выделил необходимые данные и создал `input.xlsx`
   
3. Папки `result` не должно быть в репозитории, иначе программа выдаст ошибку.