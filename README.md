# ОФУКБ: запуск Power Query из Excel на Python

В репозитории лежит минимальный набор файлов для запуска Python-версии Power Query из книги Excel.

## Файлы

- `wtf2_xmlsafe_cellonly.py` - Python-скрипт, который извлекает M-код Power Query из `.xlsx`, выполняет поддерживаемые запросы и записывает результат обратно в копию книги без изменения служебных частей Power Query.
- `test.xlsx` - тестовая Excel-книга с Power Query запросами, под которую сделан текущий скрипт.
- `README.md` - краткая инструкция по содержимому репозитория и запуску скрипта.

## Установка зависимостей

```bash
pip install pandas requests beautifulsoup4 lxml openpyxl
```

## Запуск

Запуск с регистрационным номером банка:

```bash
python3 wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000
```

Результат будет сохранен в файл:

```text
test_regnum_1000_python_filled.xlsx
```

Показать найденные загружаемые таблицы без выполнения запросов:

```bash
python3 wtf2_xmlsafe_cellonly.py test.xlsx --list
```

Запуск с подробным логом:

```bash
python3 wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000 --verbose
```
