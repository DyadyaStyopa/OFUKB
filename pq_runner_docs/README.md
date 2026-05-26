# Power Query XLSX Runner

`wtf2_xmlsafe_cellonly.py` выполняет Power Query запросы из файла Excel `.xlsx`
и записывает результаты обратно в копию книги.

Скрипт рассчитан на структуру файла `test.xlsx`: он извлекает M-код из
`customXml/item1.xml`, выполняет поддерживаемые операции Power Query на Python
и заполняет связанные Excel-таблицы.

## Зависимости

Установите Python-зависимости:

```bash
pip install pandas requests beautifulsoup4 lxml openpyxl
```

## Базовый запуск

Из рабочей директории проекта:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx
```

Результат будет сохранен рядом с исходным файлом:

```text
test_python_filled.xlsx
```

## Запуск для конкретного банка

Чтобы заменить регистрационный номер банка в URL ЦБ:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000
```

Результат:

```text
test_regnum_1000_python_filled.xlsx
```

## Использование HTML-кэша

По умолчанию скрипт использует папку:

```text
pq_html_cache
```

Можно указать другую папку:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000 --cache-dir pq_html_cache
```

Чтобы принудительно скачать HTML заново:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000 --no-cache
```

## Диагностика

Показать список найденных Power Query загрузок без выполнения:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --list
```

Сохранить извлеченный M-код:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --dump-m
```

Подробный лог:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000 --verbose
```

Максимальная диагностика с debug-файлами:

```bash
python3 pq_runner_docs/wtf2_xmlsafe_cellonly.py test.xlsx --regnum 1000 --debug
```

## Безопасность записи XLSX

Скрипт меняет только XML листов `xl/worksheets/sheet*.xml`.
Служебные части Power Query (`xl/tables`, `xl/queryTables`, `xl/connections.xml`,
`customXml/DataMashup`, `xl/workbook.xml`) должны остаться неизменными.

После сохранения выполняется проверка ZIP/XML структуры `.xlsx`. Если скрипт
обнаружит изменение служебных частей, он завершится с ошибкой.
