# ОФУКБ: запуск Power Query из Excel на Python

Этот репозиторий содержит Python-скрипт для выполнения части Power Query запросов из Excel-книги `test.xlsx` без запуска Excel.

Скрипт был сделан под структуру запросов в `test.xlsx`: он извлекает M-код из `.xlsx`, выполняет поддерживаемые операции Power Query на Python и записывает результат в новую копию книги.

## Что лежит в репозитории

- `OFUKB_CBR_PQ_alt_parser.py` - основной Python-скрипт. Выполняет Power Query запросы, скачивает или читает HTML-страницы ЦБ из кэша, собирает таблицы через `pandas` и записывает результат в `.xlsx` через ZIP/XML.
- `test.xlsx` - тестовая Excel-книга с Power Query запросами. Скрипт ориентирован именно на структуру этой книги.
- `pq_excel_gui.py` - графическая оболочка для запуска парсера без терминала.
- `assets/app_icon.icns` и `assets/app_icon.ico` - иконки приложения для сборки macOS `.app` и Windows `.exe`.
- `README.md` - эта инструкция.

## Что делает скрипт

1. Открывает `.xlsx` как ZIP-архив.
2. Извлекает M-код Power Query из `customXml/item1.xml`.
3. Находит запросы, которые загружаются в Excel-таблицы на листах книги.
4. При необходимости заменяет регистрационный номер банка в URL ЦБ через параметр `--regnum`.
5. Выполняет поддерживаемые операции Power Query на Python.
6. Записывает результаты в копию `.xlsx`.
7. Проверяет, что служебные части Power Query в книге не были испорчены.

Скрипт не является универсальным интерпретатором M. Он поддерживает тот набор операций, который используется в текущей книге: `Web.Page`, `Web.Contents`, `Web.BrowserContents`, `Html.Table`, `Table.PromoteHeaders`, `Table.TransformColumnTypes`, `Table.ReplaceErrorValues`, `Table.RemoveColumns`, `Table.RenameColumns`, `Table.Skip`, `Table.NestedJoin`, `Table.ExpandTableColumn`, `Table.ReorderColumns`, `Table.TransformColumns`.

## Требования

Нужен Python 3. Обязательные Python-зависимости для backend-скрипта:

```bash
pip install pandas requests beautifulsoup4 lxml
```

Для графического интерфейса используется `tkinter`. На macOS и Windows он обычно входит в стандартную установку Python. На Linux может потребоваться системный пакет, например `python3-tk`.

Опциональные зависимости:

```bash
pip install openpyxl pyinstaller
```

- `openpyxl` можно использовать для дополнительной ручной проверки, что итоговая книга открывается как Excel-файл. Сам скрипт записывает `.xlsx` напрямую через ZIP/XML и не требует `openpyxl` для основной работы.
- `pyinstaller` нужен только для сборки GUI в `.app` или `.exe`; для обычного запуска из Python он не требуется.


## Графический запуск

Для запуска через окно используйте GUI:

```bash
python3 pq_excel_gui.py
```

В окне можно выбрать `test.xlsx`, указать `regnum`, путь выходного файла, папку HTML-кэша, лог и debug-папку. GUI запускает backend в отдельном процессе и показывает вывод выполнения в окне.

Текущая архитектура GUI рассчитана на будущую упаковку в приложение: backend вызывается как Python-модуль, а не как отдельная команда через внешний интерпретатор Python.

## Нативное macOS-приложение

В репозитории также есть новая SwiftUI-оболочка для macOS в папке `Sources/OFUKBMacGUI`. Она использует тот же backend `OFUKB_CBR_PQ_alt_parser.py`, но выглядит как обычное macOS-приложение: выбор Excel-файла через Finder-панель или drag-and-drop, отдельные блоки для `regnum`, результата, режимов запуска и лога.

Собрать и открыть приложение локально:

```bash
./script/build_and_run.sh
```

Проверить сборку и запуск процесса:

```bash
./script/build_and_run.sh --verify
```

Собрать `.app` без открытия окна:

```bash
./script/build_and_run.sh --build-only
```

После сборки приложение лежит здесь:

```text
dist/OFUKB CBR PQ.app
```

Важно: macOS-приложение всё равно требует доступный `python3` и Python-зависимости из `requirements.txt`, потому что бизнес-логика остается в Python backend.

## Упаковка без терминала/IDE

### macOS: сборка `.app`

Для macOS основной вариант - нативная SwiftUI-оболочка:

```bash
./script/build_and_run.sh --build-only
```

После сборки открывать нужно файл:

```text
dist/OFUKB CBR PQ.app
```

Проверка из терминала:

```bash
open "dist/OFUKB CBR PQ.app"
```

Внутрь `.app` копируются SwiftUI GUI, `OFUKB_CBR_PQ_alt_parser.py`, `requirements.txt` и иконки. Python и зависимости из `requirements.txt` должны быть установлены на компьютере, где запускается приложение.

Если нужен именно установщик `.pkg`, сначала соберите `.app`, а затем сделайте package отдельной командой:

```bash
productbuild --component "dist/OFUKB CBR PQ.app" /Applications "dist/OFUKB CBR PQ.pkg"
```

Для распространения `.app` или `.pkg` другим пользователям macOS может потребоваться подпись приложения, а для публичной доставки - notarization.

### Windows: сборка `.exe`

```bash
python -m pip install pandas requests beautifulsoup4 lxml pyinstaller
python -m PyInstaller --noconsole --onedir --name "OFUKB CBR PQ" --icon assets/app_icon.ico --add-data "OFUKB_CBR_PQ_alt_parser.py;." pq_excel_gui.py
```

Результат будет в папке:

```text
dist\OFUKB CBR PQ\OFUKB CBR PQ.exe
```

Для распространения пользователям обычно удобнее `--onedir`, потому что зависимости `pandas`, `lxml`, `requests` и файлы Tcl/Tk для GUI проще диагностировать в распакованной папке. Поверх собранной папки можно сделать установщик через Inno Setup или NSIS.

Сборки не подписаны. На macOS возможны предупреждения Gatekeeper, на Windows - предупреждения SmartScreen.

## Быстрый запуск

Из корня репозитория:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000
```

После выполнения появится файл:

```text
test_regnum_1000_python_filled.xlsx
```

`1000` - регистрационный номер банка. Его можно заменить на другой номер:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1481
```

Тогда результат будет сохранен как:

```text
test_regnum_1481_python_filled.xlsx
```

## Запуск без замены regnum

Если не передавать `--regnum`, скрипт использует регистрационный номер, который уже прописан в M-коде книги:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx
```

Результат:

```text
test_python_filled.xlsx
```

## Явное имя выходного файла

Можно указать путь для результата через `-o` или `--output`:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 -o result.xlsx
```

## Просмотр найденных запросов

Чтобы только посмотреть, какие Power Query запросы будут выполнены и куда они загружаются:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --list
```

Этот режим не создает новый `.xlsx`.

## Сохранить извлеченный M-код

Чтобы выгрузить M-код из книги в отдельный файл:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --dump-m
```

Будет создан файл рядом с книгой:

```text
test.Section1.m
```

## HTML-кэш

Скрипт получает данные с сайта ЦБ из HTML-страниц. По умолчанию он использует папку:

```text
pq_html_cache
```

Если HTML для нужного URL уже есть в кэше, повторная загрузка из интернета не выполняется.

Можно указать другую папку кэша:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --cache-dir path/to/cache
```

Чтобы игнорировать кэш и скачать HTML заново:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --no-cache
```

## Логи и debug

Подробный лог в консоль и файл:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --verbose
```

Указать конкретный файл лога:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --verbose --log-file run.log
```

Максимальная диагностика:

```bash
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --debug
```

В debug-режиме скрипт сохраняет промежуточные CSV/JSON-снимки шагов Power Query и диагностический отчет по XML/ZIP структуре итоговой книги.

## Безопасность записи XLSX

Главная цель текущей версии - не ломать Excel-книгу при записи результата.

Скрипт меняет только XML листов:

```text
xl/worksheets/sheet*.xml
```

Он не должен менять служебные части Power Query:

```text
xl/tables/*
xl/queryTables/*
xl/connections.xml
customXml/item1.xml
xl/workbook.xml
```

После сохранения скрипт проверяет ZIP/XML структуру `.xlsx` и дополнительно проверяет, что служебные части остались byte-for-byte такими же, как в исходном файле. Если это условие нарушено, выполнение завершается ошибкой.


## Добавление новых Power Query запросов

Скрипт может сработать с дополнительными Power Query запросами, если они добавлены в `test.xlsx` или в книгу с очень похожей внутренней структурой.

Новые запросы с высокой вероятностью будут обработаны, если:

- запрос загружается на лист Excel как таблица, а не остается только connection-only;
- M-код хранится в DataMashup внутри `customXml/item*.xml`;
- запрос использует уже поддержанные операции Power Query: `Web.Page`, `Web.Contents`, `Web.BrowserContents`, `Html.Table`, `Table.PromoteHeaders`, `Table.TransformColumnTypes`, `Table.ReplaceErrorValues`, `Table.RemoveColumns`, `Table.RenameColumns`, `Table.Skip`, `Table.NestedJoin`, `Table.ExpandTableColumn`, `Table.ReorderColumns`, `Table.TransformColumns`;
- HTML-страницы имеют структуру таблиц, похожую на уже используемые страницы ЦБ;
- выходная Excel-таблица уже создана в книге и связана с Power Query/queryTable.

Новый запрос, скорее всего, потребует доработки backend-скрипта, если он использует другие M-функции, например `Table.SelectRows`, `Table.AddColumn`, `Table.Group`, `Table.Combine`, `Table.Unpivot`, `Excel.CurrentWorkbook`, `Csv.Document`, работу с JSON/XML/API, пользовательские M-функции, параметры или сложные вложенные `let`-выражения.

Важно: скрипт не является универсальным Power Query runtime. Он реализует Python-исполнитель конкретного семейства запросов из текущей книги. Также безопасная запись не создает новые Excel-таблицы с нуля и не меняет `xl/tables` / `xl/queryTables`; она рассчитана на обновление данных в уже существующих таблицах книги.

## Типичный сценарий работы

1. Открыть терминал в корне репозитория.
2. Установить зависимости.
3. Запустить скрипт с нужным `--regnum`.
4. Проверить созданный файл `*_python_filled.xlsx` в Excel.
5. При проблемах запустить повторно с `--debug` и посмотреть debug-папку.

Пример:

```bash
pip install pandas requests beautifulsoup4 lxml
python3 OFUKB_CBR_PQ_alt_parser.py test.xlsx --regnum 1000 --verbose
```
