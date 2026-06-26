# python_calamine (Pure Python, 3.9+)

Чистая Python-реализация API библиотеки [python-calamine](https://github.com/dimastbk/python-calamine).  
Работает на Python 3.9+, без Rust/C-расширений — просто положи папку `python_calamine/` в каталог проекта.

## Установка

```
pip install openpyxl xlrd
```

Для ODS дополнительных зависимостей **не нужно** (используется встроенный `zipfile` + `xml.etree`).

| Формат | Зависимость |
|--------|-------------|
| `.xlsx`, `.xlsm` | `openpyxl` |
| `.xls` | `xlrd >= 2.0` |
| `.ods` | *(нет)* |

## Быстрый старт

```python
from python_calamine import CalamineWorkbook, load_workbook

# Через фабричные методы
wb = CalamineWorkbook.from_path("book.xlsx")
wb = CalamineWorkbook.from_filelike(open("book.xlsx", "rb"))
wb = CalamineWorkbook.from_object("book.ods")   # авто-определение типа

# Или через удобную функцию
wb = load_workbook("book.xls")

# Контекстный менеджер
with CalamineWorkbook.from_path("book.xlsx") as wb:
    sheet = wb.get_sheet_by_name("Sheet1")
    data = sheet.to_python()          # список списков
    for row in sheet.iter_rows():     # итератор
        print(row)
```

## API

### `CalamineWorkbook`

```python
wb.path                      # str | None  — путь к файлу (None для filelike)
wb.sheet_names               # List[str]
wb.sheets_metadata           # List[SheetMetadata]
wb.table_names               # List[str]  — требует load_tables=True

wb.get_sheet_by_name(name)   # -> CalamineSheet
wb.get_sheet_by_index(idx)   # -> CalamineSheet
wb.get_table_by_name(name)   # -> CalamineTable  (только xlsx + load_tables=True)
wb.close()
```

### `CalamineSheet`

```python
sheet.name                   # str
sheet.height                 # int — число строк с данными
sheet.width                  # int — ширина самой широкой строки
sheet.total_height           # height - 1
sheet.total_width            # width - 1
sheet.start                  # tuple[int, int] | None
sheet.end                    # tuple[int, int] | None
sheet.merged_cell_ranges     # list[(start, end)] | []

sheet.to_python(skip_empty_area=True, nrows=None)  # List[List[...]]
sheet.iter_rows()            # Iterator[List[...]]
```

### `CalamineTable` (только xlsx)

```python
table.name      # str
table.sheet     # str — имя листа
table.columns   # List[str] — заголовки
table.height    # int
table.width     # int
table.start / end
table.to_python()  # данные без строки заголовков
```

### Типы значений ячеек

| Python-тип | Когда |
|------------|-------|
| `str` | Текст или пустая ячейка (`""`) |
| `float` | Числа (включая целые) |
| `bool` | Логические значения |
| `datetime.date` | Дата без времени |
| `datetime.datetime` | Дата со временем |
| `datetime.timedelta` | Длительность (ODS `time`) |

### Именованные таблицы xlsx

```python
wb = CalamineWorkbook.from_path("book.xlsx", load_tables=True)
print(wb.table_names)
tbl = wb.get_table_by_name("Table1")
print(tbl.columns)
print(tbl.to_python())
```

### Исключения

```python
from python_calamine import (
    CalamineError,      # базовый класс
    PasswordError,      # файл защищён паролем
    WorksheetNotFound,  # лист не найден
    TableNotFound,      # таблица не найдена
    WorkbookClosed,     # книга закрыта
    TablesNotLoaded,    # load_tables=True не был указан
    TablesNotSupported, # формат не поддерживает таблицы (xls, ods)
    XmlError,           # ошибка XML
    ZipError,           # ошибка ZIP-архива
)
```

## Структура файлов

```
your_project/
├── python_calamine/
│   └── __init__.py       ← вся реализация в одном файле
├── main.py
└── ...
```

## Отличия от оригинала

| Функция | Оригинал | Эта реализация |
|---------|----------|----------------|
| Python 3.9+ | ✅ | ✅ |
| xlsx / xlsm | ✅ | ✅ |
| xls | ✅ | ✅ |
| ods | ✅ | ✅ |
| Именованные таблицы (xlsx) | ✅ | ✅ |
| `merged_cell_ranges` (xlsx) | ✅ | ✅ |
| `merged_cell_ranges` (read_only) | — | возвращает `[]` |
| Скорость на больших файлах | Быстрее (Rust) | Медленнее |
| Зависимости | нет (wheel) | openpyxl, xlrd |
