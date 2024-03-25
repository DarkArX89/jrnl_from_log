# Journal_from_Log

**Journal_from_Log** - это приложение для формирование журналов входящих/исходящих файлов программы "ViPNet Деловая почта". Приложение считывает настройки для журналов из файла "settings.ini", находит и парсит log-файлы, а затем формирует 2 журнала в формате ".xlsx". Реализован пользовательский интерфейс при помощи PySimpleGUI. Приложение разрабатывалось для конкретной цели и ежедневно используется в работе. 

## Использованные технологии:
- Python 3.9
- Beautiful Soup 4
- PySimpleGUI 
- lxml
- openpyxml

## Инструкция по запуску
Клонировать репозиторий и перейти в него в командной строке:

```
git clone https://github.com/yandex-praktikum/jrnl_from_log.git
```

```
cd jrnl_from_log
```

Cоздать и активировать виртуальное окружение:

```
python3 -m venv env
```

```
source env/bin/activate
```

Установить зависимости из файла requirements.txt:

```
python3 -m pip install --upgrade pip
```

```
pip install -r requirements.txt
```

#### В файле "settings.ini" задать настройки:
расположение log-файлов программы "ViPNet Деловая почта":
```
log_catalog = C:/wmail/
```
каталог для создаваемых журналов:
```
jrnl_catalog = C:/wmail/
```
кодировка log-файлов (по умолчанию стоит utf-16, т.к. именно она используется в log-файлах ViPNet):
```
encoding = utf-16
```
ФИО ответственного за создание журналов:
```
username = User U.U.
```
часовой пояс:
```
time_zone = +5
```
Стартовые номера файлов (по умолчанию, оба равны 1):
```
start_input_number = 1
start_output_number = 1
```
Запустить приложение скриптом jrnl_from_log.bat или командой

```
python jrnl_from_log.py
```
