import datetime as dt
import configparser
import PySimpleGUI as sg

from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook
from os.path import isfile, isdir


def date_from_log(date):
    '''
    Функция извлекает дату из записи даты-времени лога. В логе формат такой:
    yyyy-mm-ddThh:mm:ss.xxxZ
    '''
    result = date.split('T')[0]
    return result


def format_date(date):
    y, m, d = date.split('-')
    format_date = '.'.join([d, m, y])
    return format_date


def time_to_timezone(time):
    '''
    Функция приводит время к нужному часовому поясу.
    Т.к. не было необходимости проверять пограничные значения при смене дня -
    просто прибавляет значение TIME_ZONE ко времени.
    '''
    hour = str(int(time[:2]) + int(TIME_ZONE))
    if len(hour) < 2:
        hour = '0'+hour
    time = time.replace(time[:2], hour, 1)
    return time


def log_files(day, log_config):
    config = etree.parse(log_config)

    logs = []
    for elem in config.getiterator(tag='eventlog-fileset-entry'):
        path = elem.get('path')
        old = date_from_log(elem.get('oldest-record'))
        new = date_from_log(elem.get('newest-record'))
        logs.append((path, old, new))

    logs_for_jrnl = []
    for log in logs:
        if log[1] <= day <= log[2]:
            logs_for_jrnl.append(log[0])
    return logs_for_jrnl


def parse_log(log_name, day):
    full_name = LOG_CATALOG + log_name
    with open(full_name, 'r', encoding=ENCODING) as f:
        log = f.read()
    log_soup = BeautifulSoup(log, 'html.parser')

    input_set = set()
    output_set = set()

    tags = log_soup.find_all(type='INF')

    for tag in tags:
        time = tag['time']
        date = date_from_log(time)
        if date == day:
            # извлекаем имя файла
            filename = tag.find('wm-attachment-filename') or tag.find(
                'wm-file-path')
            if filename is None:
                continue

            # извлекаем время и корректируем часовой пояс
            time = time_to_timezone(time.split('T')[1][:-8])

            # форматируем дату
            date = format_date(date) + ' ' + time

            # извлекаем общие теги
            in_or_out = tag.find('wmap-rule-name').get_text().lower()
            address = tag.find('wm-user-name').get_text()

            # создаём кортеж значений и добавляем его в множество
            element = (date, address, filename.get_text())
            if 'входящие' in in_or_out:
                input_set.add(element)
            else:
                output_set.add(element)
    return (input_set, output_set)


def append_number_and_user(log, start_number):
    '''Функция добавляет в лог номер для каждого файла и имя формирующего'''
    update_log = []
    for element in log:
        update_log.append([start_number] + list(element) + [USERNAME])
        start_number += 1
    return update_log, start_number


def export_to_xls(name, header, log_list):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for element in log_list:
        ws.append(element)
    wb.save(name)


if __name__ == '__main__':
    # считываем настройки из settings
    print('Считываем настройки из settings.ini...')
    if not isfile('settings.ini'):
        raise Exception('Не найден settings.ini!')
    settings = configparser.ConfigParser()
    settings.read('settings.ini', encoding='utf-8')
    try:
        LOG_CATALOG = settings.get('Settings', 'log_catalog')
        JRNL_CATALOG = settings.get('Settings', 'jrnl_catalog')
        ENCODING = settings.get('Settings', 'encoding')
        START_INPUT_NUMBER = settings.get('Settings', 'start_input_number')
        START_OUTPUT_NUMBER = settings.get('Settings', 'start_output_number')
        USERNAME = settings.get('Settings', 'username')
        TIME_ZONE = settings.get('Settings', 'time_zone')
    except:
        raise Exception('В settings.ini не хватает полей! Необходимые поля: ',
                        'log_catalog, jrnl_catalog, encoding, username,'
                        'time_zone, start_input_number, start_output_number'
                        )

    # определяем дату формирования журнала (предыдущий рабочий день)
    today = dt.datetime.date(dt.datetime.today())
    week_day = dt.datetime.weekday(today)
    if week_day == 0:
        day = str(today - dt.timedelta(days=3))
    else:
        day = str(today - dt.timedelta(days=1))

    # создаём окно интерфейса
    layout = [
        [sg.Push(), sg.Text('Каталог log-файлов:'),
         sg.InputText(LOG_CATALOG), sg.FolderBrowse()],
        [sg.Push(), sg.Text('Каталог с журналами:'),
         sg.InputText(JRNL_CATALOG), sg.FolderBrowse()],
        [sg.Push(), sg.Text('Кодировка журналов:'),
         sg.InputText(ENCODING)],
        [sg.Push(), sg.Text('Ответственный:'),
         sg.InputText(USERNAME)],
        [sg.Push(), sg.Text('Часовой пояс:'),
         sg.InputText(TIME_ZONE)],
        [sg.Push(), sg.Text('Дата создаваемого журнала:'),
         sg.InputText(day)],
        [sg.Push(), sg.Text('Стартовый номер входящих фалов:'),
         sg.InputText(START_INPUT_NUMBER)],
        [sg.Push(), sg.Text('Стартовый номер исходящих фалов:'),
         sg.InputText(START_OUTPUT_NUMBER)],
        [sg.Output(size=(80, 10), key='-OUTPUT-')],
        [sg.Button('Сформировать'), sg.Button('Выход')]
    ]
    window = sg.Window(
        'Формирование журналов входящих/исхоящих файлов VipNet',
        layout,
        size=(590, 430)
    )

    config_change = False
    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Выход'):
            break
        if event == 'Сформировать':
            # очищаем окно вывода
            window['-OUTPUT-'].Update('')
            # проверяем заполненные поля формы и данные в settings
            if values[0] != LOG_CATALOG:
                if not isdir(values[0]):
                    print('НЕВЕРНО УКАЗАН КАТАЛОГ LOG-ФАЙЛОВ: такого '
                          'каталога не существует!')
                else:
                    LOG_CATALOG = values[0]
                    config_change = True
                    print('ВНИМАНИЕ: каталог Log-файлов изменён на '
                          f'{LOG_CATALOG}')
            if values[1] != JRNL_CATALOG:
                if not isdir(values[1]):
                    print('НЕВЕРНО УКАЗАН КАТАЛОГ С ЖУРНАЛАМИ: такого '
                          'каталога не существует!')
                else:
                    JRNL_CATALOG = values[1]
                    config_change = True
                    print('ВНИМАНИЕ: каталог с журналами изменён на '
                          f'{JRNL_CATALOG}')
            if values[2] != ENCODING:
                ENCODING = values[2]
                config_change = True
                print(f'ВНИМАНИЕ: кодировка изменена на {ENCODING}!'
                      'Если такой кодировки не существует, то программа'
                      'будет работать неправильно!')
            if values[3] != USERNAME:
                USERNAME = values[3]
                config_change = True
                print(f'ВНИМАНИЕ: ответственный изменён на {USERNAME}')
            if values[4] != TIME_ZONE:
                if (values[4][0] != '+' and values[4][0] != '-' or
                        len(values[4][1:]) > 2 or
                        not str(values[4][1:]).isdigit()):
                    print('ЧАСОВОЙ ПОЯС ВВЕДЁН НЕВЕРНО! Часовой пояс должен '
                          'начинаться со знака "+" или "-", и содержать '
                          'только цифры!')
                else:
                    config_change = True
                    TIME_ZONE = values[4]
                    print(f'ВНИМАНИЕ: часовой пояс изменён на {TIME_ZONE}')
            if values[5] != day:
                day = values[5]
                print('ВНИМАНИЕ: день создаваемого журнала изменён на '
                      f'{day}!')
            if values[6] != START_INPUT_NUMBER:
                if not str(values[6]).isdigit():
                    print('СТАРТОВЫЙ НОМЕР ВВЕДЁН НЕВЕРНО: он должен быть'
                          'числом!')
                else:
                    START_INPUT_NUMBER = values[6]
                    print('ВНИМАНИЕ: стартовый номер вхоядщих фалов изменён '
                          f'на {START_INPUT_NUMBER}!')
            if values[7] != START_OUTPUT_NUMBER:
                if not str(values[7]).isdigit():
                    print('СТАРТОВЫЙ НОМЕР ВВЕДЁН НЕВЕРНО: он должен быть'
                          'числом!')
                else:
                    START_OUTPUT_NUMBER = values[7]
                    print('ВНИМАНИЕ: стартовый номер исходящих фалов изменён '
                          f'на {START_OUTPUT_NUMBER}!')

            # ищем подходящие Log-файлы
            print(f'Будет сформирован журнал за {day}')
            log_config_file = LOG_CATALOG + 'wmail.cfg'
            logs = log_files(day, log_config_file)
            print('Используемые лог-файлы: ', *logs)
            # создаём 2 пустых множества: для входящих и исходящих файлов
            input_set = set()
            output_set = set()

            for name in logs:
                # парсим log-файл, результат пишем в 2 разных множества
                inp, out = parse_log(name, day)
                # добавляем полученные из log-файлов данные в общие множества
                input_set.update(inp)
                output_set.update(out)
            # преобразуем множества в списки и сортируем
            input_log = sorted(list(input_set))
            output_log = sorted(list(output_set))

            # добавляем номера записей и имя формирующего
            print('Создаём журналы...')
            input_log, START_INPUT_NUMBER = append_number_and_user(
                input_log, int(START_INPUT_NUMBER))
            output_log, START_OUTPUT_NUMBER = append_number_and_user(
                output_log, int(START_OUTPUT_NUMBER))

            input_header = [
                '№', 'Дата, время получения', 'Имя массива или имя файла',
                'Отправитель', 'ФИО, подпись лица, ответственного за получение'
            ]
            output_header = [
                '№', 'Дата, время отправки', 'Имя массива или имя файла',
                'Получатель', 'ФИО, подпись лица, ответственного за отправку'
            ]

            jrnl_input_name = JRNL_CATALOG + 'jnl-in_'\
                + format_date(day) + '.xlsx'
            jrnl_output_name = JRNL_CATALOG + 'jnl-out_'\
                + format_date(day) + '.xlsx'

            export_to_xls(jrnl_input_name, input_header, input_log)
            export_to_xls(jrnl_output_name, output_header, output_log)

            # заносим новые значения в settings.ini
            settings.set('Settings', 'start_input_number',
                         str(START_INPUT_NUMBER))
            settings.set('Settings', 'start_output_number',
                         str(START_OUTPUT_NUMBER))
            if config_change:
                settings.set('Settings', 'log_catalog', LOG_CATALOG)
                settings.set('Settings', 'jrnl_catalog', JRNL_CATALOG)
                settings.set('Settings', 'username', USERNAME)
                settings.set('Settings', 'time_zone', TIME_ZONE)
            with open('settings.ini', 'w', encoding='utf-8') as settings_file:
                settings.write(settings_file)

            print('Работа завершена!')
