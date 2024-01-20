import openpyxl
import os
from datetime import datetime
from forex4you import forex4you_scrap
from lifefinance import lifefinance_scrap
from signalstart import signalstart_scrap
from myfxbook_strategies import myfxbook_strategies_scrap
from bybit import bybit_scrap
from myfxbook_members import myfxbook_members_scrap
import linecache
import sys


def print_exception():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print('⛔ ОШИБКА\nПерешли мне текст ошибки: ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))


def run_main(date):
    try:
        print(f"🚀 Процесс считывания запущен."f"Установлена дата: по {date}.\n")
        stop_date = datetime.strptime(date, '%d.%m.%Y')
        dict_for_traders = {'Объем': [],
                            'Валютная пара': [],
                            'Тип сделки': [],
                            'Время Открытия': [],
                            'Цена Открытия': [],
                            'Время Закрытия': [],
                            'Цена Закрытия': [],
                            'Прибыль': [],
                            'Ссылка': [],
                            }
        months_in_numbers = {"янв.": "01",
                             "февр.": "02",
                             "мар.": "03",
                             "апр.": "04",
                             "мая": "05",
                             "июня": "06",
                             "июля": "07",
                             "авг.": "08",
                             "сент.": "09",
                             "окт.": "10",
                             "нояб.": "11",
                             "дек.": "12",
                             }
        current_dir = os.path.dirname(os.path.abspath(__file__))
        input_excel = openpyxl.load_workbook(fr'{current_dir}\resources\input\input hrefs.xlsx')
        sheet = input_excel['Лист1']
        list_of_input_hrefs = sheet['A']
        for href in list_of_input_hrefs:
            if href.value is None:
                continue
            elif 'forex4you' in href.value:
                forex4you_scrap(current_dir, href, months_in_numbers, stop_date, dict_for_traders)
            elif 'litefinance' in href.value:
                lifefinance_scrap(current_dir, href, stop_date, dict_for_traders)
            elif 'myfxbook' in href.value:
                if 'strategies' in href.value:
                    myfxbook_strategies_scrap(current_dir, href, stop_date, dict_for_traders)
                else:
                    myfxbook_members_scrap(current_dir, href, stop_date, dict_for_traders)
            elif 'signalstart' in href.value:
                signalstart_scrap(current_dir, href, stop_date, dict_for_traders)
            elif 'bybit.com' in href.value:
                bybit_scrap(current_dir, href, stop_date, dict_for_traders)
            elif 'mql5' in href.value:
                pass
        print(f'✅ Работа успешно завершена!\n')
    except Exception as e:
        print_exception()


# run_main('01.07.2023')
