import os
import time
from selenium import webdriver
import openpyxl
import io
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException
import re
import logging


class Drover:
    def __init__(self, current_dir,base_folder):
        self.current_dir = current_dir
        self.base_folder = base_folder

    client_exe = "your_client_exe_path"
    currency_list = ["EURUSD", "GOLD", "GBPUSD", "USDJPY"]
    days_to_check = 6


class Scraper:

    def __init__(self,
                 current_dir,
                 bd_dir,
                 driver,
                 href,
                 site_name,
                 xpathes_dict
                 ):
        self.current_dir = current_dir
        self.bd_dir = bd_dir
        self.driver = driver
        self.href = href
        self.dict_for_traders = {'Объем': [],
                                 'Валютная пара': [],
                                 'Тип сделки': [],
                                 'Время Открытия': [],
                                 'Цена Открытия': [],
                                 'Время Закрытия': [],
                                 'Цена Закрытия': [],
                                 'Прибыль': [],
                                 'Ссылка': [],
                                 }
        self.months_in_numbers = {"янв.": "01",
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
        self.site_name = site_name
        self.xpathes_dict = xpathes_dict

    def site_open(self):
        print(f'Перехожу по ссылке трейдера: {self.href.value}\n')
        self.driver.get(self.href.value)
        print(f'Успешно перешел по ссылке {self.href.value}\n')
        time.sleep(2)
        try:
            WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located(
                    ("xpath", self.xpathes_dict['trader_name']))
            )
        except:
            logging.error(
                f'Проверьте, существует ли трейдер! '
                f'Не смог найти имя трейдера по ссылке {self.href.value}'
            )
            return None
        name = remove_special_chars(self.driver.find_element(
            "xpath",
            self.xpathes_dict['trader_name']).text)
        print(f'Имя трейдера = {name}\n')
        excel_name = fr'{self.bd_dir}\{self.site_name}\excel\{name}.xlsx'
        htm_name = fr'{self.bd_dir}\{self.site_name}\htm\{name}.htm'
        df_for_trader = pd.DataFrame(self.dict_for_traders)
        df_for_trader_old = pd.DataFrame(self.dict_for_traders)
        last_row_for_compair = ''
        if os.path.isfile(excel_name):
            file_exist = True
            df_for_trader_old = pd.read_excel(excel_name, dtype=str)
            last_row_for_compair = [
                str(df_for_trader_old.iloc[0]['Время Закрытия']),
                float(remove_special_chars(
                    df_for_trader_old.iloc[0]['Цена Закрытия']))]
        else:
            file_exist = False
        return {
            'name': name,
            'excel_name': excel_name,
            'htm_name': htm_name,
            'file_exist': file_exist,
            'df_for_trader': df_for_trader,
            'df_for_trader_old': df_for_trader_old,
            'last_row_for_compair': last_row_for_compair
        }

    def site_scrap(self, init_dict):
        pass

    def excel_save(self, init_dict, scrap_results):
        scrap_results.to_excel(init_dict['excel_name'], sheet_name='Sheet1',
                               index=False)
        # Дальнейший код нужен для красивого форматирования колонок в excel
        wb = openpyxl.load_workbook(init_dict['excel_name'])
        ws = wb.active
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                finally:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
        wb.save(init_dict['excel_name'])
        print(
            f'✅ Успешно сформировал excel '
            f'с сигналами трейдера {init_dict["name"]}\n'
        )

    def htm_save(self, init_dict, scrap_results):
        text = ''
        for index, row in scrap_results.iterrows():
            text = text + \
                   '<tr align = right>' \
                   fr'<td>{row["Объем"]}</td>' \
                   fr'<td nowrap>{row["Время Открытия"]}</td>' \
                   fr'<td>{row["Тип сделки"]}</td>' \
                   fr'<td class=mspt>0</td>' \
                   fr'<td>{row["Валютная пара"]}</td>' \
                   fr'<td style="mso-number-format:0\.00000;">{row["Цена Открытия"]}</td>' \
                   fr'<td style="mso-number-format:0\.00000;">{row["Объем"]}</td>' \
                   fr'<td style="mso-number-format:0\.00000;">0</td>' \
                   fr'<td class=msdate nowrap>{row["Время Закрытия"]}</td>' \
                   fr'<td style="mso-number-format:0\.00000;">{row["Цена Закрытия"]}</td>' \
                   fr'<td class=mspt>0</td>' \
                   fr'<td class=mspt>0</td>' \
                   fr'<td class=mspt>0</td>' \
                   fr'<td class=mspt>0</td>' \
                   '</tr>'
        with io.open(fr'{self.current_dir}\resources\template1.htm', 'r',
                     encoding='utf-8') as f:
            html_string = f.read()
        htm = html_string.replace("SSSSS", text)
        with io.open(init_dict['htm_name'], 'w', encoding='utf-8') as file:
            file.write(htm)
        print(
            f'✅ Успешно сформировал htm с сигналами трейдера {init_dict["name"]}\n')

    def scrap_all(self):
        init_dict = self.site_open()
        if init_dict is None:
            return
        scrap_results = self.site_scrap(init_dict)
        self.excel_save(init_dict, scrap_results)
        self.htm_save(init_dict, scrap_results)


class Litefinance(Scraper):
    def site_scrap(self, init_dict):
        count_on_page = 1
        for o in (range(1, 11)):
            time.sleep(2)
            count = len(self.driver.find_elements("xpath", fr'//div[@class = "content_row"]'))
            print(f'Начинаю обработку c {count_on_page} по {count} запись на странице {o}\n')

            for l in range(count_on_page, count + 1):
                currency = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]/descendant::a[2]'
                ).text
                if currency is None:
                    continue
                date_close = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][3]'
                ).text
                date_close = datetime.strptime(
                    date_close,
                    '%d.%m.%Y %H:%M:%S'
                ) + timedelta(hours=-1)
                date_open = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][2]'
                ).text
                date_open = datetime.strptime(
                    date_open,
                    '%d.%m.%Y %H:%M:%S'
                ) + timedelta(hours=-1)
                type_of_trade = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][4]'
                ).text.lower()
                if type_of_trade == 'покупка':
                    type_of_trade = 'buy'
                else:
                    type_of_trade = 'sell'
                obj = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][5]'
                ).text.replace(".", ",")
                currency = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]/descendant::a[2]'
                ).text.replace("XAUUSD", "GOLD")
                price_open = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][6]'
                ).text.replace(" ", "")
                price_close = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][7]'
                ).text.replace(" ", "")
                points = self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{l}]'
                    fr'/descendant::div[@class = "content_col"][8]'
                ).text.replace(".", ",")
                new_row_to_compare = [
                    date_close.strftime('%Y.%m.%d %H:%M'),
                    float(remove_special_chars(price_close))]
                if (new_row_to_compare == init_dict['last_row_for_compair']
                        and init_dict['file_exist']):
                    init_dict['df_for_trader'] = pd.concat(
                        [init_dict['df_for_trader'],
                         init_dict['df_for_trader_old']],
                        ignore_index=True
                    )
                    return init_dict['df_for_trader']
                else:
                    new_row = pd.Series({
                        'Объем': obj,
                        'Валютная пара': currency,
                        'Тип сделки': type_of_trade,
                        'Время Открытия': date_open.strftime(
                            '%Y.%m.%d %H:%M'),
                        'Цена Открытия': price_open,
                        'Время Закрытия': date_close.strftime(
                            '%Y.%m.%d %H:%M'),
                        'Цена Закрытия': price_close,
                        'Прибыль': points,
                        'Ссылка': self.href.value,
                    })
                    init_dict['df_for_trader'].loc[
                        len(init_dict['df_for_trader'])] = new_row
                count_on_page += 1
            # Если количечество сделок меньше 50 на странице - остановить обработку
            if count % 50 != 0:
                break
            self.driver.execute_script(
                "arguments[0].scrollIntoView();",
                self.driver.find_element(
                    "xpath",
                    fr'(//div[@class = "content_row"])[{count}]')
            )
        return init_dict['df_for_trader']


class Forex4you(Scraper):

    def site_scrap(self, init_dict):
        try:
            element = WebDriverWait(self.driver, 1).until(ec.presence_of_element_located(("xpath", r"//div[contains(text(), 'Allow All')]")))
            element.click()
        except:
            pass
        self.driver.find_element(
            "xpath",
            fr'//label[contains(text(), "Весь период")]'
        ).click()
        for o in (range(1, 25)):
            time.sleep(2)
            count = 0
            while count == 0:
                count = len(self.driver.find_elements(
                    "xpath",
                    fr'//tbody//tr[@data-ng-repeat = '
                    fr'"trade in $fxGrid.$data track by trade.id"]'
                    fr'//td[@data-ng-bind="::trade.symbol"]')
                )
            print(f'Начинаю обработку {count} '
                  f'записей на странице {o}\n'
                  )
            if count > 20:
                count = 20
            for l in list(range(1, count+1)):
                try:
                    currency = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                    ).text
                    if currency is None:
                        continue
                    date_close = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        '//../preceding-sibling::td[1]'
                    ).text
                    for i in self.months_in_numbers:
                        date_close = date_close.replace(
                            i, self.months_in_numbers[i]
                        )
                    date_close = datetime.strptime(
                        date_close,
                        '%d %m %Y г., %H:%M:%S'
                    )
                    date_open = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../preceding-sibling::td[2]'
                    ).text
                    for i in self.months_in_numbers:
                        date_open = date_open.replace(
                            i,
                            self.months_in_numbers[
                                i]
                        )
                    date_open = datetime.strptime(
                        date_open,
                        '%d %m %Y г., %H:%M:%S'
                    )
                    type_of_trade = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../following-sibling::td[1]'
                    ).text.lower()
                    obj = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../preceding-sibling::td[3]').text.replace(
                        ".", ","
                    )
                    currency = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                    ).text.replace("XAUUSD", "GOLD")
                    price_open = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../following-sibling::td[2]').text.replace(
                        " ",
                        ""
                    )
                    price_close = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../following-sibling::td[3]').text.replace(
                        " ",
                        ""
                    )
                    points = self.driver.find_element(
                        "xpath",
                        fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                        fr'//../following-sibling::td[4]'
                    ).text.replace(".", ",")
                    new_row_to_compare = [
                        date_close.strftime('%Y.%m.%d %H:%M'),
                        float(remove_special_chars(price_close))]
                    if (new_row_to_compare == init_dict['last_row_for_compair']
                            and init_dict['file_exist']):
                        init_dict['df_for_trader'] = pd.concat(
                            [init_dict['df_for_trader'],
                             init_dict['df_for_trader_old']],
                            ignore_index=True
                        )
                        return init_dict['df_for_trader']
                    else:
                        new_row = pd.Series({
                            'Объем': obj,
                            'Валютная пара': currency,
                            'Тип сделки': type_of_trade,
                            'Время Открытия': date_open.strftime(
                                '%Y.%m.%d %H:%M'),
                            'Цена Открытия': price_open,
                            'Время Закрытия': date_close.strftime(
                                '%Y.%m.%d %H:%M'),
                            'Цена Закрытия': price_close,
                            'Прибыль': points,
                            'Ссылка': self.href.value,
                        })
                        init_dict['df_for_trader'].loc[
                            len(init_dict['df_for_trader'])] = new_row
                except:
                    pass
            if count % 10 != 0:
                break
            try:
                self.driver.find_element(
                    "xpath",
                    fr'(//a[@data-fx-grid-set-page="$fxGridPaginator.getNextPage()"])[1]'
                ).click()
            except:
                break
        return init_dict['df_for_trader']


def make_hrefs_list(hrefs_file):
    input_excel = openpyxl.load_workbook(hrefs_file)
    sheet = input_excel['Лист1']
    list_of_input_hrefs = sheet['A']
    return list_of_input_hrefs


def remove_special_chars(string):
    pattern = r'[^\w\s]'
    return re.sub(pattern, '', string)


def run_main():
    try:
        logging.basicConfig(
            level=logging.ERROR,
            filename='main.log',
            datefmt='%d.%m.%Y %H:%M:%S',
            filemode='w',
            format='%(asctime)s, %(levelname)s, %(message)s'
        )

        current_dir = os.path.dirname(os.path.abspath(__file__))
        bd_dir = current_dir + r'\resources\БАЗА ДАННЫХ'
        input_lists = [
            make_hrefs_list(bd_dir + r'\litefinance hrefs.xlsx'),
            make_hrefs_list(bd_dir + r'\forex4you hrefs.xlsx')
        ]
        litefinance_xpathes = {
            'trader_name': fr'//div[@class = "page_header_part traders_body"]//h2'
        }
        forex4you_xpathes = {
            'trader_name': fr'//span[@data-ng-bind= "::$headerCtrl.leader.displayName"]'
        }

        options = webdriver.ChromeOptions()
        options.add_argument('chromedriver_binary.chromedriver_filename')
        # options.add_argument('headless')
        options.add_argument("window-size=1920,1080")
        driver = webdriver.Chrome(options=options)
        driver.maximize_window()

        for site in input_lists:
            for href in site:
                if href is None:
                    continue
                elif 'forex4you' in href.value:
                    forex4you = Forex4you(
                        current_dir,
                        bd_dir,
                        driver,
                        href,
                        'forex4you',
                        forex4you_xpathes
                    )
                    forex4you.scrap_all()
                elif 'litefinance' in href.value:
                    litefinance = litefinance(
                        current_dir,
                        bd_dir,
                        driver,
                        href,
                        'litefinance',
                        litefinance_xpathes
                    )
                    litefinance.scrap_all()
        driver.quit()
    except Exception as exept:
        logging.error(exept)
        print('Произошла ошибка. Описание ошибки смотрите в логах.')


if __name__ == '__main__':
    logging.basicConfig(
        level=logging.ERROR,
        filename='main.log',
        datefmt='%d.%m.%Y %H:%M:%S',
        filemode='w',
        format='%(asctime)s, %(levelname)s, %(message)s'
    )
    current_dir = os.path.dirname(os.path.abspath(__file__))
    bd_dir = current_dir + r'\resources\БАЗА ДАННЫХ'
    litefinance_list = make_hrefs_list(bd_dir + r'\litefinance hrefs.xlsx'),
    forex4you_list = make_hrefs_list(bd_dir + r'\forex4you hrefs.xlsx')
    input_lists = [make_hrefs_list(bd_dir + r'\litefinance hrefs.xlsx'),
                   make_hrefs_list(bd_dir + r'\forex4you hrefs.xlsx')]
    litefinance_xpathes = {
        'trader_name': fr'//div[@class = "page_header_part traders_body"]//h2'
    }
    forex4you_xpathes = {
        'trader_name': fr'//span[@data-ng-bind= "::$headerCtrl.leader.displayName"]'
    }
    options = webdriver.ChromeOptions()
    options.add_argument('chromedriver_binary.chromedriver_filename')
    # options.add_argument('headless')
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    for site in input_lists:
        for href in site:
            if href is None:
                continue
            elif 'forex4you' in href.value:
                forex4you = Forex4you(
                    current_dir,
                    bd_dir,
                    driver,
                    href,
                    'forex4you',
                    forex4you_xpathes
                )
                forex4you.scrap_all()
            elif 'litefinance' in href.value:
                litefinance = Litefinance(
                    current_dir,
                    bd_dir,
                    driver,
                    href,
                    'litefinance',
                    litefinance_xpathes
                )
                litefinance.scrap_all()
    driver.quit()