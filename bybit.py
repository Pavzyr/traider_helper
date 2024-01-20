import time
from selenium import webdriver
import openpyxl
import io
import pandas as pd
from datetime import datetime, timedelta
import re
from selenium.webdriver.chrome.service import Service


def remove_special_chars(string):
    pattern = r'[^\w\s]'
    return re.sub(pattern, '', string)


def bybit_scrap(current_dir, href, stop_date, dict_for_traders):
    site = 'bybit'
    df_for_trader = pd.DataFrame(dict_for_traders)
    print(f'Перехожу по ссылке трейдера: {href.value}\n')
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument('chromedriver_binary.chromedriver_filename')
    # options.add_argument('headless')
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    driver.get(href.value)
    print(f'Успешно перешел по ссылке {href.value}\n')
    name = remove_special_chars(driver.find_element("xpath",
                                                    r'(//h1[@class="leader-detail__lf-name"])[1]').text)
    print(f'Имя трейдера = {name}\n')
    driver.find_element("xpath",
                        fr'//div[contains(text(), "Статистика скопированных сделок")]').click()
    time.sleep(5)

    for o in (range(1, 50)):
        time.sleep(2)
        count = 0
        while count == 0:
            count = len(driver.find_elements("xpath",
                                             fr'//div[@class = "order-list-title"]//span[contains(text(), '
                                             fr'"Завершённые сделки")]/../../..//div[@class="order-list-table-box"]'
                                             fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'))
        print(f'Начинаю обработку {count} записей на странице {o}\n')
        for l in list(range(1, count + 1)):
            currency = driver.find_element("xpath",
                                           fr'(//div[@class = "order-list-title"]//span[contains(text(), '
                                           fr'"Завершённые сделки")]/../../..//div[@class="order-list-table-box"]'
                                           fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                           fr'//td[@class = "ant-table-cell ant-table-cell-fix-left '
                                           fr'ant-table-cell-fix-left-last"]'
                                           fr'//span[@class = "c"])[{l}]').text.replace("USDT", "USD")
            if currency is not None:
                date_close = driver.find_element("xpath",
                                                 fr'//div[@class = "order-list-title"]//span[contains(text(), '
                                                 fr'"Завершённые сделки")]/../../..//div[@class="order-list-table-box"]'
                                                 fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                                 fr'[{l}]//td[@class = "ant-table-cell"][6]').text
                date_close = datetime.strptime(date_close, '%Y-%m-%d %H:%M:%S') + timedelta(hours=-1)
                date_open = driver.find_element("xpath",
                                                fr'//div[@class = "order-list-title"]//span[contains(text(), '
                                                fr'"Завершённые сделки")]/../../..'
                                                fr'//div[@class="order-list-table-box"]'
                                                fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                                fr'[{l}]//td[@class = "ant-table-cell"][4]').text
                date_open = datetime.strptime(date_open, '%Y-%m-%d %H:%M:%S') + timedelta(hours=-1)
                type_of_trade = driver.find_element("xpath",
                                                    fr'(//div[@class = "order-list-title"]//span[contains(text(), '
                                                    fr'"Завершённые сделки")]/../../..'
                                                    fr'//div[@class="order-list-table-box"]'
                                                    fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                                    fr'//td[@class = "ant-table-cell '
                                                    fr'ant-table-cell-fix-left ant-table-cell-fix-left-last"]'
                                                    fr'//span[2])[{l}]').text.lower()
                if type_of_trade == 'лонг':
                    type_of_trade = 'buy'
                else:
                    type_of_trade = 'sell'
                price_open = driver.find_element("xpath",
                                                 fr'//div[@class = "order-list-title"]//span[contains(text(), '
                                                 fr'"Завершённые сделки")]/../../..'
                                                 fr'//div[@class="order-list-table-box"]'
                                                 fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                                 fr'[{l}]//td[@class = "ant-table-cell"][3]').text.replace(",", "")
                price_open = price_open.split('.')
                price_open1 = "".join(filter(str.isdigit, price_open[0]))
                price_open2 = "".join(filter(str.isdigit, price_open[1]))
                price_open = f'{price_open1}.{price_open2}'
                price_open = float(price_open)
                price_close = driver.find_element("xpath",
                                                  fr'//div[@class = "order-list-title"]//span[contains(text(), '
                                                  fr'"Завершённые сделки")]/../../..'
                                                  fr'//div[@class="order-list-table-box"]'
                                                  fr'//tbody//tr[@class="ant-table-row ant-table-row-level-0"]'
                                                  fr'[{l}]//td[@class = "ant-table-cell"][5]').text.replace(",", "")
                price_close = price_close.split('.')
                price_close1 = int("".join(filter(str.isdigit, price_close[0])))
                price_close2 = int("".join(filter(str.isdigit, price_close[1])))
                price_close = f'{price_close1}.{price_close2}'
                price_close = float(price_close)

                df_for_trader.loc[len(df_for_trader.index)] = [
                    "0",
                    currency,
                    type_of_trade,
                    date_open.strftime('%Y.%m.%d %H:%M'),
                    price_open,
                    date_close.strftime('%Y.%m.%d %H:%M'),
                    price_close,
                    "0",
                    href.value,
                ]
        if date_close < stop_date:
            break
        driver.execute_script("arguments[0].scrollIntoView();",
                              driver.find_element("xpath",
                                                  fr'//button[@class = "ant-pagination-item-link"]'
                                                  fr'//span[@aria-label = "right"]'))
        time.sleep(2)
        driver.find_element("xpath",
                            fr'//button[@class = "ant-pagination-item-link"]//span[@aria-label = "right"]').click()
        time.sleep(2)
    driver.quit()
    df_for_trader.to_excel(
        fr'{current_dir}\resources\output excel\{name} {site}.xlsx',
        sheet_name='Sheet1',
        index=False)
    wb = openpyxl.load_workbook(
        fr'{current_dir}\resources\output excel\{name} {site}.xlsx')
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
    wb.save(fr'{current_dir}\resources\output excel\{name} {site}.xlsx')
    print(f'✅ Успешно сформировал excel с сигналами трейдера {name}\n')
    text = ''
    for index, row in df_for_trader.iterrows():
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
    with io.open(fr'{current_dir}\resources\template1.htm', 'r', encoding='utf-8') as f:
        html_string = f.read()
    htm = html_string.replace("SSSSS", text)
    with io.open(fr'{current_dir}\resources\output htm\{name} {site}.htm',
                 'w', encoding='utf-8') as file:
        file.write(htm)
    print(f'✅ Успешно сформировал htm с сигналами трейдера {name}\n')
