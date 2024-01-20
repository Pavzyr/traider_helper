import time
from selenium import webdriver
import openpyxl
import io
import pandas as pd
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import re


def remove_special_chars(string):
    pattern = r'[^\w\s]'
    return re.sub(pattern, '', string)


def lifefinance_scrap(current_dir, href, stop_date, dict_for_traders):
    site = 'litefinance'
    df_for_trader = pd.DataFrame(dict_for_traders)
    print(f'Перехожу по ссылке трейдера: {href.value}\n')
    options = webdriver.ChromeOptions()
    options.add_argument('chromedriver_binary.chromedriver_filename')
    # options.add_argument('headless')
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get(href.value)
    print(f'Успешно перешел по ссылке {href.value}\n')
    WebDriverWait(driver, 20).until(
        ec.presence_of_element_located(
            ("xpath", fr'//div[@class = "page_header_part traders_body"]//h2'))
    )
    name = remove_special_chars(driver.find_element("xpath", fr'//div[@class = "page_header_part traders_body"]//h2').text)
    print(f'Имя трейдера = {name}\n')
    count_on_page = 1
    for o in range(1, 40):
        time.sleep(2)
        count = len(driver.find_elements("xpath", fr'//div[@class = "content_row"]'))
        print(f'Начинаю обработку c {count_on_page} по {count} запись на странице {o}\n')
        for l in range(count_on_page, count + 1):
            currency = driver.find_element("xpath",
                                           fr'(//div[@class = "content_row"])[{l}]/descendant::a[2]').text
            if currency is not None:
                date_close = driver.find_element("xpath",
                                                 fr'(//div[@class = "content_row"])[{l}]'
                                                 fr'/descendant::div[@class = "content_col"][3]').text
                date_close = datetime.strptime(date_close, '%d.%m.%Y %H:%M:%S') + timedelta(hours=-1)
                date_open = driver.find_element("xpath",
                                                fr'(//div[@class = "content_row"])[{l}]'
                                                fr'/descendant::div[@class = "content_col"][2]').text
                date_open = datetime.strptime(date_open, '%d.%m.%Y %H:%M:%S') + timedelta(hours=-1)
                type_of_trade = driver.find_element("xpath",
                                                    fr'(//div[@class = "content_row"])[{l}]'
                                                    fr'/descendant::div[@class = "content_col"][4]'
                                                    ).text.lower()
                if type_of_trade == 'покупка':
                    type_of_trade = 'buy'
                else:
                    type_of_trade = 'sell'
                obj = driver.find_element("xpath",
                                          fr'(//div[@class = "content_row"])[{l}]'
                                          fr'/descendant::div[@class = "content_col"][5]'
                                          ).text.replace(".", ",")
                currency = driver.find_element("xpath",
                                               fr'(//div[@class = "content_row"])[{l}]/descendant::a[2]') \
                    .text.replace("XAUUSD", "GOLD")
                price_open = driver.find_element("xpath",
                                                 fr'(//div[@class = "content_row"])[{l}]'
                                                 fr'/descendant::div[@class = "content_col"][6]'
                                                 ).text.replace(" ", "")
                price_close = driver.find_element("xpath",
                                                  fr'(//div[@class = "content_row"])[{l}]'
                                                  fr'/descendant::div[@class = "content_col"][7]'
                                                  ).text.replace(" ", "")
                points = driver.find_element("xpath",
                                             fr'(//div[@class = "content_row"])[{l}]'
                                             fr'/descendant::div[@class = "content_col"][8]'
                                             ).text.replace(".", ",")
                df_for_trader.loc[len(df_for_trader.index)] = [
                    obj,
                    currency,
                    type_of_trade,
                    date_open.strftime('%Y.%m.%d %H:%M'),
                    price_open,
                    date_close.strftime('%Y.%m.%d %H:%M'),
                    price_close,
                    points,
                    href.value
                ]
            count_on_page += 1
        # Если количечество сделок меньше 50 на странице - остановить обработку
        if count % 50 != 0:
            break
        if date_close < stop_date:
            break
        driver.execute_script("arguments[0].scrollIntoView();",
                              driver.find_element("xpath", fr'(//div[@class = "content_row"])[{count}]'))
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