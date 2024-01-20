import time
from selenium import webdriver
import openpyxl
import io
import pandas as pd
from datetime import datetime, timedelta
import re


def remove_special_chars(string):
    pattern = r'[^\w\s]'
    return re.sub(pattern, '', string)


def myfxbook_strategies_scrap(current_dir, href, stop_date, dict_for_traders):
    site = 'myfxbook'
    df_for_trader = pd.DataFrame(dict_for_traders)
    print(f'Перехожу по ссылке трейдера: {href.value}\n')
    options = webdriver.ChromeOptions()
    # options.add_argument('headless')
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get(href.value)
    print(f'Успешно перешел по ссылке {href.value}\n')
    name = remove_special_chars(driver.find_element("xpath",
                                                    r'//span[@itemprop = "name" and @class = "text-overflow-ellipsis '
                                                    r'overflow-hidden white-space-nowrap '
                                                    r'max-width-100 display-inline-block"]').text)
    print(f'Имя трейдера = {name}\n')
    for o in (range(2, 50)):
        time.sleep(2)
        count = 0
        while count == 0:
            count = len(driver.find_elements("xpath", fr'//tbody//tr[@class]'))
        print(f'Начинаю обработку {count} записей на странице {o - 1}\n')
        for l in list(range(1, count)):
            currency = driver.find_element("xpath",
                                           fr'(//tbody//tr[@class])[{l}]//td[3]').text
            if currency is not None:
                date_close = driver.find_element("xpath",
                                                 fr'(//tbody//tr[@class])[{l}]//td[2]').text
                date_close = datetime.strptime(date_close, '%m.%d.%Y %H:%M') - timedelta(hours=1)
                date_open = driver.find_element("xpath",
                                                fr'(//tbody//tr[@class])[{l}]//td[1]').text
                date_open = datetime.strptime(date_open, '%m.%d.%Y %H:%M') - timedelta(hours=1)
                type_of_trade = driver.find_element("xpath",
                                                    fr'(//tbody//tr[@class])[{l}]//td[4]').text.lower()
                obj = driver.find_element("xpath",
                                          fr'(//tbody//tr[@class])[{l}]//td[5]').text.replace(".", ",")
                currency = driver.find_element("xpath",
                                               fr'(//tbody//tr[@class])[{l}]//td[3]').text \
                    .replace("XAUUSD", "GOLD")
                price_open = driver.find_element("xpath",
                                                 fr'(//tbody//tr[@class])[{l}]//td[8]') \
                    .text.replace(",", "")
                price_close = driver.find_element("xpath",
                                                  fr'(//tbody//tr[@class])[{l}]//td[9]') \
                    .text.replace(",", "")
                points = driver.find_element("xpath",
                                             fr'(//tbody//tr[@class])[{l}]//td[11]').text.replace(".", ",")
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
        if date_close < stop_date:
            break
        for i in range(30):
            try:
                driver.execute_script("arguments[0].scrollIntoView();",
                                      driver.find_element("xpath", fr'(//li[@class = "next"]//a)[1]'))
                driver.find_element("xpath",
                                    fr'(//li[@class = "next"]//a)[1]').click()
                break
            except:
                try:
                    element = driver.find_element("xpath",
                                                  fr'//button[contains(text(), "Continue to Myfxbook.com")]')
                    element.click()
                except:
                    pass
                try:
                    element = driver.find_element("xpath",
                                                  fr'//button[@id = "closeModalWebNotification"]')
                    element.click()
                except:
                    pass
            time.sleep(1)

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
