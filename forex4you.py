import time
from selenium import webdriver
import openpyxl
import io
import pandas as pd
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import re


def remove_special_chars(string):
    pattern = r'[^\w\s]'
    return re.sub(pattern, '', string)


def forex4you_scrap(current_dir, href, months_in_numbers, stop_date, dict_for_traders):
    site = 'forex4you'
    df_for_trader = pd.DataFrame(dict_for_traders)
    print(f'Перехожу по ссылке трейдера: {href.value}\n')
    options = webdriver.ChromeOptions()
    # options.add_argument('headless')
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    driver.get(href.value)
    print(f'Успешно перешел по ссылке {href.value}\n')
    WebDriverWait(driver, 20).until(
        ec.presence_of_element_located(
            ("xpath", fr'//span[@data-ng-bind= "::$headerCtrl.leader.displayName"]'))
    )
    name = remove_special_chars(driver.find_element("xpath",
                                                    fr'//span[@data-ng-bind= "::$headerCtrl.leader.displayName"]').text)
    print(f'Имя трейдера = {name}\n')
    try:
        element = WebDriverWait(driver, 1).until(
            ec.presence_of_element_located(("xpath", r"//div[contains(text(), 'Allow All')]")))
        element.click()
    except:
        pass
    driver.find_element("xpath", fr'//label[contains(text(), "Весь период")]').click()
    for o in (range(2, 50)):
        time.sleep(2)
        count = 0
        while count == 0:
            count = len(driver.find_elements("xpath",
                                             fr'//tbody//tr[@data-ng-repeat = '
                                             fr'"trade in $fxGrid.$data track by trade.id"]'
                                             fr'//td[@data-ng-bind="::trade.symbol"]'))
        print(f'Начинаю обработку {count - 10} записей на странице {o-1}\n')
        for l in list(range(1, count - 9)):
            currency = driver.find_element("xpath",
                                           fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]').text
            if currency is not None:
                date_close = driver.find_element("xpath",
                                                 fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                                 '//../preceding-sibling::td[1]').text
                for i in months_in_numbers:
                    date_close = date_close.replace(i, months_in_numbers[i])
                date_close = datetime.strptime(date_close, '%d %m %Y г., %H:%M:%S')
                date_open = driver.find_element("xpath",
                                                fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                                fr'//../preceding-sibling::td[2]').text
                for i in months_in_numbers:
                    date_open = date_open.replace(i, months_in_numbers[i])
                date_open = datetime.strptime(date_open, '%d %m %Y г., %H:%M:%S')
                type_of_trade = driver.find_element("xpath",
                                                    fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                                    fr'//../following-sibling::td[1]').text.lower()
                obj = driver.find_element("xpath",
                                          fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                          fr'//../preceding-sibling::td[3]').text.replace(".", ",")
                currency = driver.find_element("xpath",
                                               fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]').text \
                    .replace("XAUUSD", "GOLD")
                price_open = driver.find_element("xpath",
                                                 fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                                 fr'//../following-sibling::td[2]').text.replace(" ",
                                                                                                 "")
                price_close = driver.find_element("xpath",
                                                  fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                                  fr'//../following-sibling::td[3]').text.replace(" ",
                                                                                                  "")
                points = driver.find_element("xpath",
                                             fr'(//td[@data-ng-bind="::trade.symbol"])[{l}]'
                                             fr'//../following-sibling::td[4]').text.replace(".", ",")
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

        try:
            driver.find_element("xpath",
                                r"//div[contains(text(), 'Allow All')]").click()
        except:
            pass
        driver.find_element("xpath",
                            fr'(//a[@data-fx-grid-set-page="$fxGridPaginator.getNextPage()"])[1]').click()
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
                '<tr align = right>'\
                fr'<td>{row["Объем"]}</td>'\
                fr'<td nowrap>{row["Время Открытия"]}</td>'\
                fr'<td>{row["Тип сделки"]}</td>'\
                fr'<td class=mspt>0</td>'\
                fr'<td>{row["Валютная пара"]}</td>'\
                fr'<td style="mso-number-format:0\.00000;">{row["Цена Открытия"]}</td>'\
                fr'<td style="mso-number-format:0\.00000;">{row["Объем"]}</td>'\
                fr'<td style="mso-number-format:0\.00000;">0</td>'\
                fr'<td class=msdate nowrap>{row["Время Закрытия"]}</td>'\
                fr'<td style="mso-number-format:0\.00000;">{row["Цена Закрытия"]}</td>'\
                fr'<td class=mspt>0</td>'\
                fr'<td class=mspt>0</td>'\
                fr'<td class=mspt>0</td>'\
                fr'<td class=mspt>0</td>'\
                '</tr>'
    with io.open(fr'{current_dir}\resources\template1.htm', 'r', encoding='utf-8') as f:
        html_string = f.read()
    htm = html_string.replace("SSSSS", text)
    with io.open(fr'{current_dir}\resources\output htm\{name} {site}.htm',
                 'w', encoding='utf-8') as file:
        file.write(htm)
    print(f'✅ Успешно сформировал htm с сигналами трейдера {name}\n')