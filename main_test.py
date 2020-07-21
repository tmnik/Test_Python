import xlsxwriter
import time
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from win32com.client import Dispatch


def browser_usd():
    driver_usd = webdriver.Chrome()
    driver_usd.maximize_window()
    driver_usd.wait = WebDriverWait(driver_usd, 5)
    driver_usd.get("https://yandex.ru/")
    time.sleep(5)
    driver_usd.find_element_by_link_text('USD').click()
    driver_usd.switch_to.window(driver_usd.window_handles[1])
    time.sleep(5)
    url_usd = driver_usd.current_url
    driver_usd.quit()
    return url_usd


def browser_eur():
    driver_eur = webdriver.Chrome()
    driver_eur.maximize_window()
    driver_eur.wait = WebDriverWait(driver_eur, 5)
    driver_eur.get("https://yandex.ru/")
    time.sleep(5)
    driver_eur.find_element_by_link_text('EUR').click()
    driver_eur.switch_to.window(driver_eur.window_handles[1])
    time.sleep(5)
    url_eur = driver_eur.current_url
    driver_eur.quit()
    return url_eur


def load_table(url):
    page = requests.get(url).text
    soup = BeautifulSoup(page, 'html5lib')
    table_list = soup.find(class_='quote__data')
    table_list_items = table_list.find_all('tr')
    return table_list_items


def create_table(i_usd, i_euro):
    create_title(i_usd, i_euro)
    [dates_usd, exchanges_usd, changes_usd,
     dates_euro, exchanges_euro, changes_euro] = arrays_dates(i_usd, i_euro)
    col = list(range(7))
    row = 1
    acc_format_usd = workbook.add_format({'num_format': ' [$$-en-US]* # ##0.0000 ; [$$-en-US]* -# ##0.0000\ ; [$$-en-US]* "-"???? ; -@ '})
    acc_format_eur = workbook.add_format({'num_format': ' [$€-x-euro2]* # ##0.0000 ; [$€-x-euro2]* -# ##0.0000\ ; [$€-x-euro2]* "-"???? ; -@ '})
    data_format = workbook.add_format({'num_format': 'DD.MM.YY;@'})
    for i in range(len(i_euro[1:])):
        worksheet.write(row, col[0], dates_usd[i], data_format)
        worksheet.write_number(row, col[1], exchanges_usd[i], acc_format_usd)
        worksheet.write_number(row, col[2], changes_usd[i], acc_format_usd)
        worksheet.write(row, col[3], dates_euro[i], data_format)
        worksheet.write_number(row, col[4], exchanges_euro[i], acc_format_eur)
        worksheet.write_number(row, col[5], changes_euro[i], acc_format_eur)
        row += 1


def create_title(i_usd, i_euro):
    bold = workbook.add_format({'bold': True})
    for j in range(3):
        worksheet.write(0, j, i_usd[0].contents[j].text, bold)
        worksheet.write(0, j + 3, i_euro[0].contents[j].text, bold)


def arrays_dates(i_usd, i_euro):
    dates_usd, exchanges_usd, changes_usd = list(), list(), list()
    dates_euro, exchanges_euro, changes_euro = list(), list(), list()
    for cell in i_usd[1:]:
        dates_usd.append(cell.contents[0].text)
        exchanges_usd.append(float(cell.contents[1].text.replace(',', '.')))
        changes_usd.append(float(cell.contents[2].text.replace(',', '.')))
    for cell in i_euro[1:]:
        dates_euro.append(cell.contents[0].text)
        exchanges_euro.append(float(cell.contents[1].text.replace(',', '.')))
        changes_euro.append(float(cell.contents[2].text.replace(',', '.')))
    return dates_usd, exchanges_usd, changes_usd, dates_euro, exchanges_euro, changes_euro


def calculation_euro_usd():
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 6, 'EUR/USD', bold)
    formula_format = workbook.add_format({'num_format': '0.0000'})
    worksheet.write_array_formula('G2:G11', '{=E2:E11/B2:B11}', formula_format)


def check_for_numbers():
    bold = workbook.add_format({'bold': True})
    worksheet.write('I2', 'Проверка числовых ячеек (подсчет автосуммы):', bold)
    formula_format = workbook.add_format({'num_format': '0.0000'})
    worksheet.write_formula('J2', '=SUM(B2:B11,C2:C11,E2:E11,F2:F11,G2:G11)', formula_format)


def auto_column_width():
    excel = Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(r"C:\Users\medve\OneDrive\Test\test.xlsx")
    excel.Worksheets(1).Activate()
    excel.ActiveSheet.Columns.AutoFit()
    wb.Save()
    wb.Close()


def main():
    url_usd = browser_usd()
    url_eur = browser_eur()
    items_usd = load_table(url_usd)
    items_eur = load_table(url_eur)
    create_table(items_usd, items_eur)
    calculation_euro_usd()
    check_for_numbers()


workbook = xlsxwriter.Workbook('test.xlsx')
worksheet = workbook.add_worksheet()
main()
workbook.close()
auto_column_width()
import sending_report

