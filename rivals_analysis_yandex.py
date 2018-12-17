from xlutils.copy import copy as xlcopy
import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

url = 'https://taxi.yandex.ru/#index'

driver = webdriver.Firefox()

driver.get(url)

addr_from = driver.find_element_by_id('addressFrom')
addr_to = driver.find_element_by_id('addressTo')

sleep(5)

streets = []
cost_lst = []

excel_file = xlrd.open_workbook('rivals.xls')
sheet = excel_file.sheet_by_index(0)

row_number = sheet.nrows

if row_number > 0:
    for row in range(0, row_number):
        streets.append(str(sheet.row(row)[0]).replace('text:', '').replace("'", ''))
        streets.append(str(sheet.row(row)[1]).replace('text:', '').replace("'", ''))


def add_address_from():
    if len(streets) > 0:
        street = streets[0]
        addr_from.clear()
        addr_from.send_keys(street)
        sleep(2)
        addr_from.send_keys(Keys.ARROW_DOWN)
        sleep(2)
        addr_from.send_keys(Keys.ENTER)
        del streets[0]
        sleep(1)


def add_address_to():
    if len(streets) > 0:
        street = streets[0]
        addr_to.clear()
        addr_to.send_keys(street)
        sleep(2)
        addr_to.send_keys(Keys.ARROW_DOWN)
        sleep(2)
        addr_to.send_keys(Keys.ENTER)
        sleep(3)
        cost = driver.find_element_by_xpath("//div[@class='routestats__price']/span[contains(@class,'text')]").text
        cost_lst.append(str(cost))
        print(cost)
        del streets[0]
        sleep(1)


def write():
    write_book = xlcopy(excel_file)
    w_sheet = write_book.get_sheet(0)

    for i, price in enumerate(cost_lst):
        w_sheet.write(i, 4, price)

    write_book.save('result_yandex.xls')
    print('Парсинг завершен')
    print('Создан файл result_yandex.xls')
    driver.close()


for i in range(len(streets)):
    add_address_from()
    add_address_to()

write()
print(cost_lst)
