from xlutils.copy import copy as xlcopy
import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep

url = 'https://vladivostok.rutaxi.ru/index.html'

driver = webdriver.Firefox()

driver.get(url)


addr_from = driver.find_element_by_css_selector('#routers div:nth-child(1) input[placeholder="Откуда?"]')
addr_to = driver.find_element_by_css_selector('#routers div:nth-child(2) input[placeholder="Куда?"]')
addr_from_house = driver.find_element_by_css_selector('#routers div:nth-child(1) input[placeholder="№ дома"]')
addr_to_house = driver.find_element_by_css_selector('#routers div:nth-child(2) input[placeholder="№ дома"]')


streets = []
cost_lst = []
houses = []

excel_file = xlrd.open_workbook('rivals_1.xls')
sheet = excel_file.sheet_by_index(0)

row_number = sheet.nrows

if row_number > 0:
    for row in range(0, row_number):
        streets.append(str(sheet.row(row)[0]).replace('text:', '').replace("'", ''))
        streets.append(str(sheet.row(row)[2]).replace('text:', '').replace("'", ''))
        houses.append(str(sheet.row(row)[1]).replace('text:', '').replace("'", '').replace('number:', '').replace('.0', ''))
        houses.append(str(sheet.row(row)[3]).replace('text:', '').replace("'", '').replace('number:', '').replace('.0', ''))


def add_address_from():
    if len(streets) > 0 and len(houses) > 0:
        street = streets[0]
        house = houses[0]
        addr_from.clear()
        addr_from.send_keys(street)
        sleep(2)
        addr_from.send_keys(Keys.ARROW_DOWN)
        sleep(2)
        addr_from.send_keys(Keys.ENTER)
        sleep(1)
        addr_from_house.clear()
        addr_from_house.send_keys(house)
        del streets[0]
        del houses[0]
        sleep(1)


def add_address_to():
    if len(streets) > 0 and len(houses) > 0:
        street = streets[0]
        house = houses[0]
        addr_to.clear()
        addr_to.send_keys(street)
        sleep(2)
        addr_to.send_keys(Keys.ARROW_DOWN)
        sleep(2)
        addr_to.send_keys(Keys.ENTER)
        sleep(1)
        addr_to_house.clear()
        addr_to_house.send_keys(house)
        sleep(3)
        cost = driver.find_element_by_css_selector('#cost span.new_price').text
        cost_lst.append(str(cost))
        print(cost)
        del streets[0]
        del houses[0]
        sleep(1)


def write():
    write_book = xlcopy(excel_file)
    w_sheet = write_book.get_sheet(0)

    for i, price in enumerate(cost_lst):
        w_sheet.write(i, 4, price)

    write_book.save('result_rutaxi.xls')
    print('Парсинг завершен')
    print('Создан файл result_rutaxi.xls')
    driver.close()


for i in range(len(streets)):
    add_address_from()
    add_address_to()

write()
print(cost_lst)

