import csv                                            #Много много импортоффф
import pandas as pd
from selenium import webdriver
from selenium.webdriver import Chrome,ChromeService
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
import  openpyxl
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook,cell,writer , worksheet
from selenium.webdriver.common.action_chains import ActionChains


def parsing():
    '''Функция парсит данные с сайта используя файл csv и сохряняет их в список'''
    browser = webdriver.Chrome(service= ChromeService(ChromeDriverManager().install())) #Устанавливаем браузер
    try:
        my_result = []
        browser.get('https://metal-calculator.ru/page/app') #Переходим на сайт
        name = browser.find_element(by=By.XPATH, value='//*[@id="app"]/div/div[2]/div/div[1]/div/div[1]/div[2]/nav/ul/li[3]/a') #Находим название аллюминий
        name.click()
        forma = browser.find_element(by=By.XPATH, value= '//*[@id="app"]/div/div[2]/div/div[1]/div/div[2]/div[2]/nav/ul/li[4]/a') #Находим тип лист\плита
        forma.click()
        calculator = browser.find_element(by=By.XPATH, value= '//*[@id="app"]/div/div[2]/div/div[2]/div/div/div/div[2]/div[2]/div/form/div') #Находиим сам калькулятор
        with open('tz (1).csv', 'r', encoding='cp1251') as file: # Открываем файл csv
            reader = csv.reader(file, delimiter=';')
            for  row in reader: #Находим столбцы (индексы) с которыми будем работать дальше
                marka = row[4]       #Марка
                tolschina = row[5]   #Толщина
                width = row[6]        #Ширина
                heught = row[7]       #Длина
                result = row[8]       # Вес - пока он у нас пустой т.е ''
                # Находим выпадающий список
                selects = calculator.find_element(by = By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(1) > label > select')
                # Жмякаем на него чтобы выпал
                selects.click()
                # Видим все что есть в выпадающем списке
                select = Select(selects)
                # Находим инпут толщина
                tolschina_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(2) > label > input')
                # На всякий случай очищаем потому что фиг знает что там может быть
                tolschina_input.clear()
                tolschina_input.send_keys(row[5])
                width_input = calculator.find_element(by=By.CSS_SELECTOR,value = '#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(3) > label > input')
                width_input.clear()
                width_input.send_keys(row[6])
                hight_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(4) > label > input')
                hight_input.clear()
                hight_input.send_keys(row[7])
                amount_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(5) > label > input')
                amount_input.clear()
                amount_input.send_keys('1')
                if marka in selects.text:
                    select.select_by_visible_text(marka)
                    button = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > a')
                    button.click()
                    result = browser.find_element(by= By.CSS_SELECTOR, value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div:nth-child(3) > div > div:nth-child(2) > div.unit-60 > ul > li.block-price-weight > div > span.result-item-value')
                    if row[8] == '':
                        row[8] = result.text
                        my_result.append(row)
                else:
                    select.select_by_visible_text('Прочее')
                    button = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > a')
                    button.click()
                    result = browser.find_element(by= By.CSS_SELECTOR, value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div:nth-child(3) > div > div:nth-child(2) > div.unit-60 > ul > li.block-price-weight > div > span.result-item-value')
                    if row[8] == '':
                        row[8] = result.text
                        my_result.append(row)

        df = pd.DataFrame(my_result,columns= ['Наименование','Код  артикул','Металл','Сорамент','Марка','Толщина','Ширина','Длина','Вес'])
        df.to_excel('new_1.xlsx',index = False)
                                  
    except TypeError:
        print('error')

#parsing()

