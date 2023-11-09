import csv
from selenium import webdriver
from selenium.webdriver import Chrome,ChromeService
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
from selenium.webdriver.common.action_chains import ActionChains


def parsing():
    '''Функция парсит данные с сайта используя файл csv и сохряняет их в список'''
    browser = webdriver.Chrome(service= ChromeService(ChromeDriverManager().install()))
    try:
        my_result = []
        browser.get('https://metal-calculator.ru/page/app')
        name = browser.find_element(by=By.XPATH, value='//*[@id="app"]/div/div[2]/div/div[1]/div/div[1]/div[2]/nav/ul/li[3]/a')
        name.click()
        forma = browser.find_element(by=By.XPATH, value= '//*[@id="app"]/div/div[2]/div/div[1]/div/div[2]/div[2]/nav/ul/li[4]/a')
        forma.click()
        calculator = browser.find_element(by=By.XPATH, value= '//*[@id="app"]/div/div[2]/div/div[2]/div/div/div/div[2]/div[2]/div/form/div')
        with open('tz (1).csv', 'r', encoding='cp1251') as file:
            reader = csv.reader(file, delimiter=';')
            for  row in reader:
                marka = row[4]
                tolschina = row[5]
                width = row[6]
                heught = row[7]
                result = row[8]
                selects = calculator.find_element(by = By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(1) > label > select')
                selects.click()
                select = Select(selects)
                tolschina_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(2) > label > input')
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
                       
            with open('my.xlsx','a') as file:
                for row in my_result:
                    file.writelines(row)
           
            

                
                
           
     
            
    except Exception:
        print('error')

parsing()

