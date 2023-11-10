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
        my_result = [] # Это временный список в который мы добавим все ряды из csv  после того как добавим в них вес
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
                # Вводим туда значение из столбца толщина в таблице 
                tolschina_input.send_keys(row[5])
                # Находим инпут ширина
                width_input = calculator.find_element(by=By.CSS_SELECTOR,value = '#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(3) > label > input')
                # На всякий случай очищаем
                width_input.clear()
                # передаем туда значение из столбца толщина
                width_input.send_keys(row[6])
                # Ищем инпут длина
                hight_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(4) > label > input')
                # Тоже очищаем
                hight_input.clear()
                # вводим туда значение из столбца толщина
                hight_input.send_keys(row[7])
                # Находим инпут колиество
                amount_input = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > div:nth-child(5) > label > input')
                # Опять очищаем
                amount_input.clear()
                # Передаем туда количество 1 штука т.е '1'
                amount_input.send_keys('1')
                # Ищем марку в выпадающем списке
                if marka in selects.text:
                    # если нашли выбирем ту марку которую нашли
                    select.select_by_visible_text(marka)
                    # Находим кнопку рассчитать
                    button = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > a')
                    # Нажимаем
                    button.click()
                    # Находим место в которое вываливается рассчитанный результат
                    result = browser.find_element(by= By.CSS_SELECTOR, value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div:nth-child(3) > div > div:nth-child(2) > div.unit-60 > ul > li.block-price-weight > div > span.result-item-value')
                    # Если вес не указан, т.е вес = ''
                    if row[8] == '':
                        # Заменяем пустую сртроку на посчитанный результат
                        row[8] = result.text
                        # Закидываем собранный ряд во временный список
                        my_result.append(row)
                else:
                    # А если марка не в выпадающем списке выбираем прочее
                    select.select_by_visible_text('Прочее')
                    # Тоже находим кнопку рассчитать
                    button = calculator.find_element(by=By.CSS_SELECTOR,value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div.units-row.units-split.details-container > div.unit-40 > div > form > div > a')
                    # Тоже на нее жмакаем
                    button.click()
                    # Находим куда вывалился получившийся результат
                    result = browser.find_element(by= By.CSS_SELECTOR, value='#app > div > div.app-content-wrap > div > div.unit-70.app-content.section.section-details.non-border > div > div > div > div:nth-child(3) > div > div:nth-child(2) > div.unit-60 > ul > li.block-price-weight > div > span.result-item-value')
                    # Если вес не указан
                    if row[8] == '':
                        # Заменяем пустую строку на этот результат
                        row[8] = result.text
                        # Тоже закидывааем собранный ряд во временный список
                        my_result.append(row)
        
        #Cоздаем датафрейм из временного списка который был ы начале пустым а теперь туда запихан весь csv  с добавленным где это было нужно весом
        #При создании датафрейма указываем заголовки столбцов которые будем запихивать в exel
        df = pd.DataFrame(my_result,columns= ['Наименование','Код  артикул','Металл','Сорамент','Марка','Толщина','Ширина','Длина','Вес'])
        #Записываем весь наш список в exel
        df.to_excel('new_1.xlsx',index = False)
                                  
    except Exception: #Обрабатываем все исключения сразу ( я заню что это говнокод но мне реально лень except построчно прописывать)
        print('error')

#parsing()  #Так как функция запусается из GUI тут ее вызов закомментирван ( можно раскомментить и запустить а коментарий к комментарию это наркомания)

