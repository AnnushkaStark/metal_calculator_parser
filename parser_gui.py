from typing import Self, TextIO                    # Импортируем дохрена всего и стразу
from PyQt6 import QtCore
from PyQt6.QtWidgets import *                      # Потому что QT  капризная хрень и просто на всяукий случай а файл и так весит тонну
from PyQt6.QtWidgets import QWidget
import sys
from PyQt6 import *
import sqlite3 as sql
from parser_1 import *
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
from selenium.webdriver.common.action_chains import ActionChains
from parser_function import *
from openpyxl import Workbook

class Parser(Ui_MainWindow,QMainWindow):                  # Инициализируем окошшшко
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()    
        self.pushButton.clicked.connect(self.start_parsing)  # И одну кнопку
        
     

    def start_parsing(self):
        parsing()                  # вызываем функцию парсинг из файла parser_function.py
        result = QMessageBox()     # Добавляем выпадающее окошко 
        result.setText('The end')  # C надписью конец чтобы когда закончит парсить и запишет нам exel
        result.exec()              # Оно выпало 
                  
        

app = QApplication(sys.argv)
window = Parser()
sys.exit(app.exec())