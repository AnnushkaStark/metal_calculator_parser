from typing import Self, TextIO
from PyQt6 import QtCore
from PyQt6.QtWidgets import *
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

class Parser(Ui_MainWindow,QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()
        self.pushButton_start.clicked.connect(self.start_parsing)
        self.textEdit.insertPlainText('')
     

    def start_parsing(self):
        return  parsing()
            
        

app = QApplication(sys.argv)
window = Parser()
sys.exit(app.exec())