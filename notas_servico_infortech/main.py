import pandas as pd
from Login import Login
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

browser = webdriver.Chrome()
browser.maximize_window()

login = Login()

cnpj = str(login.get_cnpj())
senha = str(login.get_senha())

