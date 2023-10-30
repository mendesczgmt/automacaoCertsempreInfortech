import time
import pandas as pd

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from Login import Login
from Formatacao import Formatacao

planilhaTeste = pd.read_excel('C:\\Users\\Suporte\\Downloads\\TestePlanilha.xlsx', sheet_name='1')

planilhaTeste['Entrada'] = planilhaTeste['Entrada'].astype(str)

planilhaTeste = planilhaTeste.drop(0)

print(planilhaTeste)


