import time
import pandas as pd
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

shetname = 'Primeira'
caminhoNotas = 'C:\\Users\\Suporte\\downloads\\Conferir Pagamentos.xlsx'

lerPlanilha = pd.read_excel(caminhoNotas, sheet_name=shetname)

browser = webdriver.Chrome()
browser.maximize_window()
browser.get("https://safe2pay.com.br/login?session=timeout")

email_login = browser.find_element(By.XPATH, '//*[@id="inputLoginMail"]').send_keys('financeirocertsempre@gmail.com')
senha_login = browser.find_element(By.XPATH, '//*[@id="inputLoginPassword"]').send_keys('Certsempre@123@')
entrar = browser.find_element(By.XPATH, '//*[@id="SignInModal"]/div/div/div[3]/button').click()

while True:
    try:
        botao_transacoes = browser.find_element(By.XPATH, '/html/body/div[4]/div/div[3]/div/div/div/div/div/h4').click()
        break
    except:
        continue

while True:
    try:
        calendario = browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[1]/form/div/div[11]/div/p/span/button').click()
        break
    except:
        continue

while True:
    try:
        limpa_data = browser.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/div[1]/form/div/div[11]/div/p/div/ul/li[2]/div/div[1]/button').click()
        break
    except:
        continue





time.sleep(300)