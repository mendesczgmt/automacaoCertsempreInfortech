from selenium import webdriver
from selenium.webdriver.common.by import By
from Login import Login

login = Login()
cnpj = str(login.get_cnpj())
senha = str(login.get_senha())


def fazer_login(browser, cnpj, senha):
    entrar_login = browser.find_element(By.XPATH, '//*[@id="Login"]').send_keys(cnpj)
    entrar_senha = browser.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(senha)
    entrar_button = browser.find_element(By.XPATH, '//*[@id="botao-logar"]').click()


