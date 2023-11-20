import pandas as pd
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from Formatacao import Formatacao

formatacao = Formatacao()

def abrirNavegador():
    browser = webdriver.Chrome()
    browser.maximize_window()

def abrirPlanilha():
    caminhoNotas='C:\\Users\\Suporte\\Downloads\\Antecipações teste.xlsx'
    sheetname='inicial'
    lerPlanilha=pd.read_excel(caminhoNotas, sheet_name=sheetname)
    return lerPlanilha

def abrirGerenciamento(browser):
    browser.get("https://acsafeweb.safewebpss.com.br/gerenciamentoac")
    browser.switch_to.window(browser.window_handles[-1])

def logoGerenciamento(browser):
        while True:
            try:
                logoGerenciamento = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/app-sidebar/div/a').click()
                break
            except:
                continue
        
        return logoGerenciamento



def pesquisaProtocolo(browser):
    logoGerenciamento




abrirNavegador()
abrirPlanilha()
pesquisaProtocolo()


'''
def obrirPlanilha(caminhoNotas='C:\\Users\\Suporte\\Downloads\\Antecipações teste.xlsx', sheetname='inicial'):
    lerPlanilha=pd.read_excel(caminhoNotas, sheet_name=sheetname)
    print(lerPlanilha)


def pesquisaProtocolo(lerPlanilha, browser):
    for x in range(int(formatacao.quantidadeNotas())):
        while True:
            try:
                logoGerenciamento = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/app-sidebar/div/a').click()
                break
            except:
                continue

        while True:
            try:
                relatorio = browser.find_element(By.XPATH, '//*[@id="link6"]').click()
                break
            except:
                continue

        while True:
            try:
                relatorioEmissao = browser.find_element(By.XPATH, '//*[@id="link9"]').click()  
                break
            except:
                continue
        
        while True:
            try:
                campoProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[4]/div[1]/input')
                break
            except:
                continue

        while True:
            try:
                campoProtocolo.send_keys(lerPlanilha['Protocolo'][x])
                break
            except:
                continue

        while True:
            try:
                botaoPesquisarProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[9]/div/div/div/button[1]').click()
                break
            except:
                continue

        while True:
            try:
                lupaProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[2]/table/tbody/tr/td[2]/i[1]').click()
                break
            except:
                continue

abrirPlanilha()
abrirGerenciamento(browser)
pesquisaProtocolo(abrirPlanilha, browser)'''