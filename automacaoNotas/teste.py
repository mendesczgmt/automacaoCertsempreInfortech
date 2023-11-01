import time
import pandas as pd

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from Login import Login
from Formatacao import Formatacao

browser = webdriver.Chrome()
lerPlanilha = pd.read_excel('C:\\Users\\Suporte\\Downloads\\PlanilhaSara.xlsx', sheet_name='16 a 20.10')
lerPlanilha['Protocolo'] = lerPlanilha['Protocolo'].astype(str)
lerPlanilha['Documento'] = lerPlanilha['Documento'].astype(str)
lerPlanilha['Valor do Boleto'] = lerPlanilha['Valor do Boleto'].astype(str)

browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/login.jsf")
loginCpf = browser.find_element(By.ID, "j_username")
password = browser.find_element(By.ID, "j_password")
botaoEntrar = browser.find_element(By.ID, "commandButton_entrar")

login = Login()
formatacao = Formatacao()

cpf = str(login.get_cpf())
senha = str(login.get_senha())

loginCpf.send_keys(str(login.get_cpf()))
password.send_keys(str(login.get_senha()))
botaoEntrar.click()

for x in range(7):
    try:
        botaoContinuar = browser.find_element(By.ID, 'formMensagens:commandButton_confirmar')
        botaoContinuar.click()
    except:
        print("")

browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/paginas/nfse/NFSe_EmitirNFse.jsf")

browser.execute_script("window.open('', '_blank');")
browser.switch_to.window(browser.window_handles[-1])
browser.maximize_window()
browser.get("https://acsafeweb.safewebpss.com.br/gerenciamentoac")

for x in range(int(formatacao.quantidadeNotas())):
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

    campoProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[4]/div[1]/input')

    while True:
        try:
            campoProtocolo.send_keys(lerPlanilha['Protocolo'][x])
            break
        except:
            continue
    
    botaoPesquisarProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[9]/div/div/div/button[1]').click()

    while True:
        try:
            lupaProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[2]/table/tbody/tr/td[2]/i[1]').click()
            break
        except:
            continue

    while True:
        try:
            time.sleep(5)
            cep = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[6]/div[1]/input').get_attribute("value")
            print(str(cep))
            break
        except:
            continue
    
    browser.get("https://acsafeweb.safewebpss.com.br/gerenciamentoac")


    