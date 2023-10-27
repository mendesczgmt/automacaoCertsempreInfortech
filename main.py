import time
import pandas as pd

from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from Login import Login

browser = webdriver.Chrome()

browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/login.jsf")
browser.maximize_window()
loginCpf = browser.find_element(By.ID, "j_username")
password = browser.find_element(By.ID, "j_password")
botaoEntrar = browser.find_element(By.ID, "commandButton_entrar")

login = Login()

cpf_string = str(login.get_cpf())
senha_string = str(login.get_senha())

loginCpf.send_keys(str(login.get_cpf()))
password.send_keys(str(login.get_senha()))
botaoEntrar.click()
for x in range(7):
    try:
        botaoContinuar = browser.find_element(By.ID, 'formMensagens:commandButton_confirmar')
        botaoContinuar.click()
    except:
        print("")
        
browser.execute_script("window.open('', '_blank');")
browser.switch_to.window(browser.window_handles[-1])
browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/paginas/nfse/NFSe_EmitirNFse.jsf")
    
botaoNao = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:selectradio_retencao_iss"]/tbody/tr/td[3]/div')
botaoNao.click()
formulario = browser.find_element(By.ID, 'form_emitir_nfse:calendar_competencia_input')
formulario.send_keys("27/10/2023")
botaoProsseguir = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar')
botaoProsseguir.click()

time.sleep(2)
caixa_cpf_cnpj = browser.find_element(By.NAME, 'form_emitir_nfse:inputmask_cpf_cnpj').send_keys("12.730.842/0001-88")
time.sleep(2)

botao_cpf_cnpj = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:commandbutton_buscar_cpfcnpj"]/span').click()
time.sleep(2)

botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar').click()
time.sleep(2)

itemListaServico = browser.find_element(By.ID, 'form_emitir_nfse:selectOneMenu_lista_servico_label').click()

time.sleep(2)

servico = browser.find_element(By.ID, 'form_emitir_nfse:selectOneMenu_lista_servico_panel').click()

time.sleep(2)

descricaoDetalhada = browser.find_element(By.ID, 'form_emitir_nfse:intputtextarea_descricao_detalhada')
descricaoDetalhada.click()
descricaoDetalhada.send_keys("Referente a emissão do Certificado Digital")

time.sleep(2)

botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar_mensagem').click()
time.sleep(2)

botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar').click()
time.sleep(2)

valorServico = browser.find_element(By.ID, 'form_emitir_nfse:intputmask_valor_servico')
valorServico.click()
valorServico.send_keys("200,00")

time.sleep(60)


time.sleep(60)




