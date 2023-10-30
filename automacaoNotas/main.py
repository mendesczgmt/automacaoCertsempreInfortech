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


browser.execute_script("window.open('', '_blank');")
browser.switch_to.window(browser.window_handles[-1])
browser.maximize_window()
browser.get("https://acsafeweb.safewebpss.com.br/gerenciamentoac")
time.sleep(30)

relatorio = browser.find_element(By.XPATH, '//*[@id="link6"]').click()
time.sleep(1.5)

relatorioEmissao = browser.find_element(By.XPATH, '//*[@id="link9"]').click()

campoProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[4]/div[1]/input')

time.sleep(1.5)
campoProtocolo.click()

campoProtocolo.send_keys(lerPlanilha['Protocolo'][0])

botaoPesquisarProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[9]/div/div/div/button[1]').click()
time.sleep(4)

lupaProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[2]/table/tbody/tr/td[2]/i[1]').click()

cep = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[6]/div[1]/input').text

print(str(cep))

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
        
browser.execute_script("window.open('', '_blank');")
browser.switch_to.window(browser.window_handles[-1])
browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/paginas/nfse/NFSe_EmitirNFse.jsf")
    
botaoNao = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:selectradio_retencao_iss"]/tbody/tr/td[3]/div')
botaoNao.click()
formulario = browser.find_element(By.ID, 'form_emitir_nfse:calendar_competencia_input')
formulario.send_keys(str(formatacao.get_data()))
botaoProsseguir = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar')
botaoProsseguir.click()

time.sleep(2)

caixa_cpf_cnpj = browser.find_element(By.NAME, 'form_emitir_nfse:inputmask_cpf_cnpj')
caixa_cpf_cnpj.send_keys(lerPlanilha['Documento'][0])
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

valorServico.send_keys(lerPlanilha['Valor do Boleto'][0])

lerPlanilha = lerPlanilha.drop[1]


