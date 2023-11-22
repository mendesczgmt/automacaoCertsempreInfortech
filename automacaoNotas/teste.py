import time
import pandas as pd
import pyautogui
import openpyxl
import os

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import email.message
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from Login import Login
from Formatacao import Formatacao
from selenium import webdriver

browser = webdriver.Chrome()
browser.maximize_window()

############## Leitura da planilha ##############

shetname = '13 a 18.11'
caminhoNotas = 'C:\\Users\\Suporte\\downloads\\Emissões NF Novembro JP.xlsx'

lerPlanilha = pd.read_excel(caminhoNotas, sheet_name=shetname, dtype={'Documento': str})

lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)
lerPlanilha['Produto'] = lerPlanilha['Produto'].astype(str)
lerPlanilha['Nome do AVP'] = lerPlanilha['Nome do AVP'].astype(str)
lerPlanilha['Protocolo'] = lerPlanilha['Protocolo'].astype(str)
lerPlanilha['Documento'] = lerPlanilha['Documento'].astype(str)
lerPlanilha['Valor do Boleto'] = lerPlanilha['Valor do Boleto'].astype(str)
lerPlanilha['E-mail do Titular'] = lerPlanilha['E-mail do Titular'].astype(str)

###################################################

browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/login.jsf")
browser.execute_script("window.open('', '_blank');")
browser.switch_to.window(browser.window_handles[-1])
browser.get("https://acsafeweb.safewebpss.com.br/gerenciamentoac")
browser.switch_to.window(browser.window_handles[0])

loginCpf = browser.find_element(By.ID, "j_username")
password = browser.find_element(By.ID, "j_password")
botaoEntrar = browser.find_element(By.ID, "commandButton_entrar")

login = Login()
formatacao = Formatacao()

cpf = str(login.get_cpf())
senha = str(login.get_senha())
contador = 0


while True:
    try:
        loginCpf.send_keys(str(login.get_cpf()))
        time.sleep(0.5)
        password.send_keys(str(login.get_senha()))
        time.sleep(0.5)
        botaoEntrar.click()
        break
    except:
        continue

# for x in range(7):
#     try:
#         botaoContinuar = browser.find_element(By.ID, 'formMensagens:commandButton_confirmar')
#         botaoContinuar.click()
#     except:
#         print("Não consigo")

browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/paginas/nfse/NFSe_EmitirNFse.jsf")

for x in range(int(formatacao.quantidadeNotas())):
    
    ######### OPÇÕES DO NAVEGADOR #########

    options = webdriver.ChromeOptions()
    nome_cliente = lerPlanilha['Nome'][x]
    caminho_pdf = f'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\Notas fiscais\\11 - NOVEMBRO 2023\\{nome_cliente}.pdf'

    options.add_experimental_option('prefs', {
        'download.default_directory': caminho_pdf,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'plugins.always_open_pdf_externally': True,
        'safebrowsing.enabled': True,
        'profile.default_content_settings.popups': 0
    })

    ######## OPÇÕES DO NAVEGADOR #########

    while True:
        try:
            botaoNao = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:selectradio_retencao_iss"]/tbody/tr/td[3]/div').click()
            break
        except:
            continue

    while True:
        try:
            formulario = browser.find_element(By.ID, 'form_emitir_nfse:calendar_competencia_input').send_keys(str(formatacao.get_data()))
            break
        except:
            continue

    while True:
        try:
            botaoProsseguir = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar').click()
            break
        except:
            continue
    
    time.sleep(1.5)
    caixa_cpf_cnpj = browser.find_element(By.NAME, 'form_emitir_nfse:inputmask_cpf_cnpj')
    caixa_cpf_cnpj.clear()
    time.sleep(1.5)
    caixa_cpf_cnpj.send_keys(lerPlanilha['Documento'][x])
    time.sleep(1.5)
    botao_cpf_cnpj = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:commandbutton_buscar_cpfcnpj"]/span').click()
    time.sleep(1.5)

    while True:
        try:
            browser.find_element(By.XPATH, '//*[@id="j_idt47:msgPrincipal"]/div') ## Spam sem cadastro
            browser.switch_to.window(browser.window_handles[-1])
            
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
            
            campoProtocolo = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-lista/form/div/div/div[1]/div[2]/pages-filter/div/div[4]/div[1]/input')
            
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
            
            while True:
                try:
                    razaoSocial = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[3]/div[1]/input').get_attribute("value")
                    if(razaoSocial != ""):
                        break
                except:
                    continue

            while True:
                try:
                    cep = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[6]/div[1]/input').get_attribute("value")
                    if(cep != ""):
                        break
                except:
                    continue

            while True:
                try:
                    numero = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[5]/div[2]/input').get_attribute("value")
                    if(numero != ""):
                        break
                except:
                    continue

            while True:
                try:
                    logradouro = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[5]/div[1]/input').get_attribute("value")
                    if(logradouro != ""):
                        break
                except:
                    continue

            while True:
                try:
                    bairro = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[5]/div[3]/input').get_attribute("value")
                    if(bairro != ""):
                        break
                except:
                    continue

            while True:
                try:
                    municipio = browser.find_element(By.XPATH, '/html/body/app-root/div/div/app-pages/div[1]/div/div/ng-component/app-relatorio-emissao-gestao/form/div/div/div[1]/div[2]/div[6]/div[3]/input').get_attribute("value")
                    if(municipio != ""):
                        break
                except:
                    continue

            else:
                continue

            browser.switch_to.window(browser.window_handles[0])

            while True:
                try:
                    campoRazaoSocial = browser.find_element(By.ID,'form_emitir_nfse:inputtext_nomeempresarial_nome').send_keys(razaoSocial)
                    break
                except:
                    continue

            while True:
                try:
                    campoCep = browser.find_element(By.ID,'form_emitir_nfse:sispmjp_endereco:inputmask_cep').click()
                    campoCep = browser.find_element(By.ID,'form_emitir_nfse:sispmjp_endereco:inputmask_cep').send_keys(cep)
                    break
                except:
                    continue

            while True:
                time.sleep(1.5)
                if(campoCep != ""):
                    botaoBuscarCep = browser.find_element(By.ID,'form_emitir_nfse:sispmjp_endereco:commandButton_buscar').click()
                    break
                else:
                    continue
            
            while True:
                try:
                    time.sleep(2)
                    botaoSelecionarEstado = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:select_estado"]/div[3]').click()
                    time.sleep(1)
                    botaoSelecionarEstado = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:select_estado_panel"]/div/ul/li[2]').click()
                    break
                except:
                    break

            while True:
                try:
                    time.sleep(1.5)
                    botaoSelecionarCidade = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio"]/div[3]').click()
                    time.sleep(1)
                    municipio = municipio.lower()
                    cidades = ['//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio_panel"]/div/ul/li[2]', '//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio_panel"]/div/ul/li[3]', '//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio_panel"]/div/ul/li[136]', '//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio_panel"]/div/ul/li[212]','//*[@id="form_emitir_nfse:sispmjp_endereco:select_municipio_label"]']
                    for y in cidades:
                        cidade = browser.find_element(By.XPATH, f'{y}').text
                        if(cidade.lower() == municipio):
                            cidadeClique = browser.find_element(By.XPATH, f'{y}').click()
                            break
                    break
                except:
                    print("erro nas cidades")
                    break

            while True:
                try:
                    campoLogradouro = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputtext_lagradouro"]').clear()
                    campoLogradouro = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputtext_lagradouro"]').send_keys(logradouro)
                    break
                except:
                    continue
            while True:
                try:
                    campoNumero = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputmask_numero"]').clear()
                    campoNumero = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputmask_numero"]').send_keys(numero)
                    break
                except:
                    continue
            while True:
                try:
                    campoBairro = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputtext_bairro"]').clear()
                    campoBairro = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:sispmjp_endereco:inputtext_bairro"]').send_keys(bairro)
                    break
                except:
                    continue
        except:
            break                
    while True:
        try:
            botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar').click()
            break
        except:
            print("erro ao clicar")
            continue

    while True:
        try:
            itemListaServico = browser.find_element(By.ID, 'form_emitir_nfse:selectOneMenu_lista_servico_label').click()
            break
        except:
            continue

    while True:
        try:
            servico = browser.find_element(By.ID, 'form_emitir_nfse:selectOneMenu_lista_servico_panel').click()
            time.sleep(2)
            break
        except:
            continue

    while True:
        try:
            descricaoDetalhada = browser.find_element(By.ID, 'form_emitir_nfse:intputtextarea_descricao_detalhada').send_keys("Referente a emissão do Certificado Digital")
            break
        except:
            continue

    while True:
        try:
            botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar_mensagem').click()
            break
        except:
            continue
    while True:
        try:
            botaoContinuar = browser.find_element(By.ID, 'form_emitir_nfse:commandButton_continuar').click()
            break
        except:
            continue

    while True:
        try:
            valor = (lerPlanilha['Valor do Boleto'][x])
            valor = valor + '00'
            valorServico = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:intputmask_valor_servico"]').click()
            valorServico = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:intputmask_valor_servico"]').send_keys(valor)
            #valorServico = browser.find_element(By.ID, 'form_emitir_nfse:intputmask_valor_servico').click()
            #valorServico = browser.find_element(By.ID, 'form_emitir_nfse:intputmask_valor_servico').send_keys(valor)
            break
        except:
            continue

    # //*[@id="j_idt19:msgPrincipal"]/div/ul/li/span[1] span acesso negado

    while True:
        try:
            botaoEmitir = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:commandButton_emitir"]/span[2]').click()
            break
        except:
            continue

    while True:
        try:
            spanConfirmarEmitir = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:commandbutton_confirmdialog_sim"]/span').click()
            break
        except:
            continue 
    
    while True:
        try:
            visualisarNota = browser.find_element(By.XPATH, '//*[@id="form_confirmation_nfse:j_idt56"]/span[2]').click()
            break
        except:
            continue 

    while True:
        try:
            endereco = 'C:\\Users\\Suporte\\OneDrive\\Documentos\\Kaio\\automacaoCertsempreInfortech\\automacaoNotas\\dwn.png'
            img = pyautogui.locateCenterOnScreen(endereco)
            pyautogui.click(img.x, img.y)
            time.sleep(3)
#           img = pyautogui.locateCenterOnScreen('dwn.jpg', confidence=0.9)
#           pyautogui.click(img.x, img.y)
#           time.sleep(3)
            break
        except:
            continue

    while True:
        try:
            pyautogui.write(lerPlanilha['Nome'][x])
            break
        except:
            continue
    
    while True:
        try:
            pyautogui.press('enter')
            break
        except:
            continue
    
    time.sleep(7)
    
    nome_cliente = lerPlanilha['Nome'][x]

    caminho_pdf =f'C:\\Users\\Suporte\\Downloads\\{nome_cliente}.pdf'

    if os.path.exists(caminho_pdf):
        print(f"        DOWNLOAD CONCLUÍDO PARA {nome_cliente}")
    else:
        print(f"        ERRO NO DOWNLOAD PARA {nome_cliente}")
    

    while True:
        try:
            botaoContinuar = browser.find_element(By.XPATH, '//*[@id="form_emitir_nfse:commandbutton_fechar_visualizacao"]/span').click()
            break
        except:
            continue
    


        
    corpo_email = """
    <p>Segue em anexo a nota fiscal referente ao certificado digital</p>
    """

    msg = MIMEMultipart()
    msg['Subject'] = "Nota fiscal Certificado Digital"
    msg['From'] = 'notafiscalcertsempre@gmail.com'
    msg['To'] = lerPlanilha['E-mail do Titular'][x]
    password = 'ksbzriefhjehzjkl'

    msg.attach(MIMEText(corpo_email, 'html'))

    with open(caminho_pdf, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={nome_cliente}.pdf")
        msg.attach(part)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    s.login(msg['From'], password)

    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

    contador +=1

    print('———————————————————————————–—————–—————–—————–——–—————–——–—————–——–—————–——–—————–')
    print(f'⇨       EMAIL ENVIADO PARA O CLIENTE 📩 nº{contador}')
    print('———————————————————————————–—————–—————–—————–——–—————–——–—————–——–—————–——–—————–')



# botao home //*[@id="j_idt20:j_idt22"]/span[2]