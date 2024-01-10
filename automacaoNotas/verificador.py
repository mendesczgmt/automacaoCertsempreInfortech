import pandas as pd
import os
from Formatacao import Formatacao
formatacao = Formatacao()

def verificadorNotaBaixada():
    shetname = '26.12 a 06.01'
    caminhoNotas = 'C:\\Users\\Suporte\\downloads\\planilha janeiro sara.xlsx'
    lerPlanilha = pd.read_excel(caminhoNotas, sheet_name=shetname, dtype={'Documento': str})
    lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)
    lerPlanilha['Produto'] = lerPlanilha['Produto'].astype(str)
    #lerPlanilha['Nome do AVP'] = lerPlanilha['Nome do AVP'].astype(str)
    lerPlanilha['Protocolo'] = lerPlanilha['Protocolo'].astype(str)
    #lerPlanilha['Documento'] = lerPlanilha['Documento'].astype(str)
    lerPlanilha['Valor do Boleto'] = lerPlanilha['Valor do Boleto'].astype(str)
    #lerPlanilha['E-mail do Titular'] = lerPlanilha['E-mail do Titular'].astype(str)
    # Verifica se já tem nota baixada do mesmo cliente

    cont = 0
    cont_baixada = 0

    for x in range(int(formatacao.quantidadeNotas())):
        
        cont += 1
        nome_cliente = lerPlanilha['Nome'][x]
        caminho_pdf =f'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\Notas fiscais\\ANO 2024\\01 - JANEIRO 2024\\{nome_cliente}.pdf'

        if os.path.exists(caminho_pdf):
            cont_baixada += 1
            print(f"       {cont} - ✅ JÁ EXISTE DOWLOAD PARA ESSE CLIENTE {nome_cliente}")
        else:
            print(f"       {cont} - ❌ NÃO EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
            
    if cont_baixada > 0:
        print('\n')
        alerta = int(input(f"Alerta! Você possui {cont_baixada} notas já baixadas!\nAperte: \n[1] - PARAR [2] - CONTINUAR "))
        if alerta == 1:
            quit()
        else:
            print('Continuando...')
        