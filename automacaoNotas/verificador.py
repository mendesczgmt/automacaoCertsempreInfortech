import time
import pandas as pd
import openpyxl
import os

from Formatacao import Formatacao

formatacao = Formatacao()

shetname = '13 a 18.11'
caminhoNotas = 'C:\\Users\\Suporte\\downloads\\Emissões NF Novembro JP.xlsx'

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

for x in range(int(formatacao.quantidadeNotas())):
    
    cont += 1
    nome_cliente = lerPlanilha['Nome'][x]
    caminho_pdf =f'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\Notas fiscais\\11 - NOVEMBRO 2023\\{nome_cliente}.pdf'

    if os.path.exists(caminho_pdf):
        print(f"       {cont} - ✅ JÁ EXISTE DOWLOAD PARA ESSE CLIENTE {nome_cliente}")
    else:
        print(f"       {cont} - ❌ NÃO EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
