import pandas as pd
import os
from Formatacao import Formatacao
formatacao = Formatacao()

def verificadorNotaBaixada(lerPlanilha):
    lerPlanilha = pd.read_excel(Formatacao().get_caminhoNotas(), Formatacao().get_sheetnameNotas(), dtype={'Documento': str})
    lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)

    cont = 0
    cont_baixada = 0

    for x in range(int(formatacao.quantidadeNotas())):
        
        cont += 1
        nome_cliente = lerPlanilha['Nome'][x]
        caminho_pdf =f'C:\\Users\\kaion\\Downloads\\teste\\{nome_cliente}.pdf'

        if os.path.exists(caminho_pdf):
            cont_baixada += 1
            print(f"       {cont} - ✅ JÁ EXISTE DOWLOAD PARA ESSE CLIENTE {nome_cliente}")
            lerPlanilha.drop(index=[x], inplace=True)
        else:
            print(f"       {cont} - ❌ NÃO EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
            
    if cont_baixada > 0:
        print('\n')
        alerta = int(input(f"Alerta! Você possui {cont_baixada} notas já baixadas!\nAperte: \n[1] - RETIRAR DA PLANILHA\n[2] - CONTINUAR SEM ALTERAR "))
        if alerta == 1:
            print('Alterado Na Planilha...')
    
    return lerPlanilha

