import pandas as pd
import os
from Formatacao import Formatacao
formatacao = Formatacao()

def verificadorNotaBaixada():
    lerPlanilha = pd.read_excel(Formatacao().get_caminhoNotas(), Formatacao().get_sheetnameNotas(), dtype={'Documento': str})
    lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)

    cont = 0
    cont_baixada = 0
    listNotas = []
    
    for x in range(int(formatacao.quantidadeNotas())):
        
        cont += 1
        nome_cliente = lerPlanilha['Nome'][x]
        caminho_pdf =f'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\Notas fiscais\\ANO 2024\\01 - JANEIRO 2024\\{nome_cliente}.pdf'

        if os.path.exists(caminho_pdf):
            cont_baixada += 1
            print(f"       {cont} - ✅ JÁ EXISTE DOWLOAD PARA ESSE CLIENTE {nome_cliente}")
            listNotas = listNotas.append(nome_cliente)
        else:
            print(f"       {cont} - ❌ NÃO EXISTE DOWNLOAD PARA ESSE CLIENTE {nome_cliente}")
            
    if cont_baixada > 0:


        #achar a linha que tem o nome
        #excluir a linha do df
        print('\n')
        alerta = int(input(f"Alerta! Você possui {cont_baixada} notas já baixadas!\nAperte: \n[1] - ALTERAR NA PLANILHA\n [2] - CONTINUAR SEM ALTERAR "))
        if alerta == 1:
            lerPlanilha.drop()
        else:
            print('Continuando...')
        
        print(listNotas)

verificadorNotaBaixada()