import datetime

import pandas as pd

lerPlanilha = pd.read_excel('C:\\Users\\Suporte\\Downloads\\PlanilhaSara.xlsx', sheet_name='16 a 20.10')

class Formatacao:
    def __init__(self):
        self._data = datetime.date.today().strftime('%d/%m/%Y')
    
    def get_data(self):
        return self._data

    def quantidadeNotas(self):
        tamanho = lerPlanilha.shape
        num_linhas = tamanho[0]
        return num_linhas


    
