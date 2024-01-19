import datetime
import pandas as pd

class Formatacao:
    def __init__(self):
        self._data = datetime.date.today().strftime('%d/%m/%Y')
        self._caminhoNotas = 'C:\\Users\\Suporte\\Downloads\\teste.xlsx'
        self.sheetnameNotas = '26.12 a 06.01'
        self._valor = 0
    
    def formatarValor(self, valor):
        self._valor = str(valor)
        self._valor += "00"
        self._valor = int(self._valor)
        return self._valor

    def get_data(self):
        return self._data

    def quantidadeNotas(self):
        tamanho = lerPlanilha.shape
        num_linhas = tamanho[0]
        return num_linhas
    
    def get_caminhoNotas(self):
        return self._caminhoNotas
    
    def get_sheetnameNotas(self):
        return self.sheetnameNotas

    def get_pegarValor(self):
        self._valor = str(self._valor)
        self._valor += "00"
        self._valor = int(self._valor)
        return self._valor

lerPlanilha = pd.read_excel(Formatacao().get_caminhoNotas(), Formatacao().get_sheetnameNotas(), dtype={'Documento': str})

