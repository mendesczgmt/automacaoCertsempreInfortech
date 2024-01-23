import datetime
import pandas as pd

class Formatacao:
    def __init__(self):
        self._data = datetime.date.today().strftime('%d/%m/%Y')
        self._caminhoNotas = 'C:\\Users\\Suporte\\OneDrive\\Área de Trabalho\\fazer.xlsx'
        self.sheetnameNotas = 'page1'

    def get_caminhoNotas(self):
        return self._caminhoNotas

    def get_sheetnameNotas(self):
        return self.sheetnameNotas

    def get_data(self):
        return self._data
               


lerPlanilha = pd.read_excel(Formatacao().get_caminhoNotas(), Formatacao().get_sheetnameNotas())

print(lerPlanilha)
