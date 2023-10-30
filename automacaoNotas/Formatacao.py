import datetime

class Formatacao:
    def __init__(self):
        self._data = datetime.date.today().strftime('%d/%m/%Y')
    
    def get_data(self):
        return self._data
    
