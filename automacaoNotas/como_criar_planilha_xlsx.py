import pandas as pd

d = {'NOME': ['Kaio', 'Ricardo', 'João'], 'TIPO':['e-CPF A1', 'e-CNPJ A1', 'e-CPF A3'], 'AGR':['Camila', 'Sabrina', 'Roberta'], 'VALOR':['200','200', '300']}

dados = pd.DataFrame(data=d)

print(dados)

dados.to_excel('dados.xlsx', index=False) # A parte do index retira os indices da planilha






