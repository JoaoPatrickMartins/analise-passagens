import pandas as pd

# Função para imprimir o valor
def print_valor(valor):
    print(valor)

# Abrir a planilha Excel
df = pd.read_excel('caminho/para/sua/planilha.xlsx')

# Varre a planilha linha a linha
for index, row in df.iterrows():
    valor = row['Nr. Doc']
    print_valor(valor)
