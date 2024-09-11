import pandas as pd

# Função para imprimir o valor
def print_valor(data, valor):
    print(f"Data: {data} - Valor: {valor}")

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\teste.xlsx'

# Abrir a planilha Excel
try:
    df = pd.read_excel(caminho_arquivo)
except FileNotFoundError:
    print(f"Arquivo não encontrado: {caminho_arquivo}")
    exit()
except ImportError as e:
    print(f"Erro ao importar bibliotecas necessárias: {e}")
    exit()


# Verificar se há pelo menos 4 colunas no DataFrame
if df.shape[1] < 4:
    print("O DataFrame não possui pelo menos 4 colunas.")
    exit()

# Dropar as 5 primeiras linhas
df = df.drop(range(5))

# Varre a planilha linha a linha e obtém o valor da quarta coluna
for index, row in df.iterrows():
    valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
    data = row.iloc[0]  # iloc[0] refere-se à primeira coluna
    
    print_valor(data, valor)    
