import pandas as pd
import time

def tratar_planilha_bilhetes_vendidos(caminho_arquivo):
  # Função para imprimir o valor
  def print_valor(data, valor):
    print(f"Data: {data} - Valor: {valor}")

  # Abrindo e tratando planilha de bilhetes vendidos 
  print("Data e hora atual:", time.ctime())
  print("Abrindo e tratando planilha de bilhetes vendidos...")

  # Abrir a planilha Excel
  try:
    df = pd.read_excel(caminho_arquivo)
  except FileNotFoundError:
    print(f"Arquivo não encontrado: {caminho_arquivo}")
    return
  except ImportError as e:
    print(f"Erro ao importar bibliotecas necessárias: {e}")
    return

  # Verificar se há pelo menos 4 colunas no DataFrame
  if df.shape[1] < 4:
    print("O DataFrame não possui pelo menos 4 colunas.")
    return

  # Dropar as 5 primeiras linhas
  df = df.drop(range(5))

  # Adicionar nome às colunas e uma coluna adicional chamada status
  df.columns = ['Data Venda', 'Pax', 'Form', 'Nr. Doc', 'Rota Resumida', 'Total Tarifa']
  df['Status'] = ''
  df['Cancelamento válido'] = ''

  # Remover todas as linhas vazias
  df = df.dropna(how='all')

  # Tratamento da planilha de bilhetes vendidos finalizado 
  print("Data e hora atual:", time.ctime())
  print("Tratamento da planilha de bilhetes vendidos finalizado.")

  # Varre a planilha linha a linha e obtém o valor da quarta coluna 
  for index, row in df.iterrows():
    # Converter o campo da coluna 'Data Venda' para string
    df['Data Venda'] = df['Data Venda'].astype(str)

    #converter o campo da coluna form para string
    df['Form'] = df['Form'].astype(str)

    #converter o campo da coluna valor para string
    df['Nr. Doc'] = df['Nr. Doc'].astype(str)

  return df
