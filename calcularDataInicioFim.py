from datetime import datetime, timedelta

# Função para calcular a data de início e de fim
def calcular_data_inicio_fim(data_str):
  # Verificar se data_str é uma string e converter para datetime se necessário
  if isinstance(data_str, str):
      data = datetime.strptime(data_str, '%d/%m/%Y %H:%M:%S')
  else:
      data = data_str
  
  # Calcular a data de início e a data de fim
  data_inicio = data - timedelta(days=3)
  data_fim = data + timedelta(days=3)
  
  # Formatar as datas de volta para string no formato desejado
  data_inicio_str = data_inicio.strftime('%d/%m/%Y')
  data_fim_str = data_fim.strftime('%d/%m/%Y')
  
  return data_inicio_str, data_fim_str