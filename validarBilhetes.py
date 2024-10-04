import pyautogui
import pyperclip
import time
import pandas as pd

from calcularDataInicioFim import calcular_data_inicio_fim

def validar_cancelamento(df):
  try:
    # Varre a planilha linha a linha para validar o cancelamento
    for index, row in df.iterrows():
      form = row.iloc[2]
      valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
      data = row.iloc[0]  # iloc[0] refere-se à primeira coluna

      localizador = form + valor

      # Clicar na caixa (ajuste as coordenadas conforme necessário)
      pyautogui.click(x=1767, y=478)
      pyautogui.click(x=475, y=34)

      # Digitar "b" para consulta de bilhetes
      pyautogui.write('b')

      # Adicione uma pausa para garantir que o sistema tenha tempo de responder
      time.sleep(1)

      # Digitar 2 tabs para chegar no campo de localizador
      pyautogui.press('tab')
      pyautogui.press('tab')

      # Digitar o número do localizador
      pyautogui.write(localizador)

      # Tab para ir para o campo de data inicio
      pyautogui.press('tab')

      # Calcular data de inicio e de fim
      data_inicio, data_fim = calcular_data_inicio_fim(data)

      # Digitar a data de início
      pyautogui.write(data_inicio)

      # Digitar 2 tabs para ir para o campo de data fim
      pyautogui.press('tab')
      pyautogui.press('tab')

      # Digitar a data de fim
      pyautogui.write(data_fim)

      # Digitar 8 tabs para chegar no campo agencia 
      for _ in range(8):
        pyautogui.press('tab')

      # Digitar a agencia
      pyautogui.write('02031 - SMARTTOUR AGÊNCIA DE TURISMO')

      # Digitar 1 enter pra confirmar a agencia
      pyautogui.press('enter')

      # Pressionar F12 para realizar a consulta
      pyautogui.press('f12')

      # Adicione uma pausa para garantir que o sistema tenha tempo de responder
      time.sleep(2)

      # Pressionar F5 para fechar popup
      pyautogui.press('f5')

      # Digitar 2 tabs para ir para o campo localizador
      pyautogui.press('tab')
      pyautogui.press('tab')

      # Pressionar control + a para selecionar o texto
      pyautogui.hotkey('ctrl', 'a')

      # Limpar a área de transferência
      pyperclip.copy('')

      # Adicione uma pausa para garantir que a área de transferência esteja limpa
      #time.sleep(1)

      # Copiar o texto selecionado para a área de transferência
      pyautogui.hotkey('ctrl', 'c')

      # Adicione uma pausa para garantir que o sistema tenha tempo de copiar o texto
      #time.sleep(1)

      # Obter o texto da área de transferência
      selected_text = pyperclip.paste()

      # Verificar se há texto selecionado
      if selected_text:
        # Adicionar status ao dataframe
        df.loc[index, 'Cancelamento válido'] = 'invalido'

        # Pressionar F5
        pyautogui.press('f5')
      else:
        # Adicionar status ao dataframe
        df.loc[index, 'Cancelamento válido'] = 'valido'

    # Digitar F5 para fechar popup
    pyautogui.press('f5')

    # Salvar o DataFrame em um novo arquivo Excel apenas com cancelamento válido
    df.to_excel(r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\resultado_validado.xlsx', index=False)

    # Mensagem de conclusão
    print("Data e hora atual:", time.ctime())
    print("Processo concluído com sucesso.")
  except Exception as e:
    print(f"Ocorreu um erro durante a execução do processo de validação de cancelamento: {e}")