import pandas as pd
import pyautogui
import time
import pyperclip

from calcularDataInicioFim import calcular_data_inicio_fim


# Função para imprimir o valor
def print_valor(data, valor):
    print(f"Data: {data} - Valor: {valor}")

# Caminho para o arquivo Excel
caminho_arquivo = r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\entrada.xlsx'

#inicio do programa
print("Data e hora atual:", time.ctime())
print("Iniciando programa...")

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

# adicionar nome as colunas e uma coluna adicional chamada status
df.columns = ['Data Venda', 'Pax', 'Form', 'Nr. Doc', 'Rota Resumida', 'Total Tarifa']			

df['Status'] = ''
df['Cancelamento válido'] = ''



# Varre a planilha linha a linha e obtém o valor da quarta coluna 
# print temporário do dataframe
# for index, row in df.iterrows():
#     form = row.iloc[2]
#     valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
#     data = row.iloc[0]  # iloc[0] refere-se à primeira coluna

#     localizador = form + valor
#     print_valor(data, localizador)    



# iniciar acesso ao sistema
print("Data e hora atual:", time.ctime())
print("Iniciando acesso ao sistema...")

# Pausar entre as ações para dar tempo ao sistema de responder
pyautogui.PAUSE = 1

# Minimizar todas as janelas e ir para a área de trabalho (atalho Win + D)
pyautogui.hotkey('win', 'd')

# Adicione uma pausa para garantir que o sistema tenha tempo de responder
time.sleep(1)

# Clicar no ícone do programa na área de trabalho (ajuste as coordenadas conforme necessário)
pyautogui.click(x= 1851, y= 37)  # Coordenadas do programa

# Digitar 1 enter para abrir o programa
pyautogui.press('enter')

# Adicione uma pausa para garantir que o sistema tenha tempo de responder
time.sleep(5)

# Clicar no checkbox  e executar
# Digitar 2 tabs para selecionar o checkbox
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar espaço para marcar o checkbox
pyautogui.press('space')

# Digitar 1 enter para executar
pyautogui.press('enter')

# Clicar no checkbox  e executar
# Digitar 2 tabs para selecionar o checkbox
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar espaço para marcar o checkbox
pyautogui.press('space')

# Digitar 1 enter para executar
pyautogui.press('enter')

#login
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar para 30 segundos para garantir que o sistema tenha tempo de responder
time.sleep(10)

# Digitar o usuário
pyautogui.write('T063971')

# Digitar 1 tab para ir para o campo de senha
pyautogui.press('tab')

# Digitar a senha
pyautogui.write('12345678')

# Digitar 1 tab para ir para o botão autenticar
pyautogui.press('tab')

# Digitar 1 enter para autenticar
pyautogui.press('enter')

# selecionar viacao
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar para 30 segundos para garantir que o sistema tenha tempo de responder
time.sleep(10)

# Digitar 2 tabs para ir para o campo de viacao
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar 1 seta para baixo para selecionar a viacao
pyautogui.press('down')

# Digitar 1 tab para ir para o botão de confirmar
pyautogui.press('tab')

# Digitar 1 enter para confirmar
pyautogui.press('enter')

#confirmar msg de aguardar para acesso
# Digitar 1 enter para confirmar
pyautogui.press('enter')

#abrir menu de consulta de bilhetes
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar pra 60 segundos para garantir que o sistema tenha tempo de responder
time.sleep(30)

# fim do processo de acesso ao sistema
print("Data e hora atual:", time.ctime())
print("Acesso ao sistema concluído com sucesso.")

# gerar uma copia do dataframe
df_copy = df.copy()

try:
    # Varre a planilha linha a linha para validar a disponibilidade para cancelamento
    for index, row in df.iterrows():
        form = row.iloc[2]
        valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
        data = row.iloc[0]  # iloc[0] refere-se à primeira coluna

        localizador = form + valor

        # Clicar na caixa (ajuste as coordenadas conforme necessário)
        pyautogui.click(x=1767, y=478)
        pyautogui.click(x=19, y=34)

        # Digitar "n" ir cancelamento de bilhetes
        pyautogui.write('n')

        # Adicione uma pausa para garantir que o sistema tenha tempo de responder
        time.sleep(1)
        # # Digitar 3 tabs para chegar no campo de localizador
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')

        # Digitar o número do localizador
        # OUFBERGG localizador existente no sistema
        # GUFBEXZG localizador inexistente no sistema
        pyautogui.write(localizador)

        # Pressionar F12 para realizar a consulta
        pyautogui.press('f12')

        # Adicione uma pausa para garantir que o sistema tenha tempo de responder
        # Ajustar para 5 segundos para garantir que o sistema tenha tempo de responder
        time.sleep(1)

        # Pressionar F5 fechar popup
        pyautogui.press('f5')

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
        df.loc[index, 'Status'] = 'existente'
        if selected_text:
            df.drop(index, inplace=True)
            df_copy.loc[index, 'Status'] = 'inexistente'

            # Pressionar F5 para fechar popup
            pyautogui.press('f5')
        else:
            df.loc[index, 'Status'] = 'existente'
            df_copy.loc[index, 'Status'] = 'existente'
                

    # Print temporário do DataFrame
    # print(df)

    # Salvar o DataFrame em um novo arquivo Excel apenas com status
    df_copy.to_excel(r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\resultado_nao_validado.xlsx', index=False)

    # Mensagem de conclusão da primeira parte
    print("Data e hora atual:", time.ctime())
    print("Analise de cancelamento concluída com sucesso.")
except Exception as e:
    print(f"Ocorreu um erro durante a execução do processo de verificação da disponibilidade para cancelamento: {e}")

# iniciar validação de cancelamento
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar para 30 segundos para garantir que o sistema tenha tempo de responder
time.sleep(1)

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




