import pandas as pd
import pyautogui
import time
import pyperclip


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

# adicionar nome as colunas e uma coluna adicional chamada status
df.columns = ['Data Venda', 'Pax', 'Form', 'Nr. Doc', 'Rota Resumida', 'Total Tarifa']			

df['Status'] = ''

# Varre a planilha linha a linha e obtém o valor da quarta coluna 
# print temporário do dataframe
for index, row in df.iterrows():
    form = row.iloc[2]
    valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
    data = row.iloc[0]  # iloc[0] refere-se à primeira coluna

    localizador = form + valor
    print_valor(data, localizador)    

# iniciar acesso ao sistema
# Pausar entre as ações para dar tempo ao sistema de responder
pyautogui.PAUSE = 1

# Minimizar todas as janelas e ir para a área de trabalho (atalho Win + D)
pyautogui.hotkey('win', 'd')

# Adicione uma pausa para garantir que o sistema tenha tempo de responder
time.sleep(1)

# Clicar no ícone do programa na área de trabalho (ajuste as coordenadas conforme necessário)
pyautogui.click(x= 1851, y= 37)  # Coordenadas do programa

# Digitar 1 enter
pyautogui.press('enter')

# Adicione uma pausa para garantir que o sistema tenha tempo de responder
time.sleep(5)

# Clicar no checkbox  e executar
# Digitar 2 tabs
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar espaço
pyautogui.press('space')

# Digitar 1 enter
pyautogui.press('enter')

# Clicar no checkbox  e executar
# Digitar 2 tabs
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar espaço
pyautogui.press('space')

# Digitar 1 enter
pyautogui.press('enter')

#login
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar para 30 segundos para garantir que o sistema tenha tempo de responder
time.sleep(10)
# Digitar o usuário
pyautogui.write('T063971')

# Digitar 1 tab
pyautogui.press('tab')

# Digitar a senha
pyautogui.write('12345678')

# Digitar 1 tab
pyautogui.press('tab')

# Digitar 1 enter
pyautogui.press('enter')

# selecionar viacao
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar para 30 segundos para garantir que o sistema tenha tempo de responder
time.sleep(10)

# Digitar 2 tabs
pyautogui.press('tab')
pyautogui.press('tab')

# Digitar 1 seta para baixo
pyautogui.press('down')

# Digitar 1 tab
pyautogui.press('tab')

# Digitar 1 enter
pyautogui.press('enter')

#confirmar msg de aguardar para acesso
# Digitar 1 enter
pyautogui.press('enter')

#abrir menu de consulta de bilhetes
# Adicione uma pausa para garantir que o sistema tenha tempo de responder
# Ajustar pra 60 segundos para garantir que o sistema tenha tempo de responder
time.sleep(30)

# Varre a planilha linha a linha e obtém o valor da quarta coluna
for index, row in df.iterrows():
    form = row.iloc[2]
    valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
    data = row.iloc[0]  # iloc[0] refere-se à primeira coluna

    localizador = form + valor

    # Clicar na caixa (ajuste as coordenadas conforme necessário)
    pyautogui.click(x=1767, y=478)
    pyautogui.click(x=19, y=34)

    # Digitar "b"
    pyautogui.write('n')

    # parte com loop das consultas
    # Adicione uma pausa para garantir que o sistema tenha tempo de responder
    time.sleep(2)
    # # Digitar 2 tabs
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')

    # Digitar o número do localizador
    # OUFBERGG localizador existente no sistema
    # GUFBEXZG localizador inexistente no sistema
    pyautogui.write(localizador)

    # Pressionar F12
    pyautogui.press('f12')

    # Adicione uma pausa para garantir que o sistema tenha tempo de responder
    # Ajustar para 5 segundos para garantir que o sistema tenha tempo de responder
    time.sleep(1)

    # Pressionar F5
    pyautogui.press('f5')

    # Pressionar control + a
    pyautogui.hotkey('ctrl', 'a')


    # Limpar a área de transferência
    pyperclip.copy('')

    # Adicione uma pausa para garantir que a área de transferência esteja limpa
    time.sleep(1)

    # Copiar o texto selecionado para a área de transferência
    pyautogui.hotkey('ctrl', 'c')

    # Adicione uma pausa para garantir que o sistema tenha tempo de copiar o texto
    time.sleep(1)

    # Obter o texto da área de transferência
    selected_text = pyperclip.paste()

    # Verificar se há texto selecionado
    if selected_text:
        # adicionar status ao dataframe
        df.loc[index, 'Status'] = 'inexistente'

        # Pressionar F5
        pyautogui.press('f5')
    else:
        # adicionar status ao dataframe
        df.loc[index, 'Status'] = 'existente'

# Print temporário do DataFrame
print(df)

# Salvar o DataFrame em um novo arquivo Excel
df.to_excel(r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\resultado.xlsx', index=False)
