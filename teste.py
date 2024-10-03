import pyautogui
import time
import pyperclip

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
pyautogui.write('OUFBERGG')

# Pressionar F12
pyautogui.press('f12')

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
  print("localizador não encontrado")
else:
  print("localizador encontrado")


