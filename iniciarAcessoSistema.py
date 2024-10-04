import time
import pyautogui

def iniciar_acesso_sistema():
  # iniciar acesso ao sistema
  print("Data e hora atual:", time.ctime())
  print("Iniciando acesso ao sistema...")

  pyautogui.PAUSE = 1 # Pausar entre as ações para dar tempo ao sistema de responder

  pyautogui.hotkey('win', 'd') # Minimizar todas as janelas e ir para a área de trabalho (atalho Win + D)

  time.sleep(1) # Adicione uma pausa para garantir que o sistema tenha tempo de responder

  pyautogui.click(x=1851, y=37) # Clicar no ícone do programa na área de trabalho (ajuste as coordenadas conforme necessário)

  pyautogui.press('enter') # Digitar 1 enter para abrir o programa

  # Adicione uma pausa para garantir que o sistema tenha tempo de responder
  time.sleep(20)

  # Clicar no checkbox e executar
  # Digitar 2 tabs para selecionar o checkbox
  pyautogui.press('tab')
  pyautogui.press('tab')

  pyautogui.press('space') # Digitar espaço para marcar o checkbox

  pyautogui.press('enter') # Digitar 1 enter para executar

  # Clicar no checkbox e executar
  # Digitar 2 tabs para selecionar o checkbox
  pyautogui.press('tab')
  pyautogui.press('tab')

  pyautogui.press('space') # Digitar espaço para marcar o checkbox

  pyautogui.press('enter') # Digitar 1 enter para executar

  # login
  time.sleep(30) # Ajustar para garantir o acesso ao sistema corretamente 30 segundos para garantir que o sistema tenha tempo de responder

  pyautogui.write('T063971') # Digitar o usuário

  pyautogui.press('tab') # Digitar 1 tab para ir para o campo de senha

  pyautogui.write('12345678') # Digitar a senha

  pyautogui.press('tab') # Digitar 1 tab para ir para o botão autenticar

  pyautogui.press('enter') # Digitar 1 enter para autenticar

  # selecionar viacao
  time.sleep(30) # Ajustar para garantir o acesso ao sistema corretamente 60 segundos para garantir que o sistema tenha tempo de responder

  # Digitar 2 tabs para ir para o campo de viacao
  pyautogui.press('tab')
  pyautogui.press('tab')

  pyautogui.press('down') # Digitar 1 seta para baixo para selecionar a viacao

  pyautogui.press('tab') # Digitar 1 tab para ir para o botão de confirmar

  pyautogui.press('enter') # Digitar 1 enter para confirmar

  # confirmar msg de aguardar para acesso
  pyautogui.press('enter') # Digitar 1 enter para confirmar

  # abrir menu de consulta de bilhetes
  time.sleep(30) # Ajustar para garantir o acesso ao sistema corretamente 60 segundos para garantir que o sistema tenha tempo de responder

  # fim do processo de acesso ao sistema
  print("Data e hora atual:", time.ctime())
  print("Acesso ao sistema concluído com sucesso.")