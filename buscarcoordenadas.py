import pyautogui
import time
import keyboard  # Biblioteca para detectar teclas

print("Posicione o mouse e pressione 'p' para capturar as coordenadas. Pressione 'Esc' para sair.")

try:
    while True:
        # Verifica se a tecla 'p' foi pressionada
        if keyboard.is_pressed('p'):
            # Pega a posição atual do mouse
            x, y = pyautogui.position()
            print(f"Coordenadas: X: {x}, Y: {y}")
            time.sleep(1)  # Delay para evitar múltiplas capturas com um só pressionamento
        
        # Verifica se a tecla 'Esc' foi pressionada para sair
        if keyboard.is_pressed('esc'):
            print("Encerrando...")
            break
except KeyboardInterrupt:
    print("\nEncerrado.")
