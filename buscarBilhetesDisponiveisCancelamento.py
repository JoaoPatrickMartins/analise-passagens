import pyautogui
import pyperclip
import time
import pandas as pd

def verificar_disponibilidade_cancelamento(df, df_copy):
    try:
        # Varre a planilha linha a linha para validar a disponibilidade para cancelamento
        for index, row in df.iterrows():
            form = row.iloc[2]
            valor = row.iloc[3]  # iloc[3] refere-se à quarta coluna
            data = row.iloc[0]  # iloc[0] refere-se à primeira coluna
            
            print(f"Form: {form} - Valor: {valor} - Data: {data}")
            localizador = form + valor

            # Clicar na caixa (ajuste as coordenadas conforme necessário)
            pyautogui.click(x=1767, y=478)
            pyautogui.click(x=19, y=34)

            # Digitar "n" ir cancelamento de bilhetes
            pyautogui.write('n')

            # Adicione uma pausa para garantir que o sistema tenha tempo de responder
            time.sleep(1)
            # Digitar 3 tabs para chegar no campo de localizador
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')

            # Digitar o número do localizador
            pyautogui.write(localizador)

            # Pressionar F12 para realizar a consulta
            pyautogui.press('f12')

            # Adicione uma pausa para garantir que o sistema tenha tempo de responder
            time.sleep(1)

            # Pressionar F5 fechar popup
            pyautogui.press('f5')

            # Pressionar control + a para selecionar o texto
            pyautogui.hotkey('ctrl', 'a')

            # Limpar a área de transferência
            pyperclip.copy('')

            # Copiar o texto selecionado para a área de transferência
            pyautogui.hotkey('ctrl', 'c')

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

        # Salvar o DataFrame em um novo arquivo Excel apenas com status
        df_copy.to_excel(r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\resultado_nao_validado.xlsx', index=False)

        # Mensagem de conclusão da primeira parte
        print("Data e hora atual:", time.ctime())
        print("Analise de cancelamento concluída com sucesso.")
    except Exception as e:
        print(f"Ocorreu um erro durante a execução do processo de verificação da disponibilidade para cancelamento: {e}")
