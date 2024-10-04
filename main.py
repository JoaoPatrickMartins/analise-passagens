import time
import pandas as pd
import pyautogui

from buscarBilhetesDisponiveisCancelamento import verificar_disponibilidade_cancelamento
from iniciarAcessoSistema import iniciar_acesso_sistema
from tratarPlanilhaBilhetesVendidos import tratar_planilha_bilhetes_vendidos
from validarBilhetes import validar_cancelamento

# Caminho do arquivo de entrada
caminho_arquivo = r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\entrada.xlsx'

# Função para tratar a planilha de bilhetes vendidos
df = tratar_planilha_bilhetes_vendidos(caminho_arquivo)

# Salvar o DataFrame em excel apos tratamento para teste
#df.to_excel(r'C:\Users\joaop\Joao Patrick\Trabalho\Digit Control\Program\analise-passagens\saida.xlsx', index=False)

# função para iniciar acesso ao sistema
iniciar_acesso_sistema()

# gerar uma copia do dataframe
df_copy = df.copy()

# Função para verificar a disponibilidade de cancelamento
verificar_disponibilidade_cancelamento(df, df_copy)

# iniciar validação de cancelamento
time.sleep(1) # Ajustar para garantir o acesso ao sistema corretamente 30 segundos para garantir que o sistema tenha tempo de responder

validar_cancelamento(df) # Função para validar o cancelamento

pyautogui.hotkey('win', 'd')# Minimizar todas as janelas e ir para a área de trabalho (atalho Win + D)




