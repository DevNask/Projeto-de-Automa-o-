import pyautogui as bot
from time import sleep
import random
import os
import pygetwindow as gw
import pyperclip
import pandas as pd
import time

df = pd.read_excel(r'C:\Users\Detran\Desktop\robo (Conferência Baixa 2)\listaplaca.xlsx')  # extrai desta planilha

print(df['placa']) #apartir da coluna Placa

resultados = []  # Lista para armazenar os resultados
df_resultados = pd.DataFrame()  # DataFrame para armazenar os resultados
relatorio_path = r"C:\Users\Detran\Desktop\robo (Conferência Baixa 2)\relatorio-pepm.xlsx"  # Caminho do arquivo de relatório
print(f"Relatório gerado com sucesso! O arquivo está em: {relatorio_path}")

bot.FAILSAFE = True  # Ativa o recurso de segurança do PyAutoGUI

for placa in df ['placa']:
    sleep(2) #Aguarda 2s
    bot.moveTo(38,69)
    bot.click()
    bot.write ("pepm", interval=0.2) #digita pepm
    bot.write (str(placa), interval=0.2) #digita as placas
    bot.hotkey('enter') #pressiona enter
    
    sleep(2) #Aguarda 2s
    
    bot.moveTo(107,419)
    bot.click()
    bot.hotkey('ctrl', 'c') #Seleciona tudo
    
    sleep(2) #Aguarda 2s
    
    texto = pyperclip.paste().lower()  # Obtém o texto copiado e converte para minúsculo

    # 8. Verifica se a palavra "baixa permanente" está presente no texto copiado
    if "baixa permanente" in texto:
        resultados.append({"Placa": placa, "Status": "Baixa Permanente Encontrada"})
    else:
        resultados.append({"Placa": placa, "Status": "Veículo Ativo"})
            
    # Salva os dados a cada nova placa
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel(relatorio_path, index=False)

    # Pausa entre as iterações para tornar o comportamento mais natural
    time.sleep(random.uniform(1, 2))