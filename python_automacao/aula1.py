# Aula 1 - Automação com python

import pyautogui as gui
import pyperclip as clip
import pandas as pd
import time
import datetime

gui.PAUSE = 1

# Abre o navegador
gui.press('win')
gui.write('Google Chrome')
gui.press('enter')

# Abre o site
gui.hotkey('ctrl', 't')
clip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
gui.hotkey('ctrl', 'v')
gui.press('enter')

# Tempo para carregar o site
time.sleep(5)

# Abre o diretório
gui.click(x=390, y=341, clicks=2)
time.sleep(3)

# Clica no arquivo
gui.click(x=382, y=420)
time.sleep(0.5)

# Baixa o arquivo
"""gui.click(x=1710, y=224)
gui.click(x=1542, y=630, clicks=2)
time.sleep(7)
"""

# Abre o arquivo excel
dataframe = pd.read_excel(r'/home/ph/Downloads/Vendas - Dez.xlsx')
print(dataframe)

# Calcular os indicadores (faturamento, quantidade de produtos)
faturamento = dataframe['Valor Final'].sum()
quantidade = dataframe['Quantidade'].sum()
#print(faturamento, quantidade)

# Abre o navegador
"""gui.press('win')
gui.write('Google Chrome')
gui.press('enter')"""

# Abre o email
gui.hotkey('ctrl', 't')
clip.copy('https://mail.google.com/mail/u/0/#inbox')
gui.hotkey('ctrl', 'v')
gui.press('enter')
time.sleep(7)

# Compor o email com destinatário, assunto e corpo
gui.click(x=32, y=244)
time.sleep(10)

clip.copy('me.pauloalmeida@gmail.com')
gui.hotkey('ctrl', 'v')
gui.press('Tab')

clip.copy('Teste de automação')
gui.hotkey('ctrl', 'v')
gui.press('Tab')

text = f"""
Olá meu caro,

Este é um email contendo informações sobre o faturamento e quantidade de produtos vendidos fechados em {datetime.date.today()}.
Mas também é o primeiro email que eu mando com o programa de automação.

O faturamento é de R${faturamento:,.2f} e a quantidade de produtos vendidos é de {quantidade:,}.

Agradeço pela atenção.
Henrique, Paulo
"""
clip.copy(text)
gui.hotkey('ctrl', 'v')

# Enviar email
gui.hotkey('ctrl', 'enter')

# Fecha a aba do email
gui.hotkey('ctrl', 'w')