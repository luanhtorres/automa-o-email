import time
from webbrowser import Opera
import pyautogui
import pyperclip
import pandas as pd
import openpyxl
import numpy


link = "link-do-arquivo-para-download"

pyautogui.PAUSE = 3
pyautogui.press("winleft")
pyautogui.write("opera")
pyautogui.press("enter")

pyperclip.copy(link)

pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=334, y=298, clicks=2, button='left')
pyautogui.click(x=334, y=298, clicks=2, button='right')
pyautogui.click(x=416, y=745)
time.sleep(10)

df = pd.read_excel(r'diretorio-do-arquivo-no-seu-computador')

faturamento = df['Valor Final'].sum()
quantidade = df['Quantidade'].sum()

pyautogui.press("winleft")
pyautogui.write("email")
pyautogui.press("enter")
pyautogui.click(x=104, y=100, button='left')
pyautogui.write("seu-email-aqui")
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')

assunto = "Relatório de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')

pyautogui.press('tab')
texto = f"""
Bom dia, tudo bom?
    
O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,} 

Obrigado, até mais.
"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')


