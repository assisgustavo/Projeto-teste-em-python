import pyautogui
import time
import pyperclip
import pandas as pd

pyautogui.PAUSE = 3

# Passo 1 - Entrar no sistema da empresa
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")

# Passo 2 - Navegar até o local onde está a nossa base de dados
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(10)

# Passo 3 - Baixar a planilha de vendas
pyautogui.click(355, 260, clicks=2)
pyautogui.click(355, 260)
pyautogui.click(1160, 163)
pyautogui.click(1012, 524)

# Passo 4 - Calcular o faturamento e quantidade de produtos vendidos
tabela = pd.read_excel(r"C:/Users/Gustavo/Downloads/Vendas - Dez.xlsx")
print(tabela)

# Passo 5 - Calculando os valores
faturamento = tabela["Valor Final"].sum()
qtde_produtos = tabela["Quantidade"].sum()

# Passo 6 - Enviando o Email
# Abrir o gmail

# Enviar o email

pyautogui.hotkey("ctrl", "t")
pyautogui.write("mail.google.com")
pyautogui.press("Enter")
time.sleep(5)

# Clicar em Escrever(novo e-mail)
pyautogui.click(41, 190)

# Escrever o destinatario
pyautogui.write("assis_gustavo@hotmail.com")

# Selecionar o campo assunto
pyautogui.hotkey("tab")
pyautogui.hotkey("tab")
pyautogui.write("Relatorio de vendas de ontem")
pyautogui.hotkey("tab")

# Selecionar o campo corpo do email e escrever o email usando os indicadores calculados
texto = f""" 
Prezados, bom dia

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Abs
Gustavo"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
