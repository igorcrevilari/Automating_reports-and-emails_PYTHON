import pyautogui
import pyperclip
from time import sleep
import pandas as pd

# Todas as X e Y dependem do modelo de tela do computador.


# ENTRANDO no drive com a tabela VENDAS
sleep(10)
pyautogui.PAUSE = 1
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("drive") # link do drive
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
sleep(5)  # Adaptar de acordo com a conexão de internet

# ACESSANDO tabela e fazendo DOWNLOAD
pyautogui.click(x=424, y=346)  # Selecionando tabela
sleep(2)

pyautogui.click(x=1159, y=196)  # Download
pyautogui.click(x=952, y=586)
pyautogui.click(x=741, y=510)
sleep(5)

# ABSTRAINDO DADOS
tabela = pd.read_excel(r"caminho do arquivo", sheet_name=None) # caminho do arquivo
print(tabela)

faturamento = tabela["Valor Final", None].sum()

quantidade = tabela["Quantidade", None].sum()
print(faturamento)
print(quantidade)
sleep(2)



# ENTRAR NO EMAIL (DEVE ESTAR LOGADO)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
sleep(5)

pyautogui.click(x=80, y=203)  # Abrindo caixa de envio
sleep(5)

pyautogui.write("email@gmail.com") # email escolhido
pyautogui.press("tab")  # seleciona o email
# escreve outro email
# tab
# escreve outro email
# tab

pyautogui.press("tab")  # pula pro campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")  # escrever o assunto
pyautogui.press("tab")  # pular pro corpo do email

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs,
Nome"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# clicar no botão enviar


# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")
