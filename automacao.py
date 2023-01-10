import pyautogui as pa
import pyperclip
import time
import pandas as pd


pa.PAUSE = 2

# Primeiro passo: Acessar o navegador

pa.press("win")
pa.write("chrome")
pa.press('enter')

# Segundo passo: acessar a base de dados

pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pa.hotkey("ctrl","v")
pa.press('enter')
time.sleep(3)
pa.click(x=452, y=302, clicks=2)
time.sleep(3)

#Terceiro passo: Baixar a base de dados

pa.click(x=402, y=398)
time.sleep(2)
pa.click(x=1716, y=190)
time.sleep(2)
pa.click(x=1551, y=622)
time.sleep(10)



#Quarto passo: importar os dados e realizar uma breve análise

tabela = pd.read_excel(r"C:/Users/06471551177/Downloads/Vendas - Dez.xlsx")
Total_faturamento = tabela["Valor Final"].sum()
Total_de_produtos = tabela["Quantidade"].sum()

#Quinto passo: Escrever email inserindo os dados análisados

pa.press("win")
pa.write("chrome")
pa.press('enter')
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pa.hotkey("ctrl","v")
pa.press('enter')
pa.click(x=154, y=206)
pyperclip.copy("carlosnielsousasilva@gmail.com")
pa.hotkey("ctrl","v")
pa.press("tab")

#Sexto passo: Escrever o assunto

pyperclip.copy("Ralatório de Vendas")
pa.hotkey("ctrl", "v")
pa.press("tab")

#Sétimo passo: Escrever o corpo do email
Texto_do_corpo = f'''Olá,
Segue relatório de vendas:
faturamento: R${Total_faturamento:,.2f}
Total de Produtos vendidos: {Total_de_produtos:,}.

Qualquer dúvida estou à disposição.
att.,
Carlos Daniel.'''

pyperclip.copy(Texto_do_corpo)
pa.hotkey("ctrl", "v")

#Oitavo passo: Enviar email
pa.press("tab")
pa.press("enter")
'''
time.sleep(5)
print(pa.position())'''