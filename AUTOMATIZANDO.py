import time
import pyautogui
import numpy
import pyperclip
import time
pyautogui.PAUSE = 1

#ir para área de trabalho
pyautogui.hotkey("win", "m")
time.sleep(2)


#Abri o navegador e abrir o arquivo
pyautogui.click(x=560, y=743)
time.sleep(3)
pyautogui.click(x=584, y=42)
time.sleep(2)
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
time.sleep(2)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)

#fazer dowload do arquivo
pyautogui.click(x=427, y=284, clicks=2)
time.sleep(3)
pyautogui.click(x=442, y=295, button='right')
time.sleep(3)
pyautogui.click(x=573, y=629)
time.sleep(4)
pyautogui.press("enter")

time.sleep(2)


#Calcular indicadores
import pandas as pd

tabela = pd.read_excel(r"D:\Intensivão Python\AULA 1\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()


#Enviar por e-mail
time.sleep(3)
pyautogui.click(x=448, y=48)
time.sleep(1)
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=65, y=158)
pyautogui.click(x=959, y=293)
time.sleep(1)
pyautogui.write("matheusbreciani20@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
time.sleep(1)
pyautogui.write("Relatório de hoje")
pyautogui.press("tab")


texto = f"""
Prezados, bom dia
o faturamento de hoje foi de: {faturamento}
a quantidade de produtos vendidos foi de: {quantidade}
Abs
Matheus Breciani"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
time.sleep(1)
pyautogui.press("enter")



"""time.sleep(5)
pyautogui.position()
print(pyautogui.position())"""