import pyautogui
import time
import pyperclip
import pandas as pd

caminho = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
caminho_arquivo = r"C:\Users\igor_\Downloads\Vendas - Dez.xlsx"

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# 1. ABRIR MENU WINDOWS
pyautogui.press("win")
time.sleep(1)

# 2. ABRIR EDGE
pyautogui.write("edge")
time.sleep(0.5)
pyautogui.press("enter")
time.sleep(2)

# 3. COLAR LINK
pyautogui.hotkey('ctrl', 'l')
time.sleep(0.5)

pyperclip.copy(caminho)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.5)

pyautogui.press("enter")
time.sleep(3)

# 4. CLICAR NO ARQUIVO (AJUSTE COORDENADAS SE NECESSÁRIO)
print("Aguardando carregamento do Google Drive...")
time.sleep(3)

pyautogui.click(x=350, y=273)
time.sleep(0.5)
pyautogui.doubleClick(x=350, y=273)
time.sleep(1)

pyautogui.click(x=393, y=274, button="right")
time.sleep(0.5)
pyautogui.click(x=506, y=335)

# 5. LER EXCEL
tabela = pd.read_excel(caminho_arquivo)

print(tabela)

faturamento = tabela["Valor Final"].sum()
qtd_total = tabela["Quantidade"].sum()

print("Faturamento total:", faturamento)
print("Quantidade total:", qtd_total)

 #2 ABRIR GMAIL

pyautogui.hotkey("ctrl", "l")
time.sleep(0.5)

pyautogui.write("https://mail.google.com/")
pyautogui.press("enter")
time.sleep(5)

pyautogui.click(x=117, y=185)

pyautogui.write("iguinhodabiz@gmail.com")
