import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog
import pandas as pd

# Function to get user input
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    empresa = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")
    mes = simpledialog.askstring(title="Mês", prompt="Digite o mês (MMAAAA):")

    return empresa, mes

# Get user input
empresa, mes = get_user_input()

# Press Win + R
pyautogui.hotkey('win', 'r')
time.sleep(1)  # Wait for the Run dialog to open

# Type the path to the file
pyautogui.write('C:\\projeto\\UNICO.EXE.lnk')
time.sleep(1)  # Wait for the typing to complete

# Press Enter
pyautogui.press('enter')

# Wait for 10 seconds
time.sleep(10)

# Type 'contabil'
pyautogui.write('contabil')

# Press Tab
pyautogui.press('tab')

# Type '1234'
pyautogui.write('1234')

# Press Enter
pyautogui.press('enter')

# Wait for 5 seconds
time.sleep(5)

# Press Ctrl + B
pyautogui.hotkey('ctrl', 'b')

time.sleep(6)

# Type the company code
pyautogui.write(empresa)

pyautogui.press('enter')

# Type the date with "01" at the beginning
pyautogui.write('01' + mes)

pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

# Press Enter again
pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

# Click on specific coordinates
pyautogui.click(x=260, y=160)
pyautogui.click(x=627, y=411)

# Wait for 4 seconds
time.sleep(4)

# Click on specific coordinates
pyautogui.click(x=95, y=160)
pyautogui.click(x=97, y=282)

pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

pyautogui.press('enter')

# Wait for 2 seconds
time.sleep(2)

# Write the path with the company code and month
save_path = f'C:\\projeto\\planilhas\\balancete_{empresa}_{mes}.csv'
pyautogui.write(save_path)

pyautogui.press('enter')

# Wait for the file to be saved
time.sleep(10)  # Increased wait time to ensure the file is saved

# Load the CSV file with the correct encoding and separator
df = pd.read_csv(save_path, encoding='latin1', sep=';')

# Ensure the first column is treated as string
df.iloc[:, 0] = df.iloc[:, 0].astype(str)

# Check if the value '259' is in the first column
if '259' in df.iloc[:, 0].values:
    # Find the row with the number '259' in column A
    index_259 = df[df.iloc[:, 0] == '259'].index[0]

    # Keep only the rows up to the row with the number '259'
    df = df.iloc[:index_259 + 1]

    # Save the modified CSV file with the correct separator
    df.to_csv(save_path, index=False, encoding='latin1', sep=';')
else:
    print("O valor '259' não foi encontrado na coluna A.")

# Open the CSV file
pyautogui.hotkey('win', 'r')
time.sleep(1)
pyautogui.write(save_path)
pyautogui.press('enter')
time.sleep(4)

# Select all content in the CSV (Ctrl + T) and copy (Ctrl + C)
pyautogui.hotkey('ctrl', 't')
time.sleep(1)
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)

# Open Notepad
pyautogui.hotkey('win', 'r')
time.sleep(1)
pyautogui.write('notepad')
pyautogui.press('enter')
time.sleep(2)

# Paste the content into Notepad (Ctrl + V)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

# Select all content in Notepad (Ctrl + A) and copy (Ctrl + C)
pyautogui.hotkey('ctrl', 'a')
time.sleep(1)
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)

# Close Notepad without saving
pyautogui.hotkey('alt', 'f4')
time.sleep(1)

# Open the reconciliation file
pyautogui.hotkey('win', 'r')
time.sleep(2)
pyautogui.write('C:\\projeto\\planilhas\\CONCILIACAO_EMPRESA_XX_XXXX.xlsx')
pyautogui.press('enter')
time.sleep(5)

# Go to cell A1
pyautogui.hotkey('ctrl', 'home')
time.sleep(1)

# Paste the content (Ctrl + V)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)

pyautogui.press('f12')

time.sleep(2)

save_path_1 = f'C:\\projeto\\planilhas\\balancete\\CONCILIACAO_{empresa}_{mes}'
pyautogui.write(save_path_1)

pyautogui.press('enter')
