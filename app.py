import pyautogui
import time
import tkinter as tk
from tkinter import simpledialog

# Function to get user input
def get_user_input():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    empresa = simpledialog.askstring(title="Código da Empresa", prompt="Digite o código da empresa:")
    mes = simpledialog.askstring(title="Mês", prompt="Digite o mês (MM/AAAA):")

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

# Type the company code
pyautogui.write(empresa)

# Press Tab
pyautogui.press('tab')

# Type the month
pyautogui.write(mes)

# Press Enter
pyautogui.press('enter')