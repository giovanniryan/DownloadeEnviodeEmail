#!/usr/bin/env python
# coding: utf-8

# In[90]:


import pyautogui
import pyperclip
import time

# acessar o sitema (drive)
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v") 
pyautogui.hotkey("enter")

# acessar a base de dados
pyautogui.click(x=423, y=317, clicks=3)
time.sleep(2)

# exportar base de dados
pyautogui.click(x=416, y=367)
pyautogui.click(x=1161, y=182)
pyautogui.click(x=860, y=600)
time.sleep(3)



# <h2>Inserção da tabela</h2>
# <ul>
#     <li> Faturamento
#     <li> Qntd de produtos
# </ul> 

# In[91]:


# Calcular faturamento e quantidade de produtos 
import pandas as pd

tabela = pd.read_excel(r"C:\Users\giova\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
display(faturamento)
display(quantidade)


# <h2> Mandar pro E-mail </h2>

# In[92]:


# acessar o email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://outlook.live.com/mail/0/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)

# clicar para enviar um novo email
pyautogui.click(x=218, y=168)

# destinatário
pyautogui.write("giovannicunha14@outlook.com")
pyautogui.press("tab")

# assunto
pyperclip.copy("Revisão da base de dados")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

# descrição
texto = f"""
Prezados, bom dia.
O faturamento mensal foi de: R${faturamento:,}
A quantidade mensal foi de: {quantidade:,.2f}

Abraços,
Giovanni Ryan
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# inserir anexo
pyautogui.click(x=576, y=647)
pyautogui.click(x=836, y=538)
pyautogui.click(x=679, y=301)

# enviar e-mail
time.sleep(3)
pyautogui.click(x=352, y=658)





# In[93]:


time.sleep(7)
pyautogui.position()


# In[ ]:




