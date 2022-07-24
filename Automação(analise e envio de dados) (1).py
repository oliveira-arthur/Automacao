#!/usr/bin/env python
# coding: utf-8

# In[10]:


import pandas as pd
import pyautogui
import pyperclip
import time


# In[ ]:


pyautogui.Pause = 1
pyautogui.press('winleft')
pyautogui.write('Edge')
pyautogui.press('enter')
time.sleep(5)
pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(366,336, clicks = 2)
time.sleep(1)
pyautogui.click(437,349)
time.sleep(1)
pyautogui.click(1664,197)
time.sleep(1)
pyautogui.click(1402,704)
leitura = pd.read_excel(r'C:\Users\arthu\Downloads\Vendas - Dez.xlsx')


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[11]:


display(leitura)
faturamento = leitura['Valor Final'].sum()
quantidadededeprodutos = leitura ['Quantidade'].sum()


# In[12]:


#entrando no navegador
pyautogui.press('winleft')
pyautogui.write('Edge')
pyautogui.press('enter')
time.sleep(5)


# In[ ]:


#abrindo uma nova página de e-mail no navegador
pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')
time.sleep(5)


# In[ ]:


#clicando em novo e-mail e esrevendo o e-email para onde será enviado, assunto e o corpo do texto com as informações retiradas da tabela
pyautogui.click(126, 219)
time.sleep(2)
pyautogui.write("seugmail+diretoria@gmail.com")
pyautogui.press('tab')
pyautogui.press('tab')
assunto = ('Relatório de Vendas de Ontem')
time.sleep(2)
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
texto = f'''Prezados, Bom dia!
O faturamento de ontem foi R${faturamento:,.2f}
A numero de produtos foi de {quantidadededeprodutos}
 Arthur de Oliveira'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')


# In[ ]:


#clicando para enviar de fato o e-ma
pyautogui.hotkey('ctrl','enter')

