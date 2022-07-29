#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[24]:


import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1.25

# 1. Entrar no sistema
pyautogui.hotkey('ctrl','t') #atalho para abrir uma nova guia
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing') #copia esse link
pyautogui.hotkey('ctrl','v') #cola o link
pyautogui.hotkey('enter') #pressiona enter

# 2. Entrar na pasta ''exportar''
time.sleep(4.5) #delay para esperar o site carregar
pyautogui.press('right',presses=2) #pressiona a seta direita duas vezes
pyautogui.hotkey('enter') #aperta enter
time.sleep(4.5) #delay para esperar o site carregar
pyautogui.press('right',presses=3) #pressiona a seta direita tres vezes
pyautogui.hotkey('enter') #aperta enter

# 3. Fazer o download da base de dados
time.sleep(5.5)#delay para aguardar o carregamento da pagina
pyautogui.click(x=98, y=158)#sequencia de cliques para fazer o download
time.sleep(5.5)#delay para aguardar o carregamento da pagina
pyautogui.click(x=225, y=433)#sequencia de cliques para fazer o download
pyautogui.moveTo(x=454, y=435)#sequencia de cliques para fazer o download
pyautogui.click(x=454, y=435)#sequencia de cliques para fazer o download

pyautogui.hotkey('alt')
pyautogui.hotkey('ctrl','j')
time.sleep(4.5)
pyautogui.click(x=479, y=243)


# 4. Calcular os indicadores

# 5. Entrar no email
# 6. Mandar um email para a diretoria


# In[25]:


# 4. Calcular os indicadores
import pandas

tabela = pandas.read_excel(r"C:\Users\felipe.mendes\Downloads\Vendas - Dez.xlsx")



# In[27]:


import pyautogui
import pyperclip
import time
from datetime import date

pyautogui.PAUSE = 1.5

pyautogui.hotkey('ctrl','t')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
time.sleep(4.5)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('enter')


time.sleep(4.5)
pyautogui.click(x=73, y=196)

pyautogui.write('felipe.mmmachadodegueregue@gmail.com')

pyautogui.hotkey('tab')
pyautogui.hotkey('tab')

today = date.today()
d2 = today.strftime("%d/%m/%Y")
pyautogui.write("Relatorio de Vendas do dia "+d2)
pyautogui.hotkey('tab')

quantidade = tabela["Quantidade"].sum()

faturamento = tabela["Valor Final"].sum()

text = f"""Prezados,

Segue os resultados das vendas de hoje:

Quantidade: {quantidade:,} itens

Faturamento: {faturamento:,.2f} R$

Atenciosamente,

Felipe Martins

"""

pyautogui.write(text)

pyautogui.hotkey('ctrl','enter')

