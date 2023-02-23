# Código de automação de tarefas para enviar e-mail.  Cometário

import pyautogui
import pyperclip
import time
pyautogui.PAUSE = 1

# Passo 1 Entra no sistema da Empresa (No link do Drive) Comentário

pyautogui.hotkey('ctrl', 't')
pyperclip.copy('httpsdrive.google.comdrivefolders149xknr9JvrlEnhNWO49zPcw0PW5icxgausp=sharing')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

time.sleep(5)


# Passo 2 Navegar até o local do relatório(Entra na pasta exporta) Comentário

pyautogui.click(x=417, y=314,clicks=2)
#pyautogui.click(x=1249, y=20,clicks=1)
time.sleep(2)



# Passo 3 Exporta o relatório (fazer o download) Comentário

pyautogui.click(x=425, y=402)
pyautogui.click(x=1216, y=190)
pyautogui.click(x=1029, y=621,clicks=1)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(5)


# Passo 4 Calcular os indicadores(faturamento e quantidade de produtos) Comentário

import pandas as pd
tabela = pd.read_excel(r'CUsersdiegoDownloadsVendas - Dez.xlsx')
display(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
print(faturamento)
print(quantidade)



# Passo 5 Enviar o email para Diretória  Comentário

# abrir aba e entrar no email
pyautogui.hotkey('ctrl', 't') # abre a aba
pyperclip.copy('httpsmail.google.commailu1#inbox')
pyautogui.hotkey('ctrl','v')
pyautogui.press('enter')

time.sleep(5)
# clicar no botão escrever
pyautogui.click(x=93, y=197)
# preencher as informações do email

#distinatário
pyautogui.write(diegoreis.ads@gmail.com)
pyautogui.press('tab')
pyautogui.write('diegoreis.edf@gmail.com')
pyautogui.press('tab')


#assunto
pyautogui.click(x=922, y=390)
pyperclip.copy('Só um teste!')
pyautogui.hotkey('ctrl', 'v')
pyautogui.click(x=841, y=526)
time.sleep(2)




# corpo do e-mail

texto =f'''Boa tarde, meu amigo

Essa aula está sendo muito proveitosa,
não esperava aprender tanto!

O faturamento de ontem foi R${faturamento,.2f}
A quantidade de produtos foi deR${quantidade,}

Edmilson boa tarde, estou fazendo um teste, estou fazendo automação de envio de e-mail, agora deu certo!

Abs
Obrigado pela experiência.'''

pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.press('tab')
pyautogui.hotkey('tab','end')
pyautogui.hotkey('tab','end')
pyautogui.press('enter')
pyautogui.click(x=1351, y=373)
pyautogui.click(x=1351, y=581)
pyautogui.click(x=431, y=623,clicks=1)
time.sleep(2)
pyautogui.press('tab')
pyautogui.press('enter')

#enviar  o e-mail
pyautogui.press('tab')
pyautogui.press('enter')

time.sleep(2)



import time
time.sleep(20)
print (pyautogui.position())