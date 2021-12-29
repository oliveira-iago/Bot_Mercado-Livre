#### Importa as bibliotecas
#OS - Ações do sistema operacional (ja vem com python)
import os, sys
#Selenium - Ações no chrome (pip install selenium)
from selenium import webdriver
#WebDriverManager - Instalar chrome driver automatico (pip install webdriver_manager)
from webdriver_manager.chrome import ChromeDriverManager

import time

#Biblioteca necessária após a 
from selenium.webdriver.common.by import By


#Implementa configurações
options = webdriver.ChromeOptions()
options.add_argument('lang=pt-br') #Portugues
options.add_argument('--disable-notifications') #Sem notificações
#Aplica opções que removem o aviso 'o chrome está sendo controle por um software'
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option('useAutomationExtension', False)

#Baixa o chrome driver de acordo com a versão do chrome do usuario
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)

#Maximiza a tela
navegador.maximize_window()

navegador.get('https://lista.mercadolivre.com.br/_CustId_497429408')


itens = navegador.find_elements_by_class_name('ui-search-item__title ui-search-item__group__element')

print('itens:')
print(itens)

print(2)

itens = navegador.find_elements_by_class_name('ui-search-item__title ui-search-item__group__element')

print('itens:')
print(itens)

'''
print('loop:')
for item in itens:
    print(item.text)
    print('loopando...')

print('fim loop')
'''

time.sleep(500)