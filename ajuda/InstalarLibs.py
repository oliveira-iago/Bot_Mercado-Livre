import os
import time

os.system('pip install --user colorama') #Atualiza o PIP
os.system('cls')

# Cores
import colorama
colorama.init()
azul = '\033[1;94m' 
verde = '\033[1;92m'
amarelo = '\033[1;93m'
branco = '\033[1;97m'
limpa = '\033[0;0m'
#################

print(branco, '\n\t ========== Aguarde as bibliotecas serem baixadas =========\n', verde)

#Instala as bibliotecas no sistema operacional
os.system('pip install --upgrade --user --trusted-host pypi.org --trusted-host files.pythonhosted.org pip') #Atualiza o PIP
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org selenium') #Selenium (Ações no navegador)
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org webdriver_manager') #WebDriverManager (Instala o chrome driver automatico)
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org PyQt5') #PyQt5 (interface grafica)
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org PyQt5-tools') #ferramentas do PyQt5
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org openpyxl') #OpenPyXl (Ações no Excel)
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org tk') #Tkinter (Ações com o explorador de arquivos)
os.system('pip install --user --trusted-host pypi.org --trusted-host files.pythonhosted.org pyautogui') #PyAutoGui (Simula ações do usuário | tbm usado na criação de caixas de dialogo)


#Limpa o terminal
os.system('cls')

#Lista de bibliotecas
libs = [
    'Atualização do pip',
    'Selenium           - Para manipular o navegador',
    'WebDriverManager   - Para instalar o Chrome Driver da versão do seu chrome',
    'PyQt5 + tools      - Para a utilização de interfaces gráficas',
    'OpenPyXl           - Para a manipulação de planilhas do excel',
    'Tk (Tkinter)       - Para a manipulação do gerenciador de arquivos',
    'PyAutoGui          - Para criação de caixas de diálogos'
]

print(branco, '\n\t Foram instaladas as seguintes bibliotecas:\n', amarelo)

#Exibe as biblotecas instaladas
for lib in libs:
    print('\t', lib)

print(azul, '\n\t Tudo certo! Esta janela será fechada após 15 segundos')
time.sleep(15)
