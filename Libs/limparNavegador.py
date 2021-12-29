# 13/07/2021
# Iago Leonardo Alves de Oliveira

# Biblioteca desenvolvida para burlar a barreira do mercado livre
# Sempre que o site barra a navegação, essa função fará uma limpeza no navegador e atualizará a pagina
# Aguardará alguns segundos e depois continuará de onde parou

#Bibliotecas
#Time - usada para esperar um tempo
import time

#Função burla a barreira do mercado livre
def burlarBarreira(navegador):
    print('\n\t\tO BOT FOI BARRADO!\n\t\t(Parando de enviar solicitações por um período de tempo\n\t\tApós isso, continuará de onde parou)')

    #Recebe o link atual
    linkAtual = navegador.current_url
    
    #Volta para a pagina anterior
    navegador.back()

    #Recebe o link da pagina anterior
    linkAnterior = navegador.current_url

    #Deleta todos os cookies
    navegador.delete_all_cookies()
    #Aguarda alguns segundos
    time.sleep(60)
    print('\n\t\tAguarde mais 90 segundos...')
    time.sleep(90)
    print('\n\t\tContinuando de onde paramos...')

    #Abre uma nova aba
    navegador.execute_script('window.open();')

    #Volta para a aba antiga
    navegador.switch_to.window(navegador.window_handles[0])
    #Fecha a aba antiga
    navegador.close()

    #Atualiza qual aba o navegador está usando (senão dá conflito)
    navegador.switch_to.window(navegador.window_handles[0])

    #Acessa o link anterior
    navegador.get(linkAnterior)
    #Aguarda carregamento da pagina por completo
    navegador.implicitly_wait(2)
    
    #Deleta todos os cookies novamente
    navegador.delete_all_cookies()
    #Atualiza a pagina
    navegador.refresh()

    #Acessa o link em que estava antes de ser derrubado
    navegador.get(linkAtual)
    #Aguarda carregamento da pagina por completo
    navegador.implicitly_wait(2)