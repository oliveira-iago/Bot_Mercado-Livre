# Iago Leonardo Alves de Oliveira
# 04/04/2021

# Biblioteca desenvolvida para buscar os numero de vendas e o estoque disponivel de um produto no mercado livre
# A funçao recebe o link do produto e o usa para acessá-lo usando o chrome driver
# Retorna o numero de vendas e estoque
# Caso o mercado livre barre a navegação, é usada uma função para burlar isto e continuar de onde parou

#Biblioteca exclusiva
from Libs.limparNavegador import burlarBarreira

#Função que busca a quantidade de vendas a partir de um link
def buscarVendasEstoque(link, navegador):
    
    #Vendas se inicia vazio
    vendas = ''
   
    #Acessa o anuncio do mercado livre
    navegador.get(link)

    #Aguarda carregamento da pagina por completo (1 segundo na vdd)
    navegador.implicitly_wait(2)

    try:
        #Busca o numero de vendas a partir da classe presente no codigo HTML da pagina
        vendas = navegador.find_element_by_class_name('ui-pdp-subtitle').text

    except:
        #Chama a função que recarrega a pagina e limpa os cookies
        burlarBarreira(navegador)

        try:
            #Busca o numero de vendas a partir da classe presente no codigo HTML da pagina
            vendas = navegador.find_element_by_class_name('ui-pdp-subtitle').text

        except:
            #Chama a função que recarrega a pagina e limpa os cookies
            burlarBarreira(navegador)

            #Busca o numero de vendas a partir da classe presente no codigo HTML da pagina
            vendas = navegador.find_element_by_class_name('ui-pdp-subtitle').text
        
    #Converte em texto
    vendas = str(vendas)
    
    #Formata o texto para deixar apenas os numeros
    vendas = vendas.replace('Novo  |  ', '')
    vendas = vendas.replace('Novo', '')
    vendas = vendas.replace(' vendidos', '')
    vendas = vendas.replace(' vendido', '')
    vendas = vendas.replace('Usado', '')
        
    #Se vendas estiver vazio, entao nao vendeu
    if vendas == '':
        vendas = '0'

    #Recebe o estoque disponivel do produto
    try:
        estoqueProduto = navegador.find_element_by_class_name('ui-pdp-buybox__quantity__available').text
        estoqueProduto = estoqueProduto.replace('(', '')
        estoqueProduto = estoqueProduto.replace(')', '')
    except:
        estoqueProduto = '1'

    #Remove os textos
    estoqueProduto = estoqueProduto.replace(' ', '').replace('disponíveis', '').replace('disponível', '')
    #Converte em numero
    estoqueProduto = int(estoqueProduto)

    #Verifica se o anuncio está pausado
    try:
        navegador.execute_script("document.getElementsByClassName('ui-pdp-message ui-vpp-message undefined ui-pdp-background-color--WHITE andes-message andes-message--warning andes-message--quiet')[0].textContent")
        vendas = vendas + ' (Pausado)'
    except:
        print()    

    #Retorna a quantidade de vendas e estoque
    return vendas, estoqueProduto