import base64

import openpyxl
from selenium.common import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.wait import WebDriverWait


# Passo 1: Criar ou abrir uma planilha Excel
arquivo_excel = 'informacoes_links.xlsx'
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Informações'

# Adicionar cabeçalho à planilha
ws.append(['Nome do Link', 'Quem Faz','Fonte da Imagem'])

# Configurar o WebDriver
driver = webdriver.Chrome()

# Acessar o site
driver.get("https://sei.utfpr.edu.br/sip/login.php?sigla_orgao_sistema=UTFPR&sigla_sistema=SEI")

# Esperar o site carregar completamente
# Preencher o campo de usuário
usuario = driver.find_element(By.ID, "txtUsuario")
usuario.send_keys("")

# Preencher o campo de senha
senha = driver.find_element(By.ID, "pwdSenha")
senha.send_keys("")

# Submeter o formulário de login
senha.send_keys(Keys.RETURN)

# Esperar o clique ser processado
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Base de Conhecimento"))).click()

# Esperar o botão "Minha Base" ser clicável
minha_base_button = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.ID, "btnMinhaBase"))
)
minha_base_button.click()

# Esperar a tabela estar disponível
tabela = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'infraTable'))
)

# Passo 4: Encontrar todas as linhas da tabela
linhas = tabela.find_elements(By.TAG_NAME, 'tr')

# Passo 5: Iterar sobre as linhas e ler os dados das células
for linha in linhas[1:]:  # Ignorar a primeira linha (cabeçalho)
    try:
        colunas = linha.find_elements(By.TAG_NAME, 'td')

      #  nome_ultima_alteracao= linha.find_element(By.XPATH, './td[3]/a')
        nome=colunas[2].text
        date_ultima_alteracao=colunas[3].text

       # date_ultima_alteracao=linha.find_element(By.XPATH,'./td[4]/a')
# Pegue o texto do elemento 'a' dentro da quarta coluna
       # nome = nome_ultima_alteracao.get_attribute('title')
        # Encontrar o link na linha
        link = linha.find_element(By.CLASS_NAME, 'ancoraSigla')

        # Obter o texto do link (nome)
        nome_link = link.text
        print(f'Nome do link: {nome_link}')

        # Clicar no link
        link.click()

        time.sleep(2)  # Esperar a página carregar

        abas = driver.window_handles  # Obtém todas as abas
        driver.switch_to.window(abas[-1])  # Muda para a última aba (a nova)

        quem_faz_lista = []  # Lista para armazenar os itens de "Quem faz"

        try:
            quem_faz = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//p[contains(text(), 'Quem faz?')]"))
            )

            # A partir do elemento "Quem faz?", localizar a lista (<ul>) associada
            lista = quem_faz.find_element(By.XPATH, "./following-sibling::ul")

            # Encontrar todos os itens da lista (<li>)
            itens = lista.find_elements(By.TAG_NAME, 'li')

            # Iterar sobre os itens e imprimir o texto
            for item in itens:
                texto_item = item.find_element(By.TAG_NAME, 'p').text
                print(f'Quem faz: {texto_item}')
                quem_faz_lista.append(texto_item)  # Adiciona o texto à lista

        except Exception as e:
            print(f'Erro: {e}')

        try:
            # Localiza o parágrafo com a classe "Texto_Justificado"
            paragrafo = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "Texto_Justificado"))
            )

            # Verifica se existe uma imagem dentro do parágrafo
            try:
                imagem = paragrafo.find_element(By.TAG_NAME, 'img')  # Procura por uma tag <img> dentro do parágrafo

                # Verifica se o 'src' da imagem começa com 'data:image/png;base64'
                if imagem.get_attribute("src").startswith("data:image/png;base64"):
                    data_url='sim'
                    print("Imagem com base64 encontrada dentro da classe 'Texto_Justificado'!")
                else:
                    data_url='nao'

                    print("Imagem encontrada dentro da classe 'Texto_Justificado', mas não está em base64.")

            except NoSuchElementException:
                data_url = 'nao'

                print("Nenhuma imagem encontrada dentro da classe 'Texto_Justificado'.")

        except NoSuchElementException:
            data_url='nao'
            print("Parágrafo com a classe 'Texto_Justificado' não encontrado.")

            # Gravar os dados na planilha
        ws.append([nome_link,', '.join(quem_faz_lista), data_url,nome,date_ultima_alteracao])


        time.sleep(2)  # Esperar a página carregar

        # Voltar para a aba original
        driver.close()  # Fecha a aba atual
        driver.switch_to.window(abas[0])
    except Exception as e:
        print(f'Erro ao processar a linha: {e}')

# Salvar a planilha Excel
wb.save(arquivo_excel)

# Fechar o navegador
driver.quit()
