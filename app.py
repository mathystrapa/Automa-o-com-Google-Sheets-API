import gspread  # biblioteca para conectar com a API Google Sheets e com as planilhas
from selenium import webdriver  # biblioteca do pacote Selenium para controlar o navegador
from selenium.webdriver.chrome.service import Service   # classe do pacote Selenium para criar um objeto Service que deve ser passado como parâmetro na hora de criar o objeto webdrivrer
from webdriver_manager.chrome import ChromeDriverManager    # classe para instalar o Drive Manager do Google Chrome em sua versão mais recente automaticamente
from selenium.webdriver.support.ui import WebDriverWait     # classe para mandar o programa esperar
from selenium.webdriver.common.by import By     # módulo para garantir que o selenium atenda à classe WebDriverWait usando algo como um XPATH para achar um elemento da página
from selenium.webdriver.support import expected_conditions as EC    # módulo do selenium com várias condições pré-definidas

credential = './teste-gspread-406820-f7016171990b.json'     # credencial .json

sheet_id = '15ea91LhVlld9arN7pRt9QwXhr8CIOpGB9OsqhNGhHgM'   # ID da planilha

gs = gspread.service_account(filename=credential)    # cria um objeto usando a API Google Sheets para acessar o projeto da Google Cloud
sheets = gs.open_by_key(sheet_id)   # abre todas as planilhas do projeto
sheet1 = sheets.get_worksheet(0)    # seleciona uma planilha pelo índice 0 (primeira)
sheet2 = sheets.get_worksheet(1)    # seleciona uma planilha pelo índice 1 (segunda)

print(sheet1)   # outra planilha teste
print(sheet2)   # essa é a que contém os números

chrome_drive_manager = ChromeDriverManager().install()  # baixa a última versão do Drive Manager do Chrome e armazena o caminho no objeto chrome_drive_manager. A instalação só precisa ser feita uma vez, depois o argumento da classe Service pode ser apenas o caminho do arquivo.
service = Service(chrome_drive_manager)     # cria um objeto Service com o argumento do caminho do ChromeDriverManager
navegador = webdriver.Chrome(service=service)   # cria o navegador a ser controlado

print(chrome_drive_manager)     # mostra o caminho do Chrome Drive Manager instalado.

navegador.get('https://web.whatsapp.com/')  # o método get() leva o navegador até um link

try:

    element = WebDriverWait(navegador, 60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[5]/div/span')))     # espera um elemento específico que só aparece quando o usuário logou

    for i in range(5):

        phone = '41996746161'
        mensagem = f'mensagem automática de teste número {i}'
        navegador.get('https://web.whatsapp.com/send?phone=+55' + phone + '&text=' + mensagem)   # caminho para mandar uma mensagem pré-definida a um usuário pelo número

        try:
            botao_enviar = WebDriverWait(navegador, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')))     # espera o botão de enviar aparecer para ser clicado

            botao_enviar.click()    # clica no botão de enviar

            try:
                alerta = WebDriverWait(navegador, 1).until(EC.alert_is_present())   # espera um alerta javascript aparecer
                print('Alerta gerado:', alerta.text)

            except:
                print('Nenhum alerta presente.')
        
        except TimeoutError:
            print('O tempo para enviar a mensagem foi muito longo.')

except TimeoutError:
    print('O tempo de espera se esgotou e o usuário não fez o login. Por favor, faça o login em até 1 minuto.')