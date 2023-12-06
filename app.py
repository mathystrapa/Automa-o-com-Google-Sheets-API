import gspread  # biblioteca para conectar com a API Google Sheets e com as planilhas
from selenium import webdriver  # biblioteca do pacote Selenium para controlar o navegador
from selenium.webdriver.chrome.service import Service   # classe do pacote Selenium para criar um objeto Service que deve ser passado como parâmetro na hora de criar o objeto webdrivrer
from webdriver_manager.chrome import ChromeDriverManager    # classe para instalar o Drive Manager do Google Chrome em sua versão mais recente automaticamente
from selenium.webdriver.support.ui import WebDriverWait     # classe para mandar o programa esperar
from selenium.webdriver.common.by import By     # módulo para garantir que o selenium atenda à classe WebDriverWait usando algo como um XPATH para achar um elemento da página
from selenium.webdriver.support import expected_conditions as EC    # módulo do selenium com várias condições pré-definidas
from selenium.common.exceptions import TimeoutException     # exceção caso o tempo acabe
import time     # biblioteca para fazer o programa parar

numeros_checados = []   # lista que vai armazenar os números que já foram checados
    
def retornar_linhas_validas(planilha: gspread.worksheet.Worksheet):
    cont = 2
    while True:
        try:
            if planilha.row_values(cont) == []:
                return cont - 1
            cont += 1
        except:
            return planilha.row_count


credential = './teste-gspread-406820-f7016171990b.json'     # credencial .json
sheet_id = '15ea91LhVlld9arN7pRt9QwXhr8CIOpGB9OsqhNGhHgM'   # ID da planilha

gs = gspread.service_account(filename=credential)    # cria um objeto usando a API Google Sheets para acessar o projeto da Google Cloud

chrome_drive_manager = ChromeDriverManager().install()  # baixa a última versão do Drive Manager do Chrome e armazena o caminho no objeto chrome_drive_manager. A instalação só precisa ser feita uma vez, depois o argumento da classe Service pode ser apenas o caminho do arquivo.
service = Service(chrome_drive_manager)     # cria um objeto Service com o argumento do caminho do ChromeDriverManager
navegador = webdriver.Chrome(service=service)   # cria o navegador a ser controlado

print(chrome_drive_manager)     # mostra o caminho do Chrome Drive Manager instalado.

navegador.get('https://web.whatsapp.com/')  # o método get() leva o navegador até um link

while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[5]/div/span')) < 1:     # tenta achar um elemento que só aparece quando o login é bem-sucedido
    navegador.implicitly_wait(2)    # espera 2 segundos

while True:     # loop infinito

    sheets = gs.open_by_key(sheet_id)   # abre todas as planilhas do projeto
    sheet2 = sheets.get_worksheet(1)    # seleciona uma planilha pelo índice 1 (segunda). Essa é a planilha que contém os números que usaremos

    row = 2
    while row <= retornar_linhas_validas(sheet2):   # loop for para acessar cada linha da planilha. Note que um loop 'for' poderia gerar um erro pois este poderia gerar um erro ao deixar o loop rodar mesmo a condição não sendo mais satisfeita devido a uma alteração durante o loop. Já o loop 'while' realiza a checagem da condição a cada loop

        nome = sheet2.row_values(row)[2]    # pega o nome na linha da planilha
        numero = sheet2.row_values(row)[3].replace('-', '').replace('+', '').replace('(', '').replace(')', '').replace(' ', '').replace('.', '').replace('_', '')     # pega o número na linha da planilha e retira alguns caracteres especiais

        if len(numero) < 12 and len(numero) > 9:     # caso o número esteja dentro do padrão (10 ou 11 caracteres)
            if numero not in numeros_checados:
                numeros_checados.append(numero)     # adiciona o número à lista de números checados
                mensagem = f'Mensagem automática de teste para {nome}.'

                navegador.get('https://web.whatsapp.com/send?phone=+55' + numero + '&text=' + mensagem)   # caminho para mandar uma mensagem pré-definida para alguém pelo número

                while True:
                    try:
                        botao_enviar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')))
                        botao_enviar.click()
                        try:
                            alerta = WebDriverWait(navegador, 1).until(EC.alert_is_present())
                            alerta.dismiss()
                        except TimeoutException:
                            pass

                        break

                    except TimeoutException:
                        try:
                            botao_fechar_erro = navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')
                            botao_fechar_erro.click()
                            sheet2.delete_rows(row)
                            break
                        
                        except:
                            pass

        else:
            print(f'O número {numero} não é valido.')
            sheet2.delete_rows(row)      # deleta a linha com o número inválido

        while row == retornar_linhas_validas(sheet2):
            time.sleep(5)      # faz o programa parar pelo número de segundos passados no parâmetro. Isso é importante porque a biblioteca gspread da Google APIs tem um limite de leituras por minuto igual a 300 leituras. Portanto, caso o programa não pare após checar todos os números, mais de 300 leituras em um minuto serão feitas e a API será interrompida

        row += 1