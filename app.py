import gspread  # biblioteca para conectar com a API Google Sheets e com as planilhas
from selenium import webdriver  # biblioteca do pacote Selenium para controlar o navegador
from selenium.webdriver.chrome.service import Service   # classe do pacote Selenium para criar um objeto Service que deve ser passado como parâmetro na hora de criar o objeto webdrivrer
from webdriver_manager.chrome import ChromeDriverManager    # classe para instalar o Drive Manager do Google Chrome em sua versão mais recente automaticamente
from selenium.webdriver.support.ui import WebDriverWait     # classe para mandar o programa esperar
from selenium.webdriver.common.by import By     # módulo para garantir que o selenium atenda à classe WebDriverWait usando algo como um XPATH para achar um elemento da página
from selenium.webdriver.support import expected_conditions as EC    # módulo do selenium com várias condições pré-definidas
from selenium.common.exceptions import TimeoutException, NoSuchElementException     # exceção caso o tempo acabe
import time     # biblioteca para fazer o programa parar
import random   # biblioteca para gerar números aleatórios
import os   # biblioteca para garantir que a chave .json seja corretamente aberta em todos os sistemas operacionais

numeros_checados = []   # lista que vai armazenar os números que já foram checados

gs = gspread.service_account(filename='./teste-gspread-406820-f7016171990b.json')    # cria um objeto usando a API Google Sheets para acessar o projeto da Google Cloud

print('+---------------------------------------------------+')
print('\nIniciando a automação. AVISO: O PROGRAMA COMEÇA DO INÍCIO DA PLANILHA, PORTANTO, CERTIFIQUE-SE DE QUE NÃO HÁ NÚMEROS QUE JÁ FORAM CHECADOS\n+-----------------------------------------------------+')

print('\nAbrindo o navegador ...')

chrome_drive_manager = ChromeDriverManager().install()  # baixa a última versão do Drive Manager do Chrome e armazena o caminho no objeto chrome_drive_manager. A instalação só precisa ser feita uma vez, depois o argumento da classe Service pode ser apenas o caminho do arquivo.
service = Service(chrome_drive_manager)     # cria um objeto Service com o argumento do caminho do ChromeDriverManager
navegador = webdriver.Chrome(service=service)   # cria o navegador a ser controlado

print('\nFaça o login no Whatsapp Web pelo QR code.')

navegador.get('https://web.whatsapp.com/')  # o método get() leva o navegador até um link

while len(navegador.find_elements(By.XPATH, '//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[5]/div/span')) < 1:     # tenta achar um elemento que só aparece quando o login é bem-sucedido
    navegador.implicitly_wait(2)    # espera 2 segundos

print('\nLogin bem-sucedido. Está tudo pronto.')

row = 2
while True:     # loop infinito

    sheets = gs.open_by_key('15ea91LhVlld9arN7pRt9QwXhr8CIOpGB9OsqhNGhHgM')   # abre todas as planilhas do projeto
    sheet = sheets.get_worksheet(1)    # seleciona a planilha que contém os números que usaremos

    while sheet.row_values(row) != []:   # loop for para acessar cada linha da planilha. Note que um loop 'for' poderia gerar um erro pois este poderia gerar um erro ao deixar o loop rodar mesmo a condição não sendo mais satisfeita devido a uma alteração durante o loop. Já o loop 'while' realiza a checagem da condição a cada loop

        nome = sheet.row_values(row)[2]    # pega o nome na linha da planilha
        numero = sheet.row_values(row)[3].replace('-', '').replace('+', '').replace('(', '').replace(')', '').replace(' ', '').replace('.', '').replace('_', '')     # pega o número na linha da planilha e retira alguns caracteres especiais

        if len(numero) < 12 and len(numero) > 9:     # caso o número esteja dentro do padrão (10 ou 11 caracteres)
            if numero not in numeros_checados:
                numeros_checados.append(numero)     # adiciona o número à lista de números checados
                numero_aleatorio = random.randint(30, 97) * random.randint(2, 9) + random.randint(50, 90)   # gera um número aleatório para código de cadastro
                mensagem = f'Oi {nome}, seja bem-vindo ao HCC e obrigada por se cadastrar! Seu código de cadastro é: HCC{numero_aleatorio}'

                navegador.get('https://web.whatsapp.com/send?phone=+55' + numero + '&text=' + mensagem)   # caminho para mandar uma mensagem pré-definida para alguém pelo número

                print(f'\nEnviando mensagem para o número {numero}')
                while True:     # esse loop só acaba quando a mensagem é enviada ou o botão para fechar a mensagem de erro é clicado
                    try:
                        botao_enviar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')))    # espera o botão de enviar se tornar clicável
                        botao_enviar.click()    # clica no botão de enviar
                        sheet.update_cell(row=row, col=5, value='1')    # coloca 1 na coluna 'Número já checkado? (1 = sim)'
                        sheet.update_cell(row=row, col=6, value=f'{numero_aleatorio}')
                        print(f'\nMensagem enviada para o número {numero} com sucesso!')
                        try:    # esse bloco try existe para tratar um pequeno alerta javascript de quando uma mensagem é enviada
                            alerta = WebDriverWait(navegador, 1).until(EC.alert_is_present())
                            alerta.dismiss()
                        except TimeoutException:
                            pass

                        break

                    except TimeoutException:    # caso o botão de enviar não apareça em 10 segundos
                        try:
                            botao_fechar_erro = navegador.find_element(By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/button')   # tenta achar o botão de fechar a mensagem de erro
                            botao_fechar_erro.click()   # clica no botão de fechar a mensagem de erro
                            print(f'\nO número {numero} é inválido. Deletando da planilha ...')
                            try:
                                sheet.delete_rows(row)      # deleta o número inválido da planilha
                                print(f'\nO número {numero} foi deletado da planilha com sucsesso!')

                            except Exception as error:
                                print(f'\nUm erro inesperado ocorreu ao tentar excluir o número {numero} da planilha. Erro: {error}')

                            row -= 1
                            break

                        except NoSuchElementException:
                            pass
                            
                        except:
                            pass

        else:
            print(f'\nO número {numero} é inválido. Deletando da planilha ...')
            try:
                sheet.delete_rows(row)      # deleta o número inválido da planilha
                print(f'\nO número {numero} foi deletado da planilha com sucsesso!')
                row -= 1

            except Exception as error:
                print(f'\nUm erro inesperado ocorreu ao tentar excluir o número {numero} da planilha. Erro: {error}')

        row += 1

    print('\nEsperando novos números serem adicionados ...')
    time.sleep(5)   # para o programa por 5 segundos. Essa parada é importante pois o Google APIs tem um limite de 300 leituras por minuto a uma planilha. Dessa forma, impedimos que esse limite seja ultrapassado