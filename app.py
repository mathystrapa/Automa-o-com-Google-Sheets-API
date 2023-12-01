import gspread  # biblioteca para conectar com a API Google Sheets e com as planilhas

credential = './teste-gspread-406820-f7016171990b.json'     # credencial .json

sheet_id = '15ea91LhVlld9arN7pRt9QwXhr8CIOpGB9OsqhNGhHgM'   # ID da planilha

gs = gspread.service_account(filename=credential)    # cria um objeto usando a API Google Sheets para acessar o projeto da Google Cloud
sheets = gs.open_by_key(sheet_id)   # abre todas as planilhas do projeto
sheet1 = sheets.get_worksheet(0)    # seleciona uma planilha pelo índice 0 (primeira)
sheet2 = sheets.get_worksheet(1)    # seleciona uma planilha pelo índice 1 (segunda)

print(sheet1)
print(sheet2)   # essa é a que contém os números