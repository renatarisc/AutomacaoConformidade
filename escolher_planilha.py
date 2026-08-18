from google.oauth2.service_account import Credentials
import gspread # para manipular as planilhas do Drive

PASTA_DRIVE_ID = "1DPVCjPCSp34LCntHrKbK2E3J1b4qhw02" # pasta do Drive onde ficam as Planilhas de Controle mensais

# usado pelos scripts quando rodam sozinhos (PyCharm/terminal, fora do gui.py), sem uma planilha
# escolhida na interface - o nome muda todo mês, mas sem data certa, então continua manual aqui
NOME_PLANILHA_PADRAO = "07. Jul"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def listar():
    # lista as planilhas da pasta do Drive, da mais recentemente modificada pra mais antiga
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES)
    gc = gspread.authorize(credenciais)
    arquivos = gc.list_spreadsheet_files(folder_id=PASTA_DRIVE_ID)
    return sorted(arquivos, key=lambda arquivo: arquivo["modifiedTime"], reverse=True)
