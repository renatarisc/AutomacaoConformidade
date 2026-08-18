from selenium import webdriver # o webdriver é o motor de busca do Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from google.oauth2.service_account import Credentials
import gspread # para manipular as planilhas do Drive
import pandas as pd
import time # para fazer pausa

import pintar_celula_planilha

def main():
    # Para controlar um Chrome já aberto, precisa iniciar o Chrome em modo de depuração remota (remote debugging) e mandar o Selenium se conectar a ele
    # 1- Fecha todos os Chromes abertos
    # 2- No CMD: "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeSelenium"
    # 3- Faz login no Sistema que deseja (Siafi)
    # 4- Conecta o Selenium no navegador logado no Siafi (depois que pego os dados da Planilha de Controle)

    # ------- Acessa a Planilha de Controle da Conformidade (mensal) -------
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES) # nome do arq dentro da pasta do Projeto
    gc = gspread.authorize(credenciais)
    planilha = gc.open("07. Jul") # apenas o nome da planilha, não precisa indicar o caminho
    # aba = planilha.worksheet("RO")
    aba = planilha.worksheet("TesteNS")

    dados = pd.DataFrame(aba.get_all_records(numericise_ignore=['all'])) # get_all_records() usa a 1ª linha como cabeçalho e exige que cada coluna tenha nome único

    # as colunas OB ISS e OB PG guardam a OP relacionada pelo relacionar_valor_op.py; aqui ela é trocada pela OB correspondente
    COLUNAS_OB = [
        {"coluna": 7, "nome": "OB ISS"},
        {"coluna": 9, "nome": "OB PG"},
    ]

    VERMELHO = (1, 0, 0)
    AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255) # mesmo amarelo usado no relacionar_valor_op.py

    lista_processo = []
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        for info in COLUNAS_OB:
            valor_celula = str(dados.loc[linha, info["nome"]])

            if "OP" not in valor_celula: # ignora célula em branco, "Não encontrado" ou que já tenha uma OB
                continue

            lista_processo.append({
                "linha_planilha": linha_planilha,
                "coluna": info["coluna"],
                "OP": valor_celula,
            })

    # ------- Conecta o Selenium no navegador logado no Siafi -------
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"
    navegador = webdriver.Chrome(options=options)

    for processo in lista_processo:

        campo_comando = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.ID, "frmMenu:acessoRapido")))
        campo_comando.send_keys("gerop" + Keys.ENTER)
        time.sleep(3)

        numero_OP = processo["OP"].replace("2026OP", "")

        campo_NE = WebDriverWait(navegador, 15).until(EC.element_to_be_clickable((By.ID, "formComp:codigoOp_input")))
        campo_NE.clear()
        campo_NE.send_keys(numero_OP)
        campo_NE.send_keys(Keys.ENTER)

        try:
            nome_OB = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, "formComp:tableDocSiafiDetalhe:1:linkDocumentoHOD_lnkConsultaDoc")))
            navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", nome_OB)  # rola a página até o link, pois ele precisa estar visível na tela para não dar erro
            time.sleep(2)  # aguarda 2s para a rolagem terminar
            OB = nome_OB.text # pega o texto do link = 158385/2026OB000303
            OB = OB.split("/")[1] # = 2026OB000303

        except TimeoutException:
            # não achou a OB no Siafi ainda: pinta de vermelho mas MANTÉM a OP na célula, para tentar de novo numa próxima varredura
            pintar_celula_planilha.executar(aba, processo["linha_planilha"], processo["coluna"], VERMELHO)
            continue

        aba.update_cell(processo["linha_planilha"], processo["coluna"], OB) # substitui a OP pela OB na mesma célula (OB ISS ou OB PG)
        pintar_celula_planilha.executar(aba, processo["linha_planilha"], processo["coluna"], AMARELO_CLARO_1) # mantém o amarelo: sinaliza que a OB ainda precisa ser baixada

if __name__ == "__main__":
    main()
