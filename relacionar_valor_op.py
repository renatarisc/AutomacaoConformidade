from selenium import webdriver  # o webdriver é o motor de busca do Selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from google.oauth2.service_account import Credentials
import gspread # para manipular as planilhas do Drive
import pandas as pd
import time

import pintar_celula_planilha

def valor_brl_para_float(valor_str):
    # a numericise automática do gspread assume separador de milhar americano (vírgula)
    # e quebra valores no formato brasileiro (ex: "3.600,00" vira 3.6). Por isso os
    # valores são lidos como texto (numericise_ignore) e convertidos aqui manualmente
    valor_str = str(valor_str).strip().replace("R$", "").strip()
    valor_str = valor_str.replace(".", "").replace(",", ".")
    return float(valor_str)

def main():
    # Para controlar um Chrome já aberto, precisa iniciar o Chrome em modo de depuração remota (remote debugging)
    # e mandar o Selenium se conectar a ele:
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

    # porque cada tipo de valor tem sua própria coluna de valor e sua própria coluna de OB de destino
    COLUNAS_VALOR_OB = [
        {"nome_valor": "Valor ISS", "coluna_ob": 7, "nome_ob": "OB ISS"}, # Valor ISS -> OB ISS
        {"nome_valor": "Valor PG", "coluna_ob": 9, "nome_ob": "OB PG"},   # Valor PG -> OB PG
    ]

    AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255)

    lista_processo = []
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        for info in COLUNAS_VALOR_OB:
            valor_str = dados.loc[linha, info["nome_valor"]]
            if valor_str == "":
                continue

            # só considera se a OB correspondente ainda estiver em branco (ainda não foi relacionada)
            if dados.loc[linha, info["nome_ob"]] != "":
                continue

            conformado = {}
            conformado['valor'] = valor_brl_para_float(valor_str)
            conformado['linha_planilha'] = linha_planilha
            conformado['coluna_ob'] = info["coluna_ob"]
            lista_processo.append(conformado)

    # ------- Conecta o Selenium no navegador logado no Siafi -------
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"
    navegador = webdriver.Chrome(options=options)

    campo_comando = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, "frmMenu:acessoRapido")))
    campo_comando.send_keys("gerop" + Keys.ENTER)
    time.sleep(3)

    select_status = Select(navegador.find_element(By.ID, "formComp:status"))
    select_status.select_by_value("PENDENTE_ASSINATURA")  # = Pendente de Assinatura

    botao_pesquisar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, "formComp:botao_pesquisar")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_pesquisar)  # rola a página até o botão, pois ele precisa estar visível na tela para não dar erro
    botao_pesquisar.click()

    # ------- Espera o combo aparecer e altera para 50 OPs por página
    combo = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, "formComp:pageSizeInferior")))
    Select(combo).select_by_value("50") # se houver mais de 50 OPs, só aumentar o valor aqui
    time.sleep(3)
    WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.ID, "formComp:tabelaResultadoPesquisa:colunaSelecao")))

    ops_relacionadas = []

    for processo in lista_processo:
        valor_procurado = processo["valor"]

        # ------- Carrega a tabela dos valores x OPs do Siafi e pega todas as linhas dela
        WebDriverWait(navegador, 30).until(EC.presence_of_element_located((By.ID, "formComp:tabelaResultadoPesquisa:colunaSelecao")))
        linhas_tabela_navegador = navegador.find_elements(By.XPATH, "//table[contains(@id,'tabelaResultadoPesquisa')]//tbody/tr")

        encontrou = False

        for linha in linhas_tabela_navegador:

            valor = linha.find_element(By.XPATH, ".//span[contains(@id,'txtValor')]").text.strip()
            valor_num = valor_brl_para_float(valor)

            if valor_num == valor_procurado:
                OP = linha.find_element(By.XPATH, ".//a").text.strip()  # vai pegar o texto do link correspondente ao valor =158385/2026OP000303
                OP = OP.split("/")[1]  # = 2026OP000303
                aba.update_cell(processo["linha_planilha"], processo["coluna_ob"], OP) # grava a OP na coluna OB ISS ou OB PG, conforme a coluna do valor encontrado
                pintar_celula_planilha.executar(aba, processo["linha_planilha"], processo["coluna_ob"], AMARELO_CLARO_1) # pinta a célula da OP de amarelo claro 1
                ops_relacionadas.append(OP)
                encontrou = True
                break

        if not encontrou:
            aba.update_cell(processo["linha_planilha"], processo["coluna_ob"], "Não encontrado") # escreve na própria célula da OB, já que não achou o valor no Siafi

    print("OPs relacionadas:")
    for OP in ops_relacionadas:
        print(OP)

if __name__ == "__main__":
    main()
