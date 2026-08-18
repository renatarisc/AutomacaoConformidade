from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from google.oauth2.service_account import Credentials
import gspread # para manipular as planilhas do Drive
import pandas as pd
import time

import carregar_cores_planilha
import escolher_planilha
import pintar_celula_planilha
import baixar_ne
import anexar_ne
import encaminhar_processo

def main(nome_planilha=None):
    # nome_planilha: passado pelo gui.py com a planilha escolhida na interface; rodando o
    # script sozinho (sem gui.py), usa escolher_planilha.NOME_PLANILHA_PADRAO
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
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    # aba = planilha.worksheet("RO")
    aba = planilha.worksheet("TesteRO")

    dados = pd.DataFrame(aba.get_all_records()) # get_all_records() usa a 1ª linha como cabeçalho e exige que cada coluna tenha nome único
    cores = carregar_cores_planilha.executar(aba) # chama a def
    # para descobrir a cor --> print(cores[(3, 1)]), sendo que 3,1 é célula A3. Retorno = (0, 1, 1) = azul "ciano"

    COLUNA_TRAMITADO = dados.columns.get_loc("TRAMITADO") + 1 # calculado pelo nome do cabeçalho, não fixo, pra não quebrar se a coluna mudar de lugar
    COLUNA_DESPACHO = dados.columns.get_loc("DESPACHO") + 1

    VERMELHO = (1, 0, 0)
    BRANCO = (1, 1, 1)
    AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255) # mesmo amarelo usado nos outros scripts do pipeline

    lista_processo = []
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        # pega as informações do empenho com célula CINZA: foi assinado e deve ser baixado no Siafi
        if cores.get((linha_planilha, 8)) == (0.8, 0.8, 0.8):
            conformado = {}
            conformado['linha_planilha'] = linha_planilha
            conformado['processo'] = (dados.loc[linha, "PROCESSO"])
            conformado['NE'] = (dados.loc[linha, "NE"]) # = 2026NE510016
            conformado['despacho'] = (dados.loc[linha, "DESPACHO"])
            lista_processo.append(conformado)

    # ------- Conecta o Selenium no navegador logado no Siafi -------
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"
    navegador_siafi = webdriver.Chrome(options=options)

    for linha in lista_processo:
        numero_NE = linha["NE"].replace("2026NE", "") # = 510016
        baixar_ne.executar(navegador_siafi, numero_NE)

    # ------- Abre o Chrome maximizado -------
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # detach=True p/ impedir que o Selenium feche o navegador ao terminar a execução ou quando ocorrer um erro
    navegador_suap = webdriver.Chrome(options=options)  # navegador controlado pelo Selenium | o Chrome tem mais compatibilidade com os sites
    navegador_suap.maximize_window()

    # ------- Entra na tela de login do Suap -------
    # navegador_suap.get("https://suap.iff.edu.br/accounts/login/?next=/")  # pode ser o caminho de um arquivo local
    navegador_suap.get("http://suap.dev.iff.edu.br/accounts/login/?next=/")

    # ------- Faz o login no Suap -------
    navegador_suap.find_element(By.ID, "id_username").send_keys("1882905")
    navegador_suap.find_element(By.ID, "id_password").send_keys("Aj250104!" + Keys.ENTER)
    time.sleep(15)

    for linha in lista_processo:

        processo = linha["processo"]
        NE = linha["NE"] # = 2026NE510016
        despacho = linha["despacho"]

        # ------- Localiza o processo e carrega na página -------
        campo_busca_rapida = WebDriverWait(navegador_suap, 30).until(EC.element_to_be_clickable((By.NAME, "q")))
        campo_busca_rapida.clear()
        campo_busca_rapida.send_keys(processo + Keys.ENTER)

        try:
            anexar_ne.executar(navegador_suap, NE)
            if despacho:
                encaminhar_processo.executar(navegador_suap, despacho)
                aba.update_cell(linha["linha_planilha"], COLUNA_TRAMITADO, "OK") # só marca se a tramitação realmente aconteceu
                pintar_celula_planilha.executar(aba, linha["linha_planilha"], 8, BRANCO) # anexou e tramitou: NE totalmente concluída
            else:
                # anexou mas não tem despacho pra tramitar: NE concluída mesmo assim, e o despacho fica
                # sinalizado de amarelo pra alguém preencher manualmente depois
                pintar_celula_planilha.executar(aba, linha["linha_planilha"], 8, BRANCO)
                pintar_celula_planilha.executar(aba, linha["linha_planilha"], COLUNA_DESPACHO, AMARELO_CLARO_1)

            navegador_suap.get("http://suap.dev.iff.edu.br/")  # volta para a tela de início, onde tem o campo Busca rápida
            # navegador_suap.get("http://suap.iff.edu.br/")
            time.sleep(10)

        except Exception as e:
            # não interrompe mais o script inteiro: pinta de vermelho essa NE e segue para as próximas, igual ao baixar_ob.py
            print(f"Erro no empenho {NE} ao executar anexarNE: {e}")
            pintar_celula_planilha.executar(aba, linha["linha_planilha"], 8, VERMELHO)

if __name__ == "__main__":
    main()
