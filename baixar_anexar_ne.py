from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from google.oauth2.service_account import Credentials
from pathlib import Path
import gspread # para manipular as planilhas do Drive
import time

import carregar_cores_planilha
import credenciais_suap
import escolher_planilha
import ler_planilha
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
    aba = planilha.worksheet("RO")

    dados = ler_planilha.carregar_registros(aba, preencher_mescladas=True) # 1ª linha vira cabeçalho, resto como texto; ignora colunas sem nome; replica célula mesclada (PROCESSO/DESPACHO de um empenho com várias NEs) para todo o bloco
    cores = carregar_cores_planilha.executar(aba) # chave (linha, coluna) 1-based da planilha
    # para descobrir a cor --> print(cores[(3, 1)]), sendo que 3,1 é célula A3. Retorno = (0, 1, 1) = azul "ciano"

    cabecalho = aba.row_values(1) # cabeçalho cru: dá o nº de coluna REAL da planilha (bate com cores/update_cell/format)

    def coluna(nome):
        # (nome_real, nº_1based) da coluna, sem depender de maiúsculas/minúsculas nem espaços nas
        # pontas. O nº vem do cabeçalho cru porque ler_planilha descarta colunas de cabeçalho
        # vazio - a posição no DataFrame pode não ser a da planilha, e cores/format usam a da planilha.
        alvo = nome.strip().casefold()
        for i, atual in enumerate(cabecalho, start=1):
            if str(atual).strip().casefold() == alvo:
                return atual, i
        raise KeyError(f"coluna {nome!r} não encontrada no cabeçalho da aba RO: {cabecalho}")

    NOME_NE, COLUNA_NE = coluna("NE")
    NOME_PROCESSO, COLUNA_PROCESSO = coluna("PROCESSO")
    NOME_DESPACHO, COLUNA_DESPACHO = coluna("DESPACHO")
    _, COLUNA_TRAMITADO = coluna("TRAMITADO")

    VERMELHO = (1, 0, 0)
    BRANCO = (1, 1, 1)
    AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255) # mesmo amarelo usado nos outros scripts do pipeline
    CINZA = (0.8, 0.8, 0.8) # célula da NE cinza = empenho assinado, pronto pra baixar/anexar

    # agrupa por NÚMERO DE PROCESSO: um empenho pode ter várias NEs (uma por fornecedor) em linhas
    # diferentes do mesmo processo, muitas vezes numa célula PROCESSO mesclada. O processo é
    # buscado uma vez, todas as NEs assinadas dele são anexadas e só então ele é tramitado uma
    # única vez - tramitar tira o processo da fila, deixando as outras NEs sem como anexar.
    processos = {} # numero_processo -> {"processo", "despacho", "nes": [...], "linhas": set()}
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        if cores.get((linha_planilha, COLUNA_NE)) != CINZA: # só os empenhos com a NE cinza (assinada)
            continue

        numero_processo = str(dados.loc[linha, NOME_PROCESSO]).strip()
        grupo = processos.setdefault(numero_processo, {"processo": numero_processo, "despacho": "", "nes": [], "linhas": set()})
        grupo["nes"].append({"linha_planilha": linha_planilha, "NE": str(dados.loc[linha, NOME_NE]).strip()}) # = 2026NE510016
        grupo["linhas"].add(linha_planilha)
        despacho_linha = str(dados.loc[linha, NOME_DESPACHO]).strip()
        if despacho_linha and not grupo["despacho"]: # 1º despacho não vazio do grupo (deve ser o mesmo em todas as linhas do processo)
            grupo["despacho"] = despacho_linha

    lista_processo = list(processos.values()) # dict preserva a ordem de inserção -> mantém a ordem da planilha

    # ------- Baixa no Siafi só as NEs que ainda não estão em Downloads -------
    # (numa reexecução em que só o anexo falhou, os PDFs já estão lá - aí nem abre o Siafi)
    pasta_downloads = Path.home() / "Downloads"
    nes_para_baixar = [
        item for grupo in lista_processo for item in grupo["nes"]
        if not list(pasta_downloads.glob(f"*{item['NE']}*.pdf"))
    ]
    if nes_para_baixar:
        options = webdriver.ChromeOptions()
        options.debugger_address = "127.0.0.1:9222"
        navegador_siafi = webdriver.Chrome(options=options)
        for item in nes_para_baixar:
            baixar_ne.executar(navegador_siafi, item["NE"].replace("2026NE", "")) # = 510016
    else:
        print("Todas as NEs já estão em Downloads - pulando a etapa do Siafi.")

    # ------- Abre o Chrome maximizado -------
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # detach=True p/ impedir que o Selenium feche o navegador ao terminar a execução ou quando ocorrer um erro
    navegador_suap = webdriver.Chrome(options=options)  # navegador controlado pelo Selenium | o Chrome tem mais compatibilidade com os sites
    navegador_suap.maximize_window()

    # ------- Entra na tela de login do Suap -------
    navegador_suap.get("https://suap.iff.edu.br/accounts/login/?next=/")  # pode ser o caminho de um arquivo local
    # navegador_suap.get("http://suap.dev.iff.edu.br/accounts/login/?next=/")  # ambiente de homologação (dev)

    # ------- Faz o login no Suap -------
    navegador_suap.find_element(By.ID, "id_username").send_keys(credenciais_suap.usuario())
    navegador_suap.find_element(By.ID, "id_password").send_keys(credenciais_suap.senha() + Keys.ENTER)
    time.sleep(15)

    for grupo in lista_processo:

        processo = grupo["processo"]
        despacho = grupo["despacho"]

        nes = grupo["nes"]
        todas_anexadas = True
        for indice, item in enumerate(nes):
            eh_ultima = indice == len(nes) - 1
            try:
                # busca o processo de novo antes de CADA NE: o anexar_ne parte da página do
                # processo, e se a NE anterior falhou a página fica num estado quebrado
                campo_busca_rapida = WebDriverWait(navegador_suap, 30).until(EC.element_to_be_clickable((By.NAME, "q")))
                campo_busca_rapida.clear()
                campo_busca_rapida.send_keys(processo + Keys.ENTER)

                anexar_ne.executar(navegador_suap, item["NE"])
                pintar_celula_planilha.executar(aba, item["linha_planilha"], COLUNA_NE, BRANCO) # anexou: NE concluída
            except Exception as e:
                # não interrompe o script: pinta de vermelho essa NE e segue para as próximas, igual ao baixar_ob.py
                print(f"Erro no empenho {item['NE']} ao executar anexarNE: {e}")
                pintar_celula_planilha.executar(aba, item["linha_planilha"], COLUNA_NE, VERMELHO)
                todas_anexadas = False

            # tramita uma única vez, depois da última NE, e só se TODAS anexaram - tramitar tira o
            # processo da fila do usuário, deixando as NEs que faltaram sem como anexar numa próxima
            # rodada. Fica na página que o anexar_ne deixou (que tem o link "Encaminhar").
            if eh_ultima and todas_anexadas and despacho:
                try:
                    encaminhar_processo.executar(navegador_suap, despacho)
                    for linha_planilha in sorted(grupo["linhas"]): # todas as linhas do processo (um empenho mesclado ocupa várias)
                        aba.update_cell(linha_planilha, COLUNA_TRAMITADO, "OK") # só marca se a tramitação realmente aconteceu
                        pintar_celula_planilha.executar(aba, linha_planilha, COLUNA_PROCESSO, BRANCO) # célula do processo (ciano) -> branco: processo concluído
                except Exception as e:
                    print(f"Erro ao tramitar o processo {processo}: {e}")
            elif eh_ultima and todas_anexadas: # anexou tudo mas não há despacho pra tramitar: sinaliza o DESPACHO de amarelo pra preencher manualmente
                for linha_planilha in sorted(grupo["linhas"]):
                    pintar_celula_planilha.executar(aba, linha_planilha, COLUNA_DESPACHO, AMARELO_CLARO_1)

            navegador_suap.get("https://suap.iff.edu.br/")  # volta para a tela de início, onde tem o campo Busca rápida
            # navegador_suap.get("http://suap.dev.iff.edu.br/")  # ambiente de homologação (dev)
            time.sleep(10)

if __name__ == "__main__":
    main()
