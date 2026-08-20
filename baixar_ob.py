from google.oauth2.service_account import Credentials
from pathlib import Path
from pypdf import PdfReader # pip install pypdf - lê o texto do PDF baixado para conferir a OB
import pyautogui
import gspread # para manipular as planilhas do Drive
import pandas as pd
import time

import carregar_cores_planilha
import escolher_planilha
import pintar_celula_planilha

PASTA_DOWNLOADS = Path.home() / "Downloads" # pasta onde o Sistema salva o PDF da OB (nome sequencial: 00000001.pdf, 00000002.pdf...)

AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255) # mesmo amarelo usado no relacionar_valor_op.py e no relacionar_op_ob.py
VERMELHO = (1, 0, 0)

def cor_bate(cor_celula, cor_alvo, tolerancia=0.01):
    # a API do Sheets guarda a cor com menos precisão do que o float do Python,
    # então uma comparação exata (==) quase nunca bate mesmo com a cor visualmente igual
    if not cor_celula:
        return False
    return all(abs(c - alvo) < tolerancia for c, alvo in zip(cor_celula, cor_alvo))

def esperar_pdfs(arquivos_antes, quantidade_esperada, tentativas=30, intervalo=1):
    # espera (só uma vez, depois de todos os downloads) até a pasta Downloads ter a quantidade
    # de PDFs novos esperada, dando tempo do Sistema terminar de salvar o último arquivo
    for _ in range(tentativas):
        novos = set(PASTA_DOWNLOADS.glob("*.pdf")) - arquivos_antes
        if len(novos) >= quantidade_esperada:
            return novos
        time.sleep(intervalo)
    return set(PASTA_DOWNLOADS.glob("*.pdf")) - arquivos_antes # timeout: devolve o que tiver, mesmo incompleto

def variante_barra_processo(processo):
    # a planilha guarda o processo com ponto (23322.000745.2026-89), mas o PDF do Siafi
    # às vezes mostra com barra antes do ano (23322.000745/2026-89) - aceita as duas formas
    if "." not in processo:
        return processo
    base, ano = processo.rsplit(".", 1)
    return f"{base}/{ano}"

def pdf_contem(caminho_pdf, ob, valor, processo):
    leitor = PdfReader(str(caminho_pdf))
    conteudo = "\n".join(pagina.extract_text() or "" for pagina in leitor.pages)
    processo_ok = processo in conteudo or variante_barra_processo(processo) in conteudo
    return ob in conteudo and valor in conteudo and processo_ok

def main(nome_planilha=None):
    # nome_planilha: passado pelo gui.py com a planilha escolhida na interface; rodando o
    # script sozinho (sem gui.py), usa escolher_planilha.NOME_PLANILHA_PADRAO
    pyautogui.PAUSE = 0.5 # Pausa entre os comandos

    time.sleep(5)

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
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    # aba = planilha.worksheet("RO")
    aba = planilha.worksheet("TesteNS")

    dados = pd.DataFrame(aba.get_all_records(numericise_ignore=['all'])) # get_all_records() usa a 1ª linha como cabeçalho e exige que cada coluna tenha nome único
    cores = carregar_cores_planilha.executar(aba) # chama a def

    # a cor amarelo claro 1 sinaliza que a OB (ISS ou PG) já foi relacionada pelo relacionar_op_ob.py e ainda precisa
    # ser baixada - coluna calculada pelo nome do cabeçalho, não fixa, pra não quebrar se a coluna mudar de lugar
    COLUNAS_OB = [
        {"coluna": dados.columns.get_loc("OB ISS") + 1, "nome": "OB ISS", "nome_valor": "Valor ISS"},
        {"coluna": dados.columns.get_loc("OB PG") + 1, "nome": "OB PG", "nome_valor": "Valor PG"},
    ]

    lista_processo = []
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        for info in COLUNAS_OB:
            if not cor_bate(cores.get((linha_planilha, info["coluna"])), AMARELO_CLARO_1):
                continue
            lista_processo.append({
                "linha_planilha": linha_planilha,
                "coluna": info["coluna"],
                "OB": dados.loc[linha, info["nome"]],
                "valor": str(dados.loc[linha, info["nome_valor"]]).strip(),
                "processo": str(dados.loc[linha, "Processo"]).strip(),
            })

    lista_sucesso = [] # processos cujo pyautogui não deu erro, na ordem em que rodaram
    arquivos_antes = set(PASTA_DOWNLOADS.glob("*.pdf")) # snapshot único, antes de começar os downloads

    for processo in lista_processo:

        OB = processo["OB"]
        numero_OB = OB.replace("2026OB", "")

        try:
            pyautogui.click(293, 596) # linha de comando
            pyautogui.write(">conob")
            pyautogui.press("enter")

            pyautogui.click(553, 290) # número da OB
            pyautogui.write(numero_OB)
            pyautogui.press("enter")

            pyautogui.press("f2") # detalhar a OB que está com o cursor

            pyautogui.click(842, 64) # botão captura de tela
            pyautogui.click(1166, 524) # na tela p/ ENTER
            pyautogui.press("enter") # mudar de tela

            pyautogui.click(842, 64) # botão captura de tela
            pyautogui.click(1166, 524) # na tela p/ ENTER
            pyautogui.press("enter") # mudar de tela

            pyautogui.click(842, 64) # botão captura de tela
            pyautogui.click(1166, 524) # na tela p/ ENTER
            pyautogui.press("enter") # mudar de tela

            pyautogui.click(874, 67) # botão carrega captura de tela

            pyautogui.moveTo(346, 533, duration=0.5) # botão "Selecionar Todos'
            time.sleep(1)
            pyautogui.click()
            pyautogui.click()

            pyautogui.moveTo(618, 533, duration=0.5) # botão "Imprimir e Excluir Selecionados"
            time.sleep(1)
            pyautogui.click()

            pyautogui.moveTo(608, 579, duration=0.5) # botão "OK"
            time.sleep(1)
            pyautogui.click()

            lista_sucesso.append(processo)

            pyautogui.press("f3") # sai da tela da OB (PF3=SAI)
            pyautogui.click(293, 596) # linha de comando

        except Exception as e:
            # o pyautogui falhou nessa OB; pinta de vermelho para indicar que precisa ser baixada manualmente
            print(f"Erro ao baixar a OB {OB}: {e}")
            pintar_celula_planilha.executar(aba, processo["linha_planilha"], processo["coluna"], VERMELHO)

    # ------- Só depois de rodar todos os downloads é que confere e renomeia os PDFs, pra não atrasar o pyautogui -------
    # pareia cada PDF novo com a OB cujo conteúdo bate (em vez de assumir a ordem), porque o pyautogui não
    # percebe quando um clique cai na tela errada e não gera PDF nenhum ou um pdf errado pra a OB
    arquivos_novos = esperar_pdfs(arquivos_antes, len(lista_sucesso))

    processos_pendentes = list(lista_sucesso)
    for arquivo in arquivos_novos:
        encontrado = next(
            (p for p in processos_pendentes if pdf_contem(arquivo, p["OB"], p["valor"], p["processo"])),
            None
        )
        if encontrado:
            arquivo.rename(arquivo.with_name(f"{encontrado['OB']}.pdf"))
            processos_pendentes.remove(encontrado)
        # PDF que não bate com nenhuma OB pendente é ignorado (pode ser de outra coisa, salvo na pasta por outro motivo)

    for processo in processos_pendentes:
        print(f"PDF da OB {processo['OB']} não encontrado/conferido na pasta Downloads")
        pintar_celula_planilha.executar(aba, processo["linha_planilha"], processo["coluna"], VERMELHO)

if __name__ == "__main__":
    main()
