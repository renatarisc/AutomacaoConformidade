from google.oauth2.service_account import Credentials
from selenium import webdriver
from pypdf import PdfReader
import gspread
import re
import io
import requests

import escolher_planilha

# mesmo PDF de "Andamento do processo" do preencher_planilha_ro.py, mas aqui procurando telas
# SIAFI CONSULTA-CONNS (Nota de Lançamento de Sistema) em vez de CONSULTA-CONRO - a cadeia de
# documentos é NE -> NS -> NP (Nota de Pagamento), diferente da RO -> NC/NE do outro script
RE_PROCESSO = re.compile(r"\d{5}\.\d{6}\.\d{4}-\d{2}")
RE_DESPACHO_SEM_OCORRENCIA = re.compile(r"Despacho:\s*Sem\s+ocorr[êe]ncia", re.IGNORECASE)
RE_NUMERO_NS = re.compile(r"NUMERO\s*:\s*(2026NS\d+)")
RE_TITULO_NP = re.compile(r"TITULO DE CREDITO\s*:\s*(2026NP\d+)")
# a tela CONSULTA-CONNS separa o CNPJ do nome com um traço ("10934273/0001-67 - IMA TELECOM
# LTDA"), diferente da tela CONSULTA-CONRO usada pelo RO - por isso o "-?" opcional aqui
RE_FAVORECIDO = re.compile(r"FAVORECIDO\s*:\s*[\d/\-]+\s*-?\s*(.+)")

AZUL_DESTAQUE = {"red": 0.10, "green": 0.45, "blue": 0.91} # cor aplicada só no(s) valor(es) recém-acrescentado(s) numa célula que já existia, pra facilitar identificar o que mudou nessa passada

def numero_coluna_para_letra(n): # n: nº da coluna (A=1) -> letra correspondente
    letra = ""
    while n:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra

def numero_sem_zeros(codigo):
    # "2026NS001121" -> "1121" / "2026NP000296" -> "296"
    return str(int(re.search(r"\d+$", codigo).group()))

def baixar_pdf_da_aba(navegador, url):
    # a aba abre o PDF puro (visualizador nativo do Chrome, não uma página HTML) - baixa os bytes originais
    # direto via HTTP, reaproveitando os cookies da sessão já logada no Suap, sem gravar nada em disco
    sessao = requests.Session()
    for cookie in navegador.get_cookies():
        sessao.cookies.set(cookie["name"], cookie["value"], domain=cookie["domain"])

    resposta = sessao.get(url)
    resposta.raise_for_status()
    return io.BytesIO(resposta.content)

def extrair_dados(arquivo_pdf):
    # diferente do RO, aqui não existe "ponto de partida" (despacho sem ocorrência) na primeira vez que
    # um processo de pagamento aparece - ele só existe a partir da 2ª rodada de conferência em diante.
    # Por isso NS/NP/favorecido são varridos no PDF inteiro (a mesma NS/NP pode aparecer repetida em
    # mais de um bloco da tela, daí a deduplicação), enquanto a página inicial só é calculada quando
    # existe pelo menos um despacho "Sem Ocorrência" no PDF - main() decide o que fazer com cada caso
    leitor = PdfReader(arquivo_pdf)
    paginas = [pagina.extract_text() or "" for pagina in leitor.pages]

    if not paginas or "Processo Eletrônico" not in paginas[0]:
        return None # não é um PDF de andamento de processo

    match_processo = RE_PROCESSO.search(paginas[0]) # o número do processo vem sempre na 1ª página
    if not match_processo:
        return None
    processo = match_processo.group()

    ns_encontrados = [] # números de NS distintos, na ordem em que apareceram no PDF
    np_encontrados = [] # números de NP distintos, na ordem em que apareceram no PDF
    favorecido = ""
    pagina_despacho = None

    for i, texto in enumerate(paginas):
        if RE_DESPACHO_SEM_OCORRENCIA.search(texto):
            pagina_despacho = i # guarda a ÚLTIMA ocorrência (sobrescreve a cada match, na ordem do PDF)

        for match_ns in RE_NUMERO_NS.finditer(texto):
            numero_ns = numero_sem_zeros(match_ns.group(1))
            if numero_ns not in ns_encontrados:
                ns_encontrados.append(numero_ns)

        for match_np in RE_TITULO_NP.finditer(texto):
            numero_np = numero_sem_zeros(match_np.group(1))
            if numero_np not in np_encontrados:
                np_encontrados.append(numero_np)

        if not favorecido: # o favorecido é o mesmo em todas as telas do processo - só precisa pegar uma vez
            match_favorecido = RE_FAVORECIDO.search(texto)
            if match_favorecido:
                favorecido = match_favorecido.group(1).strip().split()[0] # só o início do nome (ex: "IMA TELECOM LTDA" -> "IMA")

    pagina_inicial = (pagina_despacho + 2) if pagina_despacho is not None else None # +1 pra pular a página do próprio despacho, +1 porque a lista é 0-based

    return {
        "processo": processo,
        "favorecido": favorecido,
        "ns_encontrados": ns_encontrados,
        "np_encontrados": np_encontrados,
        "pagina_inicial": pagina_inicial,
    }

def mesclar_valores(texto_existente, encontrados):
    # texto_existente: conteúdo atual da célula (ex: "300" ou "" se a célula estiver vazia) - encontrados:
    # valores extraídos do PDF nesta rodada, na ordem em que apareceram. Nunca substitui - só acrescenta
    # o que ainda não está na célula (ex: existente "300" + encontrado "500" -> "300,500")
    existentes = [v for v in texto_existente.split(",") if v]
    novos = [v for v in encontrados if v not in existentes]
    return existentes + novos, novos

def requisicao_texto_com_destaque(sheet_id, numero_linha, indice_coluna, texto_existente, valores_mesclados):
    # monta o texto final da célula com os valores recém-acrescentados em azul e o restante (o que já
    # existia) sem cor - só dá pra fazer isso com uma chamada "crua" da API do Sheets (updateCells +
    # textFormatRuns), o gspread não expõe formatação de parte do texto dentro da mesma célula
    texto_final = ",".join(valores_mesclados)
    indice_azul = len(texto_existente) + (1 if texto_existente else 0) # pula a vírgula separadora, se houver algo antes
    return {
        "updateCells": {
            "rows": [{
                "values": [{
                    "userEnteredValue": {"stringValue": texto_final},
                    "textFormatRuns": [
                        {"startIndex": 0},
                        {"startIndex": indice_azul, "format": {"foregroundColor": AZUL_DESTAQUE}},
                    ],
                }]
            }],
            "fields": "userEnteredValue,textFormatRuns",
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": numero_linha - 1,
                "endRowIndex": numero_linha,
                "startColumnIndex": indice_coluna,
                "endColumnIndex": indice_coluna + 1,
            },
        }
    }

def requisicao_valor_simples(sheet_id, numero_linha, indice_coluna, valor):
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"numberValue": valor}}]}],
            "fields": "userEnteredValue",
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": numero_linha - 1,
                "endRowIndex": numero_linha,
                "startColumnIndex": indice_coluna,
                "endColumnIndex": indice_coluna + 1,
            },
        }
    }

def main(nome_planilha=None):
    # nome_planilha: passado pelo gui.py com a planilha escolhida na interface; rodando o
    # script sozinho (sem gui.py), usa escolher_planilha.NOME_PLANILHA_PADRAO
    #
    # Para controlar um Chrome já aberto, precisa iniciar o Chrome em modo de depuração remota (remote debugging)
    # e mandar o Selenium se conectar a ele:
    # 1- Fecha todos os Chromes abertos
    # 2- No CMD: "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeSelenium"
    # 3- Faz login no Suap, recebe os processos na Caixa de processos e abre, em abas separadas, a tela
    #    "Andamento do processo" (URL djtools/process_progress2) de cada um que for conferir hoje
    # 4- SÓ DEPOIS de abrir todas as abas, conecta o Selenium no navegador logado no Suap

    # ------- Acessa a Planilha de Controle da Conformidade (mensal) -------
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES)
    gc = gspread.authorize(credenciais)
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    aba = planilha.worksheet("TesteNS")

    cabecalho = aba.row_values(1)
    ultima_coluna = numero_coluna_para_letra(len(cabecalho))

    indice_processo = cabecalho.index("Processo")
    indice_favorecido = cabecalho.index("Favorecido")
    indice_np = cabecalho.index("NP")
    indice_ns = cabecalho.index("NS")
    indice_pagina = cabecalho.index("PÁGINA")
    # a coluna DF segue a mesma lógica de NS/NP (mesclar_valores + destaque em azul), mas ainda não é
    # preenchida - falta o usuário indicar de qual documento/tela ela é extraída

    FORMATO_LINHA = { # mesmo estilo das linhas já existentes na tabela: texto centralizado e borda preenchida em cada célula
        "horizontalAlignment": "CENTER",
        "borders": {
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"},
        },
    }

    # um processo de pagamento pode ir e voltar várias vezes pra conferência (diferente do processo de
    # empenho do RO, que é reaproveitado o ano inteiro) - por isso, antes de processar as abas, lê a
    # planilha inteira uma vez e monta um mapa processo -> {linha, NS atual, NP atual}, pra decidir se
    # cada processo encontrado no Chrome é uma linha nova ou uma linha já existente a ser complementada
    todas_linhas = aba.get_all_values()
    registro_por_processo = {}
    for i, linha_planilha in enumerate(todas_linhas[1:], start=2):
        processo_linha = linha_planilha[indice_processo] if indice_processo < len(linha_planilha) else ""
        if processo_linha:
            registro_por_processo[processo_linha] = {
                "numero_linha": i,
                "ns": linha_planilha[indice_ns] if indice_ns < len(linha_planilha) else "",
                "np": linha_planilha[indice_np] if indice_np < len(linha_planilha) else "",
            }

    # ------- Conecta o Selenium no Chrome já logado no Suap, com as abas de Andamento do processo abertas -------
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"
    navegador = webdriver.Chrome(options=options)

    aba_original = navegador.current_window_handle

    print("Processos adicionados/atualizados na planilha:")

    for janela in navegador.window_handles:
        navegador.switch_to.window(janela)
        url = navegador.current_url

        if "djtools/process_progress2" not in url:
            continue # só processa as abas que estão na tela de Andamento do processo

        dados = extrair_dados(baixar_pdf_da_aba(navegador, url))
        if dados is None:
            continue

        processo = dados["processo"]
        registro = registro_por_processo.get(processo)

        if registro is None:
            # processo novo - insere linha nova. Sem uma passada anterior de conferência não existe
            # despacho "Sem Ocorrência" no PDF pra servir de ponto de partida, então PÁGINA fica em branco
            linha_nova = [""] * len(cabecalho)
            linha_nova[indice_processo] = processo
            linha_nova[indice_favorecido] = dados["favorecido"]
            linha_nova[indice_ns] = ",".join(dados["ns_encontrados"])
            linha_nova[indice_np] = ",".join(dados["np_encontrados"])

            resultado = aba.append_row(linha_nova, value_input_option="USER_ENTERED")
            numero_linha = int(re.search(r"![A-Z]+(\d+)", resultado["updates"]["updatedRange"]).group(1))
            aba.format(f"A{numero_linha}:{ultima_coluna}{numero_linha}", FORMATO_LINHA)

            registro_por_processo[processo] = {
                "numero_linha": numero_linha,
                "ns": linha_nova[indice_ns],
                "np": linha_nova[indice_np],
            }
        else:
            # processo já existente - só acrescenta o que ainda não está lá (nunca substitui), e destaca
            # em azul só a parte recém-acrescentada de cada célula
            numero_linha = registro["numero_linha"]
            ns_mesclado, ns_novos = mesclar_valores(registro["ns"], dados["ns_encontrados"])
            np_mesclado, np_novos = mesclar_valores(registro["np"], dados["np_encontrados"])

            requisicoes = []
            if ns_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_ns, registro["ns"], ns_mesclado))
            if np_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_np, registro["np"], np_mesclado))
            if dados["pagina_inicial"] is not None:
                requisicoes.append(requisicao_valor_simples(aba.id, numero_linha, indice_pagina, dados["pagina_inicial"]))

            if requisicoes:
                aba.spreadsheet.batch_update({"requests": requisicoes})

            registro["ns"] = ",".join(ns_mesclado)
            registro["np"] = ",".join(np_mesclado)

        print(processo)

    navegador.switch_to.window(aba_original)

if __name__ == "__main__":
    main()
