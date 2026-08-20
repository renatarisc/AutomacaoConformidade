from google.oauth2.service_account import Credentials
from selenium import webdriver
from pypdf import PdfReader # pip install pypdf - lê o texto do PDF de andamento do processo (djtools/process_progress2 do Suap)
import gspread # para manipular as planilhas do Drive
import re
import io
import requests

import escolher_planilha

RE_PROCESSO = re.compile(r"\d{5}\.\d{6}\.\d{4}-\d{2}")
RE_DESPACHO_SEM_OCORRENCIA = re.compile(r"Despacho:\s*Sem\s+ocorr[êe]ncia", re.IGNORECASE)
RE_NUMERO_RO = re.compile(r"NUMERO\s*:\s*(2026RO\d+)")
RE_DOCUMENTO_NC = re.compile(r"DOCUMENTO WEB\s*:\s*(2026NC\d+)")
RE_DOCUMENTO_NE = re.compile(r"DOCUMENTO WEB\s*:\s*(2026NE\d+)")
RE_FAVORECIDO = re.compile(r"FAVORECIDO\s*:\s*[\d/\-]+\s+(.+)")
RE_VALOR = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}") # valor em formato brasileiro (ex: "1.208,33"), como aparece na tabela de eventos do SIAFI

def valor_brl_para_float(valor_str):
    return float(valor_str.replace(".", "").replace(",", "."))

def float_para_valor_brl(valor):
    return f"{valor:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")

def numero_coluna_para_letra(n): # n: nº da coluna (A=1) -> letra correspondente
    letra = ""
    while n:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra

def numero_sem_zeros(codigo):
    # "2026RO000528" -> "528" / "2026NC000207" -> "207"
    return str(int(re.search(r"\d+$", codigo).group()))

def juntar_com_e(valores):
    # ["541", "542"] -> "541 e 542" / ["541", "542", "543"] -> "541, 542 e 543"
    if len(valores) == 1:
        return valores[0]
    return ", ".join(valores[:-1]) + f" e {valores[-1]}"

def agrupar_por_referencia(pares):
    # pares: lista de (ro, referencia) - ex: (ro_nc, nc) ou (ro_ne, ne). Quando várias ROs se referem à mesma
    # NC/NE (ex: reforços que só se somam até completar o valor), agrupa numa linha só em vez de uma linha por
    # RO. Retorna a lista de ROs (não já combinada em texto) porque quem chama ainda precisa dela pra somar os
    # valores de cada RO - preserva a ordem em que cada NC/NE apareceu pela 1ª vez no PDF
    ros_por_referencia = {}
    ordem = []
    for ro, referencia in pares:
        if referencia not in ros_por_referencia:
            ros_por_referencia[referencia] = []
            ordem.append(referencia)
        if ro not in ros_por_referencia[referencia]: # a tela do SIAFI de uma RO se repete em 2 páginas do PDF (cabeçalho + eventos) - não duplica a mesma RO
            ros_por_referencia[referencia].append(ro)
    return [(ros_por_referencia[referencia], referencia) for referencia in ordem]

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
    # arquivo_pdf: objeto tipo arquivo (ex: BytesIO com os bytes baixados da aba) - monta os dados da
    # solicitação de empenho mais recente: como o mesmo processo é usado o ano inteiro (vários empenhos ao
    # longo do ano), a última página com despacho "Sem Ocorrência" é o que marca onde ela começa
    leitor = PdfReader(arquivo_pdf)
    paginas = [pagina.extract_text() or "" for pagina in leitor.pages]

    if not paginas or "Processo Eletrônico" not in paginas[0]:
        return None # não é um PDF de andamento de processo

    match_processo = RE_PROCESSO.search(paginas[0]) # o número do processo vem sempre na 1ª página
    processo = match_processo.group() if match_processo else ""

    pagina_despacho = None
    for i, texto in enumerate(paginas):
        if RE_DESPACHO_SEM_OCORRENCIA.search(texto):
            pagina_despacho = i # guarda a ÚLTIMA ocorrência (sobrescreve a cada match, na ordem do PDF)

    if pagina_despacho is None:
        return None # não achou nenhum despacho "Sem Ocorrência" - não dá pra saber onde a solicitação mais recente começa

    pagina_inicial = pagina_despacho + 2 # +1 pra pular a página do próprio despacho, +1 porque a lista é 0-based e a planilha quer a página em 1-based

    # dentro do mesmo ciclo pode ter mais de uma rodada de RO da NC/NC e de RO da NE/NE (ex: reforços sucessivos
    # ainda não fechados por um novo despacho "Sem Ocorrência") - por isso coleta TODAS as ocorrências, na
    # ordem em que aparecem no PDF, em vez de parar na primeira
    pares_nc = [] # lista de (ro_nc, nc)
    pares_ne = [] # lista de (ro_ne, ne)
    valores_ro = {} # ro -> valor em texto BRL (ex: "1.208,33"), pego da tabela de eventos da tela do SIAFI daquela RO
    favorecido = ""

    for texto in paginas[pagina_inicial - 1:]:
        match_ro = RE_NUMERO_RO.search(texto)
        if not match_ro:
            continue

        ro = numero_sem_zeros(match_ro.group(1))
        match_valor = RE_VALOR.search(texto)

        if re.search(r"FAVORECIDO\s*:", texto):
            # tela do SIAFI com FAVORECIDO + DOCUMENTO WEB de NE = uma rodada de RO da NE / NE
            match_ne = RE_DOCUMENTO_NE.search(texto)
            if match_ne:
                ne = match_ne.group(1) # a coluna NE guarda o código completo (ex: 2026NE500026), diferente da RO da NC/NC/RO da NE
                pares_ne.append((ro, ne))
                if not favorecido: # o favorecido é o mesmo em todas as rodadas do processo - só precisa pegar uma vez
                    match_favorecido = RE_FAVORECIDO.search(texto)
                    if match_favorecido:
                        favorecido = match_favorecido.group(1).strip().split()[0] # só o início do nome (ex: "IMA TELECOM LTDA" -> "IMA")
        else:
            # tela do SIAFI sem FAVORECIDO, com DOCUMENTO WEB de NC = uma rodada de RO da NC / NC
            match_nc = RE_DOCUMENTO_NC.search(texto)
            if match_nc:
                pares_nc.append((ro, numero_sem_zeros(match_nc.group(1))))

        if match_valor and ro not in valores_ro: # a RO se repete em 2 páginas do PDF (cabeçalho + eventos) - só a de eventos tem valor
            valores_ro[ro] = match_valor.group()

    # quando várias ROs se referem à mesma NC (ou à mesma NE) - ex: reforços que só se somam até completar o
    # valor - agrupa numa linha só, com as ROs juntas na mesma célula, em vez de uma linha por RO
    grupos_nc = agrupar_por_referencia(pares_nc) # -> lista de (lista de ROs, nc), 1 item por NC distinta
    grupos_ne = agrupar_por_referencia(pares_ne) # -> lista de (lista de ROs, ne), 1 item por NE distinta

    # gera uma linha por NC/NE distinta, pareando pela ordem em que aparecem no PDF (1ª NC com a 1ª NE, etc.);
    # se um ciclo não tiver o mesmo número de NCs/NEs distintas dos dois lados, as colunas do lado que faltar ficam em branco
    total_linhas = max(len(grupos_nc), len(grupos_ne), 1)

    linhas = []
    for i in range(total_linhas):
        ros_nc, nc = grupos_nc[i] if i < len(grupos_nc) else ([], "")
        ros_ne, ne = grupos_ne[i] if i < len(grupos_ne) else ([], "")

        siafi_ro_nc = "" # valor da NC: é o mesmo valor em qualquer uma das ROs do grupo (movimentações da mesma NC)
        for ro in ros_nc:
            if ro in valores_ro:
                siafi_ro_nc = valores_ro[ro]
                break

        # SIAFI NE: soma dos valores de cada RO agrupada nessa NE (o empenho vai sendo reforçado aos poucos)
        valores_ne = [valor_brl_para_float(valores_ro[ro]) for ro in ros_ne if ro in valores_ro]
        siafi_ne = float_para_valor_brl(sum(valores_ne)) if valores_ne else ""

        # se os dois valores foram encontrados, confere se batem (mesmo depois de somar as ROs de NE) - se não
        # bater, avisa no print, já que a NC deveria lastrear exatamente o valor empenhado
        valores_batem = True
        if siafi_ro_nc and valores_ne:
            valores_batem = abs(valor_brl_para_float(siafi_ro_nc) - sum(valores_ne)) < 0.01 # tolerância de 1 centavo (arredondamento)

        linhas.append({
            "PROCESSO": processo,
            "FAVORECIDO": favorecido,
            "RO da NC": juntar_com_e(ros_nc) if ros_nc else "",
            "NC": nc,
            "SIAFI RO da NC": siafi_ro_nc,
            "RO da NE": juntar_com_e(ros_ne) if ros_ne else "",
            "NE": ne,
            "SIAFI NE": siafi_ne,
            "PÁGINA": pagina_inicial,
            "VALORES_BATEM": valores_batem, # uso interno (não é coluna da planilha) - main() usa isso pra avisar no print
        })

    return linhas

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
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES) # nome do arq dentro da pasta do Projeto
    gc = gspread.authorize(credenciais)
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    aba = planilha.worksheet("TesteRO")

    cabecalho = aba.row_values(1) # nomes das colunas, na ordem da planilha - usado pra montar a linha nova sem depender da posição fixa
    ultima_coluna = numero_coluna_para_letra(len(cabecalho))

    FORMATO_LINHA = { # mesmo estilo das linhas já existentes na tabela: texto centralizado e borda preenchida em cada célula
        "horizontalAlignment": "CENTER",
        "borders": {
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"},
        },
    }

    # ------- Conecta o Selenium no Chrome já logado no Suap, com as abas de Andamento do processo abertas -------
    options = webdriver.ChromeOptions()
    options.debugger_address = "127.0.0.1:9222"
    navegador = webdriver.Chrome(options=options)

    aba_original = navegador.current_window_handle

    print("Processos adicionados na planilha:")

    for janela in navegador.window_handles:
        navegador.switch_to.window(janela)
        url = navegador.current_url

        if "djtools/process_progress2" not in url:
            continue # só processa as abas que estão na tela de Andamento do processo

        linhas_extraidas = extrair_dados(baixar_pdf_da_aba(navegador, url))

        if linhas_extraidas is None:
            continue

        for dados in linhas_extraidas: # normalmente 1 linha, mas pode ser mais de uma se o ciclo tiver várias rodadas de RO da NC/RO da NE
            linha = [""] * len(cabecalho)
            for coluna, valor in dados.items():
                if coluna in cabecalho:
                    linha[cabecalho.index(coluna)] = valor

            resultado = aba.append_row(linha, value_input_option="USER_ENTERED") # adiciona uma linha nova - cada rodada de empenho é uma linha
            numero_linha = int(re.search(r"![A-Z]+(\d+)", resultado["updates"]["updatedRange"]).group(1))
            aba.format(f"A{numero_linha}:{ultima_coluna}{numero_linha}", FORMATO_LINHA)

            if dados["VALORES_BATEM"]:
                print(dados["PROCESSO"])
            else:
                print(f"{dados['PROCESSO']} - O valor da NC não é igual ao valor empenhado.")

    navegador.switch_to.window(aba_original)

if __name__ == "__main__":
    main()
