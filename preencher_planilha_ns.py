from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from pypdf import PdfReader
import gspread
import re
import io
import requests

import contratos_db
import escolher_planilha
import pdf_aberto_windows

# mesmo PDF de "Andamento do processo" do preencher_planilha_ro.py, mas aqui procurando telas
# SIAFI CONSULTA-CONNS (Nota de Lançamento de Sistema) em vez de CONSULTA-CONRO - a cadeia de
# documentos é NE -> NS -> NP (Nota de Pagamento), diferente da RO -> NC/NE do outro script
RE_PROCESSO = re.compile(r"\d{5}\.\d{6}\.\d{4}-\d{2}")
# NF, EMISSÃO e COMPETÊNCIA têm que vir todas da MESMA fonte (a nota fiscal em si, a fonte
# oficial) - nunca de outro documento do processo que só cite/reproduza essa informação (ex: o
# Assunto da capa do processo, um despacho, uma carta de correção). Esses outros documentos são
# preenchidos a PARTIR da nota, então usá-los não garante que a informação está correta - essa
# conferência (nota vs. os outros documentos) é o que vai ser implementado à frente. Por isso
# primeiro localiza a página que É a nota fiscal (marcadores fortes do timbre dela, não uma
# menção qualquer a "Nota Fiscal" em outro lugar) e só extrai dali; se não achar (ou a nota foi
# anexada como imagem/scan, sem texto nenhum pra ler - já visto num caso real), os três campos
# ficam em branco de propósito, sem cair pra nenhuma fonte substituta
_MARCADORES_PAGINA_NF = (
    "Documento Auxiliar da NFS-e",
    "DANFSe",
    "NOTA FISCAL DE SERVIÇOS ELETRÔNICA",
    "NOTA FISCAL DE SERVICOS ELETRONICA",
    "NOTA FISCAL DE SERVIÇOS ELETRÓNICA",  # variante com "Ó" (agudo) em vez de "Ô" (circunflexo) -
                                            # é como o Tesseract OCR lê esse acento nesse tipo de
                                            # NF (ex: modelo "Nota Salvador", digitalizada sem
                                            # camada de texto - ver conformidade.py/_extrair_texto_ocr)
    "DANFE",
)
# fallback tolerante pro título "NOTA FISCAL DE SERVIÇO(S) ELETRÔNICA" - cada prefeitura varia o
# modelo do DANFSe: singular/plural ("SERVIÇO" x "SERVIÇOS"), sufixo "- NFS-e" e, sobretudo, os
# acentos de "SERVIÇO"/"ELETRÔNICA" que o pypdf corrompe (\S? casa o caractere trocado, o "C" de
# "SERVICO" sem acento, ou o acento intacto). Visto no 1250.pdf (modelo de Campos dos Goytacazes).
# Exige o título em início de linha pra NÃO casar a linha "Tipo do Documento: Nota Fiscal de
# Serviço Eletrônica (NFS-e)" da capa de digitalização do Suap (que não é a nota em si)
RE_MARCADOR_NF = re.compile(r"(?:^|\n)NOTA FISCAL DE SERVI\S?OS? ELETR\S?NICA", re.IGNORECASE)

def localizar_texto_nf(paginas):
    # devolve (texto das páginas que SÃO a nota fiscal, concatenado; números dessas páginas,
    # 1-based) - texto "" e lista vazia se não achar nenhuma (inclusive quando ela existe mas foi
    # anexada como imagem, sem texto nenhum pra extrair). Mais de uma página encontrada pode ser a
    # mesma nota espalhada por 2 páginas OU duas notas fiscais distintas anexadas ao processo (ex:
    # uma nota e sua substituta por carta de correção) - main() avisa nesse caso pra conferência manual
    indices = [
        i for i, texto in enumerate(paginas)
        if any(marcador in texto for marcador in _MARCADORES_PAGINA_NF) or RE_MARCADOR_NF.search(texto)
    ]
    texto = "\n".join(paginas[i] for i in indices)
    paginas_encontradas = [i + 1 for i in indices]
    return texto, paginas_encontradas

# a página 1 do PDF tem acentos corrompidos na extração do pypdf (ex: "n�mero" em vez de "número") -
# "\D*?" entre o rótulo e o número evita depender de casar o acento certo. O rótulo na própria nota
# geralmente é "Número da NFS-e" ou "Número da Nota" (varia conforme o modelo/prefeitura), não
# "Nota Fiscal" - esse é o texto que aparece no Assunto da capa do processo (outra fonte, não a nota)
RE_NUMERO_NF = re.compile(r"N[uú]mero\s+da\s+(?:NFS-e|Nota(?:\s+Fiscal)?)\D*?(\d+)", re.IGNORECASE)
# modelo Campos dos Goytacazes (GISS): não tem rótulo "Número da NFS-e" - o número aparece solto
# entre os rótulos "NFS-e Substituída" e "NFS-e" da tabelinha de cabeçalho (layout em colunas que
# o pypdf achata; \S? cobre o "í" corrompido de "Substituída")
RE_NUMERO_NF_CAMPOS = re.compile(r"NFS-e Substitu\S?da\s*\n\s*(\d+)\s*\n\s*NFS-e\b", re.IGNORECASE)

def extrair_numero_nf(texto_nf):
    # texto_nf já restrito à(s) página(s) da nota (ver localizar_texto_nf) - tenta o rótulo padrão
    # e, se não achar, o layout do modelo de Campos dos Goytacazes
    m = RE_NUMERO_NF.search(texto_nf) or RE_NUMERO_NF_CAMPOS.search(texto_nf)
    return m.group(1) if m else ""
# NF de MATERIAL (almoxarifado) = DANFE / NF-e, layout diferente da NFS-e de serviço: o nº vem
# como "Nº. 000.018.186" (com zeros à esquerda) e a emissão como "DATA DA EMISSÃO\n13/07/2026"
RE_NUMERO_DANFE = re.compile(r"N[ºo°]\.\s*(\d[\d.]+\d)")
RE_DATA_EMISSAO_DANFE = re.compile(r"DATA DA EMISS[ÃA]O\s*\n?\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
# campo "Tipo" da 1ª página: um processo de PAGAMENTO de nota fiscal é de serviço ("Pagamento de
# prestador de...") ou de material/almoxarifado ("Pagamento de nota fiscal de material") - o
# segundo grupo (RE_TIPO_NF_MATERIAL) só serve pra escolher como ler a nota (DANFE, não NFS-e)
# \s+ em todos os espaços internos: no PDF de andamento o campo "Tipo" quebra em várias linhas
# ("Pagamento de prestador\nde serviço à pessoa\njurídica – contratos"), então espaço fixo não casa
RE_TIPO_PAGAMENTO_NF = re.compile(r"Pagamento\s+de\s+(?:prestador\s+de|nota\s+fiscal)", re.IGNORECASE)
RE_TIPO_NF_MATERIAL = re.compile(r"Pagamento\s+de\s+nota\s+fiscal\s+de\s+material", re.IGNORECASE)
RE_DESPACHO_SEM_OCORRENCIA = re.compile(r"Despacho:\s*Sem\s+ocorr[êe]ncia", re.IGNORECASE)
RE_NUMERO_NS = re.compile(r"NUMERO\s*:\s*(2026NS\d+)")
RE_TITULO_NP = re.compile(r"TITULO DE CREDITO\s*:\s*(2026NP\d+)")
# tela SIAFI CONSULTA-CONDARF (Arrecadação Financeira - DARF) - mesmo rótulo "NUMERO :" da tela
# de NS, mas o dígito logo depois de "DF" (sempre "8", confirmado com o usuário) é um código
# institucional fixo, não faz parte do número sequencial - por isso numero_sem_zeros() não serve
# aqui (ele pegaria "800214" inteiro), precisa de um regex e uma função próprios (ver numero_df_sem_prefixo)
RE_NUMERO_DF = re.compile(r"NUMERO\s*:\s*(2026DF8\d+)")
# a tela CONSULTA-CONNS separa o CNPJ do nome com um traço ("10934273/0001-67 - IMA TELECOM
# LTDA"), diferente da tela CONSULTA-CONRO usada pelo RO - por isso o "-?" opcional aqui
# (o rótulo "FAVORECIDO :" é o texto literal da tela do SIAFI no PDF - não muda mesmo com a
# coluna da planilha tendo sido renomeada para "Empresa") - CNPJ e nome capturados em grupos
# separados pra cruzar com o banco de contratos (ver obter_abreviacao_empresa)
RE_EMPRESA = re.compile(r"FAVORECIDO\s*:\s*([\d/\-]+)\s*-?\s*(.+)")
# nº do contrato - Instrumento de Cobrança ("Contrato: 68/2024") ou texto livre da NF ("contrato
# nº 68/2024"). Best-effort (nem todo processo cita): serve pra desambiguar a empresa no banco
# quando ela tem mais de um contrato (o nome curto da "Planilha de controle" muda por contrato)
RE_CONTRATO = re.compile(r"[Cc]ontrato\s*:?\s*n?[ºo°]?\s*(\d{1,4}/\d{4})", re.IGNORECASE)
# data de emissão da NOTA FISCAL em si (não da DPS, do DARF ou de qualquer outro documento do
# processo que também tenha um campo "Data de Emissão" - ex: o Certificado de Conformidade) -
# como o modelo da nota fiscal varia (a confirmar com o usuário caso apareça um formato muito
# diferente do testado), o rótulo aceito é flexível ("Data de Emissão" ou "Data e Hora da
# Emissão", com ou sem "da <documento>" na frente) - só o valor em si (dd/mm/aaaa) é capturado
RE_DATA_EMISSAO = re.compile(
    r"(?:Data\s+(?:e\s+Hora\s+)?d[ae]\s+emiss\D?o(?:\s+d[ae]\s+(\S+))?"
    r"|Emiss\S?o\s+da\s+NFS-e)"  # rótulo do modelo de Campos dos Goytacazes ("Emissão da NFS-e")
    r"\D*?(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)
# mês em que o SERVIÇO foi prestado (não quando a nota foi emitida - "Competência da NFS-e" é
# outra coisa, é a data de emissão disfarçada) - só aparece no texto livre da "Descrição do
# Serviço" da nota, e nem toda nota escreve isso (o usuário confirmou: se não achar, fica em
# branco mesmo, ele preenche à mão). Duas formas já vistas: rótulo "Competência Julho/2026" e a
# frase corrida "...no mês de JULHO/2026..." (NFS-e nacional). O separador tanto pode ser "/"
# quanto " de " ("julho de 2026")
RE_COMPETENCIA = re.compile(
    r"(?:Compet\D?ncia|m\D?s\s+de)\s+"
    r"(Janeiro|Fevereiro|Mar\D?o|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro)"
    r"\s*(?:/|\s+de\s+)\s*(\d{2,4})",
    re.IGNORECASE,
)
_MESES_PT = ("", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro")
# alguns modelos de NFS-e não trazem o mês da competência escrito, só um "Período de referência:
# DD/MM/AAAA a DD/MM/AAAA" no texto do serviço (ex: link de telecom, modelo de Campos dos
# Goytacazes) - nesses, a competência é tomada como o MÊS INICIAL desse período
RE_PERIODO_REFERENCIA_NF = re.compile(
    r"Per\D?odo\s+de\s+refer\Dncia:\s*(\d{2})/(\d{2})/(\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)

def extrair_periodo_referencia_nf(texto_nf):
    # devolve o intervalo bruto ("06/07/2026 a 05/08/2026") pra exibir junto da competência
    # inferida, ou "" se a nota não traz esse campo
    m = RE_PERIODO_REFERENCIA_NF.search(texto_nf)
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)} a {m.group(4)}" if m else ""
# número(s) do(s) empenho(s) vinculado(s) à nota fiscal, tirado da página do "Instrumento de
# Cobrança" (contratos.gov.br) - fica na tabela "Empenhos:" (colunas Número/Subelemento/Valor),
# entre esse rótulo e o próximo campo da página ("Repactuação:") - pode ter mais de uma linha
# nessa tabela (mais de um empenho pra mesma nota), por isso pega todas as ocorrências, não só a 1ª
RE_BLOCO_EMPENHOS = re.compile(r"Empenhos:\s*(.*?)\nRepactua", re.DOTALL)
RE_NUMERO_NE = re.compile(r"2026NE\d+")

AZUL_DESTAQUE = {"red": 0.10, "green": 0.45, "blue": 0.91} # cor aplicada só no(s) valor(es) recém-acrescentado(s) numa célula que já existia, pra facilitar identificar o que mudou nessa passada

def numero_coluna_para_letra(n): # n: nº da coluna (A=1) -> letra correspondente
    letra = ""
    while n:
        n, resto = divmod(n - 1, 26)
        letra = chr(65 + resto) + letra
    return letra

def juntar_com_e(valores):
    # [10] -> "10" / [10, 25] -> "10 e 25" / [10, 25, 40] -> "10, 25 e 40"
    valores = [str(v) for v in valores]
    if len(valores) == 1:
        return valores[0]
    return ", ".join(valores[:-1]) + f" e {valores[-1]}"

def indice_coluna(cabecalho, nome):
    # comparação sem diferenciar maiúsculas/minúsculas nem espaços nas pontas - o texto exato do
    # cabeçalho na planilha já mudou de caixa mais de uma vez neste projeto (ex: "Processo" virou
    # "PROCESSO"), então não vale a pena depender do case exato bater
    nome_normalizado = nome.strip().casefold()
    for i, coluna in enumerate(cabecalho):
        if coluna.strip().casefold() == nome_normalizado:
            return i
    raise ValueError(f"coluna {nome!r} não encontrada no cabeçalho da planilha: {cabecalho}")

def numero_sem_zeros(codigo):
    # "2026NS001121" -> "1121" / "2026NP000296" -> "296"
    return str(int(re.search(r"\d+$", codigo).group()))

def numero_df_sem_prefixo(codigo):
    # "2026DF800214" -> "214" - remove especificamente o "8" fixo logo após "DF" (não é zero à
    # esquerda, então numero_sem_zeros não serve) e só então trata o resto como número sequencial,
    # que cresce e ganha mais dígitos com o tempo do mesmo jeito que NS/NP (ex: viraria "1000"
    # quando passar de 999)
    return str(int(re.search(r"DF8(\d+)$", codigo).group(1)))

def extrair_data_emissao_nf(texto_nf):
    # texto_nf: texto já restrito à(s) página(s) que É a nota fiscal (ver localizar_texto_nf) -
    # não tem mais filtro de marcador aqui, isso já foi resolvido por quem chama
    for rotulo, data in RE_DATA_EMISSAO.findall(texto_nf):
        if (rotulo or "").upper() != "DPS": # ignora a emissão da DPS - é um documento diferente da nota fiscal em si
            return data
    return ""

def extrair_competencia_nf(texto_nf):
    # mês/ano em que o serviço foi prestado, tirado do texto livre da "Descrição do Serviço" da
    # própria nota fiscal (texto_nf já restrito a ela - ver localizar_texto_nf) - só funciona
    # quando a nota escreve isso explicitamente (nem toda escreve); se não achar, devolve vazio de
    # propósito (o usuário completa à mão nesse caso)
    match = RE_COMPETENCIA.search(texto_nf)
    if match:
        return f"{match.group(1)}/{match.group(2)}"
    # sem mês escrito: usa o mês inicial do "Período de referência", quando a nota traz esse campo
    m_periodo = RE_PERIODO_REFERENCIA_NF.search(texto_nf)
    if m_periodo and 1 <= int(m_periodo.group(2)) <= 12:
        return f"{_MESES_PT[int(m_periodo.group(2))]}/{m_periodo.group(3)}"
    return ""

def extrair_empenhos(paginas):
    # números de empenho (NE) distintos, tirados da tabela "Empenhos:" da página do Instrumento
    # de Cobrança - na ordem em que aparecem, sem duplicar
    for texto in paginas:
        match_bloco = RE_BLOCO_EMPENHOS.search(texto)
        if not match_bloco:
            continue
        empenhos = []
        for numero in RE_NUMERO_NE.findall(match_bloco.group(1)):
            if numero not in empenhos:
                empenhos.append(numero)
        if empenhos:
            return empenhos
    return []

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
    # Por isso NS/NP/empresa são varridos no PDF inteiro (a mesma NS/NP pode aparecer repetida em
    # mais de um bloco da tela, daí a deduplicação), enquanto a página inicial só é calculada quando
    # existe pelo menos um despacho "Sem Ocorrência" no PDF - main() decide o que fazer com cada caso
    try:
        leitor = PdfReader(arquivo_pdf)
        if leitor.is_encrypted:
            # PDFs de terceiros soltos na pasta Downloads (notas, DANFE, boletos) às vezes vêm
            # com senha de dono e senha de usuário em branco - tenta abrir sem senha; se nem
            # assim abrir (ou o PDF estiver corrompido), não é algo que conseguimos ler, então
            # descarta o arquivo e segue para o próximo em vez de derrubar a rodada inteira
            leitor.decrypt("")
        if not leitor.pages:
            return None
        texto_pagina1 = leitor.pages[0].extract_text() or ""
    except Exception as erro:
        print(f"  (ignorado: não foi possível ler {getattr(arquivo_pdf, 'name', 'PDF')} - {erro})")
        return None

    # checa só a 1ª página (lida acima) antes de extrair o PDF inteiro (pode ter dezenas/centenas
    # de páginas) - descarta rápido um PDF que nem é do Suap, sem gastar tempo com o resto.
    # Importante quando o PDF vem de uma varredura de pasta (ex: Downloads) em vez de uma aba já
    # confirmada do Chrome
    if "Processo Eletrônico" not in texto_pagina1:
        return None # não é um PDF de andamento de processo

    # o campo "Tipo" da 1ª página diferencia processo de pagamento (NS - serviço OU material) de
    # solicitação de empenho (RO) - importante quando o PDF vem de uma varredura de pasta em vez de
    # uma aba já confirmada: a mesma pasta Downloads pode ter PDFs dos dois tipos misturados no
    # mesmo dia (confirmado com o usuário), e sem esse filtro cada script perderia tempo extraindo
    # o PDF inteiro do outro tipo
    if not RE_TIPO_PAGAMENTO_NF.search(texto_pagina1):
        return None # não é um processo de pagamento de nota fiscal - provavelmente é do tipo RO
    eh_material = bool(RE_TIPO_NF_MATERIAL.search(texto_pagina1)) # muda só como a nota é lida (DANFE)

    match_processo = RE_PROCESSO.search(texto_pagina1) # o número do processo vem sempre na 1ª página
    if not match_processo:
        return None
    processo = match_processo.group()

    paginas = [texto_pagina1] + [pagina.extract_text() or "" for pagina in leitor.pages[1:]]

    # NF, emissão e competência vêm todas da mesma página (a nota fiscal em si) - se ela não for
    # encontrada (ou tiver sido anexada como imagem/scan, sem texto pra ler), as três ficam em
    # branco juntas, nunca uma vindo de um documento e outra de outro
    texto_nf, paginas_nf = localizar_texto_nf(paginas)
    if not texto_nf:
        nf, emissao, competencia = "", "", ""
    elif eh_material:
        # DANFE (NF-e de material): nº com zeros à esquerda ("000.018.186" -> "18186"); competência
        # é o mês do SERVIÇO prestado - não existe em compra de material, fica em branco
        match_nf = RE_NUMERO_DANFE.search(texto_nf)
        nf = str(int(re.sub(r"\D", "", match_nf.group(1)))) if match_nf else ""
        match_emissao = RE_DATA_EMISSAO_DANFE.search(texto_nf)
        emissao = match_emissao.group(1) if match_emissao else ""
        competencia = ""
    else:
        nf = extrair_numero_nf(texto_nf)
        emissao = extrair_data_emissao_nf(texto_nf)
        competencia = extrair_competencia_nf(texto_nf)
    empenhos_encontrados = extrair_empenhos(paginas) # número(s) do(s) empenho(s), tirado(s) do Instrumento de Cobrança
    if not empenhos_encontrados:
        # material às vezes não tem Instrumento de Cobrança - a NE vem no Assunto da 1ª página
        # ("... Nota de Empenho: 2026NE500836 ...")
        empenhos_encontrados = list(dict.fromkeys(RE_NUMERO_NE.findall(texto_pagina1)))

    numero_contrato = "" # desambigua a empresa no banco quando ela tem mais de um contrato
    for texto in paginas:
        match_contrato = RE_CONTRATO.search(texto)
        if match_contrato:
            numero_contrato = match_contrato.group(1)
            break

    ns_encontrados = [] # números de NS distintos, na ordem em que apareceram no PDF
    np_encontrados = [] # números de NP distintos, na ordem em que apareceram no PDF
    df_encontrados = [] # números de DF distintos, na ordem em que apareceram no PDF - um NP pode ter vários DFs (ex: um por retenção/imposto)
    empresa = ""
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

        for match_df in RE_NUMERO_DF.finditer(texto):
            numero_df = numero_df_sem_prefixo(match_df.group(1))
            if numero_df not in df_encontrados:
                df_encontrados.append(numero_df)

        if not empresa: # a empresa é a mesma em todas as telas do processo - só precisa pegar uma vez
            match_empresa = RE_EMPRESA.search(texto)
            if match_empresa:
                cnpj_empresa = match_empresa.group(1)
                nome_completo_empresa = match_empresa.group(2).strip()
                # cruza com o banco de contratos pra achar a abreviação já cadastrada em "Planilha
                # de controle" (ex: "A M GAMBA ALIMENTOS" -> "GAMBA"); passa o nº do contrato pra
                # não pegar o nome curto de outro contrato da mesma empresa. Contrato não cadastrado
                # / empresa com vários contratos e sem nº -> mantém o nome completo em vez de abreviar
                abreviacao = contratos_db.obter_abreviacao_empresa(
                    cnpj_empresa, nome_completo_empresa, numero_contrato
                )
                empresa = abreviacao or nome_completo_empresa

    pagina_inicial = (pagina_despacho + 2) if pagina_despacho is not None else None # +1 pra pular a página do próprio despacho, +1 porque a lista é 0-based

    return {
        "processo": processo,
        "empresa": empresa,
        "nf": nf,
        "emissao": emissao,
        "competencia": competencia,
        "paginas_nf": paginas_nf, # em que página(s) do PDF a nota fiscal foi localizada (1-based) - main() usa isso pra avisar se achou mais de uma
        "ns_encontrados": ns_encontrados,
        "np_encontrados": np_encontrados,
        "df_encontrados": df_encontrados,
        "empenhos_encontrados": empenhos_encontrados,
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

def requisicao_texto_simples(sheet_id, numero_linha, indice_coluna, texto):
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": texto}}]}],
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

    contratos_db.inicializar_db() # garante que a tabela de contratos existe, mesmo rodando esse script sem nunca ter aberto a tela de Cadastrar Contrato antes

    # ------- Acessa a Planilha de Controle da Conformidade (mensal) -------
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES)
    gc = gspread.authorize(credenciais)
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    aba = planilha.worksheet("NS")

    cabecalho = aba.row_values(1)
    ultima_coluna = numero_coluna_para_letra(len(cabecalho))

    indice_processo = indice_coluna(cabecalho, "Processo")
    indice_empresa = indice_coluna(cabecalho, "EMPRESA") # renomeada de "Favorecido" na planilha
    indice_nf = indice_coluna(cabecalho, "NF")
    indice_np = indice_coluna(cabecalho, "NP")
    indice_ns = indice_coluna(cabecalho, "NS")
    indice_pagina = indice_coluna(cabecalho, "PÁGINA")
    indice_df = indice_coluna(cabecalho, "DF")
    indice_emissao = indice_coluna(cabecalho, "EMISSÃO")
    indice_competencia = indice_coluna(cabecalho, "COMPETÊNCIA")
    indice_empenho = indice_coluna(cabecalho, "EMPENHO")

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
                "nf": linha_planilha[indice_nf] if indice_nf < len(linha_planilha) else "",
                "ns": linha_planilha[indice_ns] if indice_ns < len(linha_planilha) else "",
                "np": linha_planilha[indice_np] if indice_np < len(linha_planilha) else "",
                "df": linha_planilha[indice_df] if indice_df < len(linha_planilha) else "",
                "emissao": linha_planilha[indice_emissao] if indice_emissao < len(linha_planilha) else "",
                "competencia": linha_planilha[indice_competencia] if indice_competencia < len(linha_planilha) else "",
                "empenho": linha_planilha[indice_empenho] if indice_empenho < len(linha_planilha) else "",
            }

    def processar_processo(dados):
        # insere (processo novo) ou atualiza (processo já existente) a linha desse processo na
        # planilha, a partir dos dados já extraídos do PDF - reaproveitado tanto pelas abas do
        # Chrome quanto pelos PDFs abertos localmente no computador (mesmo formato de "dados",
        # a origem do PDF não importa daqui pra frente)
        processo = dados["processo"]
        registro = registro_por_processo.get(processo)

        if registro is None:
            # processo novo - insere linha nova. Sem uma passada anterior de conferência não existe
            # despacho "Sem Ocorrência" no PDF pra servir de ponto de partida, então PÁGINA fica em branco
            linha_nova = [""] * len(cabecalho)
            linha_nova[indice_processo] = processo
            linha_nova[indice_empresa] = dados["empresa"]
            linha_nova[indice_nf] = dados["nf"]
            linha_nova[indice_emissao] = dados["emissao"]
            linha_nova[indice_competencia] = dados["competencia"]
            linha_nova[indice_ns] = ",".join(dados["ns_encontrados"])
            linha_nova[indice_np] = ",".join(dados["np_encontrados"])
            linha_nova[indice_empenho] = ",".join(dados["empenhos_encontrados"])
            # "---" quando não achou nenhum DF nesse PDF - mesmo símbolo de "não aplicável" usado
            # em outras colunas (ex: Valor ISS); é uma situação legítima, não erro: processo cujo
            # serviço não tem incidência de imposto simplesmente não gera DF nenhum
            linha_nova[indice_df] = ",".join(dados["df_encontrados"]) if dados["df_encontrados"] else "---"

            resultado = aba.append_row(linha_nova, value_input_option="USER_ENTERED")
            numero_linha = int(re.search(r"![A-Z]+(\d+)", resultado["updates"]["updatedRange"]).group(1))
            aba.format(f"A{numero_linha}:{ultima_coluna}{numero_linha}", FORMATO_LINHA)

            registro_por_processo[processo] = {
                "numero_linha": numero_linha,
                "nf": linha_nova[indice_nf],
                "ns": linha_nova[indice_ns],
                "np": linha_nova[indice_np],
                "df": linha_nova[indice_df],
                "emissao": linha_nova[indice_emissao],
                "competencia": linha_nova[indice_competencia],
                "empenho": linha_nova[indice_empenho],
            }
        else:
            # processo já existente - só acrescenta o que ainda não está lá (nunca substitui), e destaca
            # em azul só a parte recém-acrescentada de cada célula
            numero_linha = registro["numero_linha"]
            ns_mesclado, ns_novos = mesclar_valores(registro["ns"], dados["ns_encontrados"])
            np_mesclado, np_novos = mesclar_valores(registro["np"], dados["np_encontrados"])
            empenho_mesclado, empenho_novos = mesclar_valores(registro["empenho"], dados["empenhos_encontrados"])
            # "---" é o placeholder de "nenhum DF encontrado ainda" - trata como célula vazia na
            # mesclagem, senão viraria literalmente "---,214" na primeira vez que um DF real aparece
            df_existente = "" if registro["df"] == "---" else registro["df"]
            df_mesclado, df_novos = mesclar_valores(df_existente, dados["df_encontrados"])

            requisicoes = []
            if ns_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_ns, registro["ns"], ns_mesclado))
            if np_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_np, registro["np"], np_mesclado))
            if empenho_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_empenho, registro["empenho"], empenho_mesclado))
            if dados["pagina_inicial"] is not None:
                requisicoes.append(requisicao_valor_simples(aba.id, numero_linha, indice_pagina, dados["pagina_inicial"]))
            if df_novos:
                requisicoes.append(requisicao_texto_com_destaque(aba.id, numero_linha, indice_df, df_existente, df_mesclado))
            elif not registro["df"]: # nunca teve DF e esse PDF também não trouxe nenhum - marca "---" (não aplicável)
                requisicoes.append(requisicao_texto_simples(aba.id, numero_linha, indice_df, "---"))
                registro["df"] = "---"
            if not registro["nf"] and dados["nf"]: # NF é único por processo (1 nota fiscal = 1 processo) - só preenche se ainda estiver vazio, nunca mescla
                requisicoes.append(requisicao_texto_simples(aba.id, numero_linha, indice_nf, dados["nf"]))
                registro["nf"] = dados["nf"]
            if not registro["emissao"] and dados["emissao"]: # mesma lógica da NF: único por processo, só preenche se ainda estiver vazio
                requisicoes.append(requisicao_texto_simples(aba.id, numero_linha, indice_emissao, dados["emissao"]))
                registro["emissao"] = dados["emissao"]
            if not registro["competencia"] and dados["competencia"]: # mesma lógica: único por processo, só preenche se ainda estiver vazio (fica em branco se a nota não informar)
                requisicoes.append(requisicao_texto_simples(aba.id, numero_linha, indice_competencia, dados["competencia"]))
                registro["competencia"] = dados["competencia"]

            if requisicoes:
                aba.spreadsheet.batch_update({"requests": requisicoes})

            registro["ns"] = ",".join(ns_mesclado)
            registro["np"] = ",".join(np_mesclado)
            if empenho_novos:
                registro["empenho"] = ",".join(empenho_mesclado)
            if df_novos:
                registro["df"] = ",".join(df_mesclado)

        # avisa no console quando a nota fiscal não pôde ser localizada/lida (NF, emissão e
        # competência ficaram em branco) ou quando mais de uma página parece ser uma nota fiscal
        # (pode ser a mesma nota em 2 páginas ou duas notas distintas - qualquer um dos casos
        # merece conferência manual, por isso avisa em vez de simplesmente escolher uma)
        if len(dados["paginas_nf"]) > 1:
            print(f"{processo} - Identificada mais de uma NF nas páginas {juntar_com_e(dados['paginas_nf'])}")
        elif not dados["nf"]:
            print(f"{processo} - NF não identificada")
        else:
            print(processo)

    # ------- Conecta o Selenium no Chrome já logado no Suap, com as abas de Andamento do processo
    # abertas - OPCIONAL: se não achar Chrome em modo debug (usuário optou por só usar PDFs abertos
    # localmente, ou simplesmente esqueceu de abrir o Chrome assim), não trava o script - só avisa
    # e segue direto pros PDFs locais logo abaixo. Esse aviso vem ANTES do cabeçalho "Processos
    # adicionados", pra não intercalar com a lista de processos que vem logo em seguida -------
    try:
        options = webdriver.ChromeOptions()
        options.debugger_address = "127.0.0.1:9222"
        navegador = webdriver.Chrome(options=options)
    except WebDriverException:
        print("Abas do Chrome com PDFs abertos não encontradas, processando só os PDFs baixados.")
        navegador = None

    print("Processos adicionados/atualizados na planilha:")

    if navegador is not None:
        aba_original = navegador.current_window_handle

        for janela in navegador.window_handles:
            navegador.switch_to.window(janela)
            url = navegador.current_url

            if "djtools/process_progress2" not in url:
                continue # só processa as abas que estão na tela de Andamento do processo

            dados = extrair_dados(baixar_pdf_da_aba(navegador, url))
            if dados is None:
                continue

            processar_processo(dados)

        navegador.switch_to.window(aba_original)

    # ------- Além das abas do Chrome, também processa PDFs baixados/abertos localmente no
    # computador - combina duas fontes: janelas/abas em primeiro plano de qualquer visualizador
    # (listar_pdfs_abertos) e PDFs baixados recentemente na pasta Downloads (listar_pdfs_recentes,
    # que existe porque visualizadores com abas dentro de uma única janela, como o Foxit, escondem
    # da detecção por janela qualquer aba que não esteja em primeiro plano). Mesmo processo que já
    # apareceu numa aba do Chrome acima simplesmente atualiza o registro já criado (
    # registro_por_processo é compartilhado entre os laços), em vez de duplicar a linha.
    caminhos_pdf = list(dict.fromkeys(
        pdf_aberto_windows.listar_pdfs_abertos() + pdf_aberto_windows.listar_pdfs_recentes()
    ))
    for caminho_pdf in caminhos_pdf:
        with open(caminho_pdf, "rb") as arquivo:
            dados = extrair_dados(arquivo)
        if dados is None:
            continue # não é um PDF de Andamento do processo (ex: outro PDF qualquer aberto no Windows)

        processar_processo(dados)

if __name__ == "__main__":
    main()
