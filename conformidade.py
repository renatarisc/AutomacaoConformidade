import calendar
import re

import webview  # mesma lib do gui.py - abre a janela de resultado em cima da instância já em execução
from pypdf import PdfReader
from selenium import webdriver
from selenium.common.exceptions import WebDriverException

import contratos_db
import pdf_aberto_windows
import preencher_planilha_ns as ns  # reaproveita localizar_texto_nf, RE_NUMERO_NF, RE_EMPRESA,
                                     # extrair_data_emissao_nf, extrair_competencia_nf, juntar_com_e -
                                     # já testados/validados nesse módulo, não duplica aqui

# ------- utilidades de texto -------

def remover_duplicatas_consecutivas(texto):
    # as páginas de Termo/Despacho/Relatório do Suap sempre repetem cada linha de texto duas
    # vezes seguidas (provavelmente uma camada de acessibilidade sobreposta ao texto visível) -
    # sem isso, capturar um trecho de texto (ex: nome da contratada) vem duplicado
    linhas = texto.split("\n")
    limpas = []
    anterior = None
    for linha in linhas:
        chave = linha.strip()
        if chave and chave == anterior:
            continue
        limpas.append(linha)
        anterior = chave
    return "\n".join(limpas)

def limpar_espacos(texto):
    return re.sub(r"\s+", " ", texto or "").strip()

def comparar_numeros(a, b):
    # ignora zeros à esquerda em cada grupo de dígitos (ex: contrato "02/2022" == "00002/2022",
    # NF "588" == "000588") - usado só pra contrato e NF, não pra processo nem CNPJ: processo é
    # formatado sempre igual (uma diferença ali seria de verdade, não só de zeros), e CNPJ tem
    # grupos de tamanho fixo com significado próprio (o "0001" da filial não pode virar "1" -
    # ver comparar_cnpjs)
    if not a or not b:
        return False
    normalizar = lambda v: re.sub(r"\d+", lambda m: str(int(m.group())), v)
    return normalizar(a) == normalizar(b)

def comparar_cnpjs(a, b):
    # compara só os dígitos, sem remover zero nenhum (cada grupo do CNPJ tem tamanho fixo e
    # significado próprio - diferente de contrato/NF, "0001" aqui não é um "1" com zero de
    # preenchimento, é literalmente o código da filial)
    if not a or not b:
        return False
    digitos = lambda v: re.sub(r"\D", "", v)
    return digitos(a) == digitos(b)

def comparar_textos(a, b):
    # comparação "solta" pra nome/texto livre: ignora maiúscula, acento e espaços - considera
    # igual se um dos dois é o início do outro (documento pode citar só o começo do texto da
    # fonte segura, ex: objeto do contrato sem o complemento entre parênteses)
    import unicodedata
    def normalizar(v):
        sem_acento = unicodedata.normalize("NFKD", v or "").encode("ascii", "ignore").decode()
        return re.sub(r"\s+", " ", sem_acento).strip().upper().rstrip(".")
    na, nb = normalizar(a), normalizar(b)
    if not na or not nb:
        return False
    return na == nb or na.startswith(nb) or nb.startswith(na)

# ------- meses / competência / período -------

_MESES_NOME = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
               "Setembro", "Outubro", "Novembro", "Dezembro"]
_MESES_INDICE = {nome.lower(): i for i, nome in enumerate(_MESES_NOME) if nome}
_MESES_INDICE["marco"] = 3  # variante sem acento

def _mes_para_numero(nome_mes):
    return _MESES_INDICE.get((nome_mes or "").lower())

def calcular_periodo(competencia):
    # "Junho/2026" -> "01/06/2026 a 30/06/2026" - período é sempre o mês cheio da competência
    if not competencia:
        return None
    match = re.match(r"([A-Za-zçÇãÃéÉêÊúÚ]+)/(\d{4})", competencia)
    if not match:
        return None
    numero_mes = _mes_para_numero(match.group(1))
    if not numero_mes:
        return None
    ano = int(match.group(2))
    ultimo_dia = calendar.monthrange(ano, numero_mes)[1]
    return f"01/{numero_mes:02d}/{ano} a {ultimo_dia:02d}/{numero_mes:02d}/{ano}"

RE_COMPETENCIA_EXPLICITA = re.compile(
    # aceita "competência Junho/2026" e "competência Junho de 2026" (2 fraseados reais
    # confirmados - ambos com a palavra "competência" explícita, então nenhum dos dois precisa
    # de seta na exibição, mesmo o segundo precisando trocar "de" por "/")
    r"competência\s*\n?\s*([A-Za-zçÇãÃéÉêÊúÚ]+)\s*(?:de\s+|/)\s*(\d{4})", re.IGNORECASE)
RE_PAGAMENTO_MES = re.compile(
    r"Pagamento de\s+([A-Za-zçÇãÃéÉêÊúÚ]+)\s+de\s+(\d{4})", re.IGNORECASE)
# ancorado no rótulo do campo (não um "DD/AAAA" solto qualquer) - um "DD/AAAA" sem contexto
# colide fácil com outros números da página (ex: achou "68/2024" dentro de "Contrato: 00068/2024"
# antes de chegar no campo certo, quando o regex não tinha esse anchor)
RE_MES_ANO_NUMERICO = re.compile(
    r"M[êe]s\s*/\s*Ano Refer[êe]ncia:\s*\n?\s*(\d{2})\s*/\s*(\d{4})", re.IGNORECASE)

def extrair_competencia_documento(texto, aceitar_numerico=False):
    # tenta, nessa ordem: "competência <Mês>/<Ano>" (já explícito, sem seta), "Pagamento de <mês>
    # de <ano>" (precisa de seta, mesma coisa dita diferente), e opcionalmente "MM / AAAA"
    # numérico (usado só no Instrumento de Cobrança, campo "Mês / Ano Referência", sempre precisa
    # de seta pois não é "Mês/Ano" nem cita a palavra competência) - devolve (bruto_pra_exibir,
    # normalizado) ou (None, None)
    m = RE_COMPETENCIA_EXPLICITA.search(texto)
    if m:
        numero_mes = _mes_para_numero(m.group(1))
        if numero_mes:
            normalizado = f"{_MESES_NOME[numero_mes]}/{m.group(2)}"
            return normalizado, normalizado  # já explícito - documento mostra só o valor final

    m = RE_PAGAMENTO_MES.search(texto)
    if m:
        numero_mes = _mes_para_numero(m.group(1))
        if numero_mes:
            normalizado = f"{_MESES_NOME[numero_mes]}/{m.group(2)}"
            return m.group(0), normalizado

    if aceitar_numerico:
        m = RE_MES_ANO_NUMERICO.search(texto)
        if m and 1 <= int(m.group(1)) <= 12:
            normalizado = f"{_MESES_NOME[int(m.group(1))]}/{m.group(2)}"
            bruto = f"{m.group(1)} / {m.group(2)}"  # só o valor em si (não o rótulo "Mês / Ano Referência:" inteiro)
            return bruto, normalizado

    return None, None

def exibir_competencia(bruto, normalizado):
    # mostra "bruto → normalizado" só quando precisou de interpretação (ex: "Pagamento de junho
    # de 2026" ou "06 / 2026") - quando já veio explícito ("competência Junho/2026"), bruto e
    # normalizado já são o mesmo texto, então mostra só o valor final, sem seta
    if not bruto:
        return None
    return normalizado if bruto == normalizado else f"{bruto} → {normalizado}"

# ------- dados da própria nota fiscal (fonte segura) -------

RE_VALOR_SERVICO_NF = re.compile(r"Valor do Servi\D?o\s*\n?\s*R\$\s*([\d.,]+)")
RE_VALOR_LIQUIDO_NF = re.compile(r"Valor L[íi]quido da NFS-e\s*\n?\s*R\$\s*([\d.,]+)", re.IGNORECASE)
# ancorado em "Prestador do Serviço" - a NF também tem o CNPJ do tomador (o próprio IFFluminense)
# logo depois, sob o mesmo rótulo "CNPJ / CPF / NIF", então não dá pra buscar o rótulo sozinho
RE_CNPJ_PRESTADOR_NF = re.compile(
    r"Prestador do Servi[çc]o\s*\n?\s*CNPJ\s*/\s*CPF\s*/\s*NIF\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})",
    re.IGNORECASE)
# número do contrato e domicílio bancário nem toda NF cita - quando cita, vem dentro do texto
# livre da "Descrição do Serviço" (confirmado no euro.pdf: "...Conforme contrato nº68/2024..." e
# "...*Dados Bancários: Caixa Econômica Federal (104) | Ag 0212 | C/c 578536626-2*")
RE_CONTRATO_NF = re.compile(r"contrato\s*n[ºo°]\s*(\d+/\d+)", re.IGNORECASE)
RE_DADOS_BANCARIOS_NF = re.compile(
    r"Dados Banc[áa]rios:\s*([^|]+?)\s*\(\s*(\d+)\s*\)\s*\|\s*Ag\s*(\d+)\s*\|\s*C/c\s*([\d-]+)",
    re.IGNORECASE)

def obter_dados_nf(paginas):
    # None se a NF não foi localizada/lida (ex: página-imagem, sem texto) - quem chama trata
    # esse caso preenchendo a Fonte segura com o motivo, ao invés de um valor
    texto_nf, paginas_nf = ns.localizar_texto_nf(paginas)
    if not texto_nf:
        return None
    match_nf = ns.RE_NUMERO_NF.search(texto_nf)
    match_valor = RE_VALOR_SERVICO_NF.search(texto_nf)
    match_valor_liquido = RE_VALOR_LIQUIDO_NF.search(texto_nf)
    match_cnpj = RE_CNPJ_PRESTADOR_NF.search(texto_nf)
    match_contrato = RE_CONTRATO_NF.search(texto_nf)
    match_bancario = RE_DADOS_BANCARIOS_NF.search(texto_nf)
    return {
        "paginas": paginas_nf,
        "nf": match_nf.group(1) if match_nf else "",
        "emissao": ns.extrair_data_emissao_nf(texto_nf),
        "competencia": ns.extrair_competencia_nf(texto_nf),
        "valor": match_valor.group(1) if match_valor else "",
        "valor_liquido": match_valor_liquido.group(1) if match_valor_liquido else "",
        "cnpj": match_cnpj.group(1) if match_cnpj else "",
        "contrato": match_contrato.group(1) if match_contrato else "",
        "banco": f"{match_bancario.group(1).strip()} ({match_bancario.group(2)})" if match_bancario else "",
        "agencia": match_bancario.group(3) if match_bancario else "",
        "conta": match_bancario.group(4) if match_bancario else "",
    }

def pagina_nf_str(dados_nf):
    if not dados_nf or not dados_nf["paginas"]:
        return ""
    return f"NF pág. {ns.juntar_com_e(dados_nf['paginas'])}"

# ------- contrato no banco (fonte segura) -------

def localizar_contrato(paginas):
    # mesma lógica de localização de empresa do preencher_planilha_ns.py (linha "FAVORECIDO :" do
    # SIAFI) - só que aqui usa o CNPJ pra puxar o contrato INTEIRO do banco, não só a abreviação
    for texto in paginas:
        match = ns.RE_EMPRESA.search(texto)
        if match:
            contrato = contratos_db.obter_contrato_por_cnpj(match.group(1))
            if contrato:
                return contrato
    return None

def empenhos_registrados(contrato):
    # empenhos "diretos" (contrato de serviço) + empenhos dentro de processos de empenho
    # (contrato de almoxarifado) - achatados numa lista só de números, pra conferência
    numeros = [e["numero_empenho"] for e in contrato.get("empenhos", [])]
    for processo in contrato.get("processos_empenho", []):
        numeros.extend(e["numero_empenho"] for e in processo.get("empenhos", []))
    return numeros

# ------- montagem da tabela -------

def linha_tabela(campo, fonte_texto, fonte_disponivel, doc_texto, doc_disponivel, bate):
    # devolve um dict (não uma string pronta) - a janela HTML (ApiConformidade) manda isso direto
    # pro JS; formatar_bloco_markdown() é que sabe transformar isso em tabela markdown, só usado
    # pelo main()/CLI (rodar `python conformidade.py` sozinho, sem o gui.py)
    return {
        "campo": campo,
        "fonte": fonte_texto,
        "fonte_disponivel": bool(fonte_disponivel),
        "documento": doc_texto,
        "documento_disponivel": bool(doc_disponivel),
        "resultado": "ok" if bate is True else ("nao" if bate is False else "indefinido"),
    }

def montar_tabela(nome_arquivo, nome_documento, pagina, linhas, observacao=None):
    # observacao: nota de destaque (vermelha na janela HTML) mostrada no fim do bloco, pra campo
    # que o próprio documento determina e não tem fonte segura pra conferir contra (ex: Competência
    # e Valor do Termo Circunstanciado - ver processar_termo_circunstanciado)
    bloco = {"arquivo": nome_arquivo, "documento": nome_documento, "pagina": pagina, "linhas": linhas}
    if observacao:
        bloco["observacao"] = observacao
    return bloco

def formatar_bloco_markdown(bloco):
    # usado só pelo main()/CLI (rodar este arquivo sozinho) - a janela HTML não passa por aqui,
    # ela recebe os dicts direto e monta a tabela em HTML
    cabecalho = f"**{bloco['arquivo']} — {bloco['documento']} (pág. {bloco['pagina']})**"
    linhas_md = []
    for linha in bloco["linhas"]:
        fonte = f"`{linha['fonte']}`" if linha["fonte_disponivel"] else f"*({linha['fonte']})*"
        doc = f"`{linha['documento']}`" if linha["documento_disponivel"] else f"*({linha['documento']})*"
        resultado = {"ok": "✅", "nao": "❌", "indefinido": "➖"}[linha["resultado"]]
        linhas_md.append(f"| {linha['campo']} | {fonte} | {doc} | {resultado} |")
    corpo = "\n".join(["| Campo | Fonte segura | Documento | Resultado |", "|---|---|---|---|", *linhas_md])
    observacao = f"\n\n⚠️ {bloco['observacao']}" if bloco.get("observacao") else ""
    return f"{cabecalho}\n\n{corpo}{observacao}"

# ------- Documento 1: Relatório Circunstanciado de Recebimento Provisório -------

RE_CNPJ_PARENTESE = re.compile(r"\(\s*(?:CNPJ\s*)?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\s*\)")
RE_CONTRATO_GENERICO = re.compile(r"[Cc]ontrato\s*n[ºo°]?\s*([\d./]+)")
RE_NF_GENERICO = re.compile(r"Nota Fiscal\s*n[ºo°]?\s*(\d+)", re.IGNORECASE)
RE_VALOR_GENERICO = re.compile(r"no valor de\s*R\$\s*([\d.,]+)", re.IGNORECASE)
RE_PERIODO_DOCUMENTO = re.compile(r"no per[íi]odo de\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})")
RE_CONTRATADA_RELATORIO = re.compile(r"prestados pela Contratada\s*\n?(.+?)\n?\s*\(", re.DOTALL)

def processar_relatorio_circunstanciado(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer):
    # usa a ÚLTIMA ocorrência, não a primeira - documento pode ser refeito/corrigido no meio do
    # processo (mesmo raciocínio do RAMR/IMR, ver processar_relatorio_avaliacao_medicao), e a
    # versão mais recente sempre substitui as anteriores
    indices = [i for i, t in enumerate(paginas) if "Relatório Circunstanciado de Recebimento Provisório" in t]
    if not indices:
        return None
    indice = indices[-1]
    texto = remover_duplicatas_consecutivas(paginas[indice])

    linhas = []

    doc_processo = ns.RE_PROCESSO.search(texto)
    doc_processo_valor = doc_processo.group() if doc_processo else None
    linhas.append(linha_tabela(
        "Processo",
        f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1),
        doc_processo_valor or "não encontrado no documento", bool(doc_processo_valor),
        comparar_textos(processo_p1, doc_processo_valor) if processo_p1 and doc_processo_valor else None,
    ))

    m_contratada = RE_CONTRATADA_RELATORIO.search(texto)
    doc_contratada = limpar_espacos(m_contratada.group(1)) if m_contratada else None
    fonte_contratada = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Contratada", f"{fonte_contratada} (BD)" if fonte_contratada else "contrato não encontrado no banco", bool(fonte_contratada),
        doc_contratada or "não encontrado no documento", bool(doc_contratada),
        comparar_textos(fonte_contratada, doc_contratada) if fonte_contratada and doc_contratada else None,
    ))

    m_cnpj = RE_CNPJ_PARENTESE.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else None
    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_contrato_doc = RE_CONTRATO_GENERICO.search(texto)
    doc_contrato = m_contrato_doc.group(1) if m_contrato_doc else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    m_nf_doc = RE_NF_GENERICO.search(texto)
    doc_nf = m_nf_doc.group(1) if m_nf_doc else None
    linhas.append(linha_tabela(
        "Nota Fiscal", *_fonte_nf(dados_nf, "nf"),
        doc_nf or "não encontrado no documento", bool(doc_nf),
        comparar_numeros(dados_nf["nf"], doc_nf) if dados_nf and dados_nf["nf"] and doc_nf else None,
    ))

    (fonte_comp_texto, fonte_comp_disp), competencia_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "competencia")
    doc_comp_bruto, doc_comp = extrair_competencia_documento(texto)
    linhas.append(linha_tabela(
        "Competência", fonte_comp_texto, fonte_comp_disp,
        exibir_competencia(doc_comp_bruto, doc_comp) or "não encontrada no documento", bool(doc_comp_bruto),
        (competencia_ref == doc_comp) if competencia_ref and doc_comp else None,
    ))

    fonte_periodo = calcular_periodo(competencia_ref) if competencia_ref else None
    m_periodo = RE_PERIODO_DOCUMENTO.search(texto)
    doc_periodo = f"{m_periodo.group(1)} a {m_periodo.group(2)}" if m_periodo else None
    linhas.append(linha_tabela(
        "Período",
        f"{fonte_periodo} (calculado c/ base na competência)" if fonte_periodo else "depende da competência", bool(fonte_periodo),
        doc_periodo or "não encontrado no documento", bool(doc_periodo),
        (fonte_periodo == doc_periodo) if fonte_periodo and doc_periodo else None,
    ))

    (fonte_valor_texto, fonte_valor_disp), valor_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "valor")
    m_valor_doc = RE_VALOR_GENERICO.search(texto)
    doc_valor = m_valor_doc.group(1) if m_valor_doc else None
    linhas.append(linha_tabela(
        "Valor", fonte_valor_texto, fonte_valor_disp,
        doc_valor or "não encontrado no documento", bool(doc_valor),
        comparar_numeros(valor_ref, doc_valor) if valor_ref and doc_valor else None,
    ))

    return montar_tabela(nome_arquivo, "Relatório Circunstanciado de Recebimento Provisório", indice + 1, linhas)

# ------- Documento 2: Despacho de Ateste de Nota Fiscal de Serviço -------

RE_CONTRATADA_DESPACHO = re.compile(r"prestados pela empresa\s*\n?(.+?)\n?\s*,\s*\n?\s*constantes", re.DOTALL)
RE_TIPO_SERVICO = re.compile(r"referente\s+(?:à|ao servi[cç]o de)\s*\n?(.+?)\n?\s*,\s*\n?\s*competência",
                              re.DOTALL | re.IGNORECASE)

def processar_despacho_ateste(nome_arquivo, paginas, contrato, dados_nf, dados_parecer):
    # última ocorrência, não a primeira - mesmo raciocínio do RAMR/IMR (ver
    # processar_relatorio_avaliacao_medicao): se o despacho foi refeito/corrigido, o mais recente vale
    indices = [i for i, t in enumerate(paginas) if "Despacho de Ateste de Nota Fiscal" in t]
    if not indices:
        return None
    indice = indices[-1]
    texto = remover_duplicatas_consecutivas(paginas[indice])

    linhas = []

    m_contratada = RE_CONTRATADA_DESPACHO.search(texto)
    doc_contratada = None
    doc_cnpj = None
    if m_contratada:
        bruto = limpar_espacos(m_contratada.group(1))
        m_cnpj_dentro = RE_CNPJ_PARENTESE.search(bruto)
        doc_cnpj = m_cnpj_dentro.group(1) if m_cnpj_dentro else None
        doc_contratada = limpar_espacos(RE_CNPJ_PARENTESE.sub("", bruto))

    fonte_contratada = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Contratada", f"{fonte_contratada} (BD)" if fonte_contratada else "contrato não encontrado no banco", bool(fonte_contratada),
        doc_contratada or "não encontrado no documento", bool(doc_contratada),
        comparar_textos(fonte_contratada, doc_contratada) if fonte_contratada and doc_contratada else None,
    ))

    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    if doc_cnpj:
        linhas.append(linha_tabela(
            "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
            doc_cnpj, True,
            comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt else None,
        ))
    else:
        linhas.append(linha_tabela(
            "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
            "CNPJ não citado nesse documento", False, None,
        ))

    m_tipo = RE_TIPO_SERVICO.search(texto)
    doc_tipo = limpar_espacos(m_tipo.group(1)) if m_tipo else None
    fonte_objeto = contrato["objeto"] if contrato else None
    linhas.append(linha_tabela(
        "Tipo de serviço",
        f"{fonte_objeto} (BD)" if fonte_objeto else ("objeto não cadastrado no banco" if contrato else "contrato não encontrado no banco"),
        bool(fonte_objeto),
        doc_tipo or "não encontrado no documento", bool(doc_tipo),
        comparar_textos(fonte_objeto, doc_tipo) if fonte_objeto and doc_tipo else None,
    ))

    m_contrato_doc = RE_CONTRATO_GENERICO.search(texto)
    doc_contrato = m_contrato_doc.group(1) if m_contrato_doc else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    m_nf_doc = RE_NF_GENERICO.search(texto)
    doc_nf = m_nf_doc.group(1) if m_nf_doc else None
    linhas.append(linha_tabela(
        "Nota Fiscal", *_fonte_nf(dados_nf, "nf"),
        doc_nf or "não encontrado no documento", bool(doc_nf),
        comparar_numeros(dados_nf["nf"], doc_nf) if dados_nf and dados_nf["nf"] and doc_nf else None,
    ))

    (fonte_comp_texto, fonte_comp_disp), competencia_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "competencia")
    doc_comp_bruto, doc_comp = extrair_competencia_documento(texto)
    linhas.append(linha_tabela(
        "Competência", fonte_comp_texto, fonte_comp_disp,
        exibir_competencia(doc_comp_bruto, doc_comp) or "não encontrada no documento", bool(doc_comp_bruto),
        (competencia_ref == doc_comp) if competencia_ref and doc_comp else None,
    ))

    return montar_tabela(nome_arquivo, "Despacho de Ateste de Nota Fiscal de Serviço", indice + 1, linhas)

# ------- Documento 3: Instrumentos de Cobrança (contratos.gov.br) -------

RE_CONTRATO_INSTRUMENTO = re.compile(r"Contrato:\s*([\d./]+)")
RE_NUMERO_NF_INSTRUMENTO = re.compile(r"N[uú]mero:\s*(\d+)")
RE_DT_EMISSAO = re.compile(r"Dt\. Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})")
RE_VALOR_FATURADO = re.compile(r"Valor Faturado:\s*R\$\s*([\d.,]+)")
RE_VALOR_LIQUIDO = re.compile(r"Valor Liquido:\s*R\$\s*([\d.,]+)")

def processar_instrumento_cobranca(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer):
    # última ocorrência, não a primeira - mesmo raciocínio do RAMR/IMR (ver
    # processar_relatorio_avaliacao_medicao); confirmado real no prevelar.pdf, que tem 2 páginas de
    # Instrumentos de Cobrança com Dt. Emissão diferente (17/07 vs 16/07 - correção posterior)
    indices = [i for i, t in enumerate(paginas) if "Instrumentos de cobrança" in t]
    if not indices:
        return None
    indice = indices[-1]
    texto = paginas[indice]  # sem duplicação de linha nessa página (é tabela rótulo:valor, não texto corrido)

    linhas = []

    doc_processo = ns.RE_PROCESSO.search(texto)
    linhas.append(linha_tabela(
        "Processo", f"{processo_p1} (pág. 1)", bool(processo_p1),
        doc_processo.group() if doc_processo else "não encontrado no documento", bool(doc_processo),
        comparar_textos(processo_p1, doc_processo.group()) if processo_p1 and doc_processo else None,
    ))

    m_contrato_doc = RE_CONTRATO_INSTRUMENTO.search(texto)
    doc_contrato = m_contrato_doc.group(1) if m_contrato_doc else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    m_nf_doc = RE_NUMERO_NF_INSTRUMENTO.search(texto)
    doc_nf = m_nf_doc.group(1) if m_nf_doc else None
    linhas.append(linha_tabela(
        "Nota Fiscal", *_fonte_nf(dados_nf, "nf"),
        doc_nf or "não encontrado no documento", bool(doc_nf),
        comparar_numeros(dados_nf["nf"], doc_nf) if dados_nf and dados_nf["nf"] and doc_nf else None,
    ))

    m_dt_emissao = RE_DT_EMISSAO.search(texto)
    doc_emissao = m_dt_emissao.group(1) if m_dt_emissao else None
    linhas.append(linha_tabela(
        "Dt. Emissão", *_fonte_nf(dados_nf, "emissao"),
        doc_emissao or "não encontrada no documento", bool(doc_emissao),
        (dados_nf["emissao"] == doc_emissao) if dados_nf and dados_nf["emissao"] and doc_emissao else None,
    ))

    (fonte_comp_texto, fonte_comp_disp), competencia_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "competencia")
    doc_comp_bruto, doc_comp = extrair_competencia_documento(texto, aceitar_numerico=True)
    linhas.append(linha_tabela(
        "Competência", fonte_comp_texto, fonte_comp_disp,
        exibir_competencia(doc_comp_bruto, doc_comp) or "não encontrada no documento", bool(doc_comp_bruto),
        (competencia_ref == doc_comp) if competencia_ref and doc_comp else None,
    ))

    (fonte_valor_texto, fonte_valor_disp), valor_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "valor")

    m_valor_fat = RE_VALOR_FATURADO.search(texto)
    doc_valor_fat = m_valor_fat.group(1) if m_valor_fat else None
    linhas.append(linha_tabela(
        "Valor Faturado", fonte_valor_texto, fonte_valor_disp,
        doc_valor_fat or "não encontrado no documento", bool(doc_valor_fat),
        comparar_numeros(valor_ref, doc_valor_fat) if valor_ref and doc_valor_fat else None,
    ))

    m_valor_liq = RE_VALOR_LIQUIDO.search(texto)
    doc_valor_liq = m_valor_liq.group(1) if m_valor_liq else None
    linhas.append(linha_tabela(
        "Valor Líquido", fonte_valor_texto, fonte_valor_disp,
        doc_valor_liq or "não encontrado no documento", bool(doc_valor_liq),
        comparar_numeros(valor_ref, doc_valor_liq) if valor_ref and doc_valor_liq else None,
    ))

    doc_empenhos = ns.extrair_empenhos([texto])
    fonte_empenhos = empenhos_registrados(contrato) if contrato else []
    doc_empenhos_str = ", ".join(doc_empenhos) if doc_empenhos else None
    bate_empenhos = all(e in fonte_empenhos for e in doc_empenhos) if doc_empenhos and fonte_empenhos else None
    linhas.append(linha_tabela(
        "Empenhos",
        f"{', '.join(fonte_empenhos)} (BD)" if fonte_empenhos else ("nenhum empenho cadastrado nesse contrato" if contrato else "contrato não encontrado no banco"),
        bool(fonte_empenhos),
        doc_empenhos_str or "não encontrado no documento", bool(doc_empenhos_str),
        bate_empenhos,
    ))

    return montar_tabela(nome_arquivo, "Instrumentos de Cobrança", indice + 1, linhas)

# ------- Documento 4: Relatório de Avaliação e Medição dos Resultados (RAMR/IMR) -------

# rótulos em CAIXA ALTA e cada um seguido do valor numa linha própria (às vezes com o mês e o "/"
# em linhas separadas também) - formato confirmado no euro.pdf, bem diferente da prosa corrida dos
# outros 3 documentos, por isso os regexes daqui são todos dedicados, não reaproveitam os genéricos
RE_COMPETENCIA_RAMR = re.compile(r"COMPET[ÊE]NCIA:\s*\n?\s*([A-Za-zçÇãÃéÉêÊúÚ]+)\s*\n?\s*/\s*\n?\s*(\d{4})", re.IGNORECASE)
RE_VIGENCIA_RAMR = re.compile(r"VIG[ÊE]NCIA:\s*\n?\s*(\d{2}/\d{2}/\d{4})\s*A\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_CONTRATO_RAMR = re.compile(r"CONTRATO N[ºo°]\s*\n?\s*([\d./]+)", re.IGNORECASE)
RE_CONTRATADO_RAMR = re.compile(r"CONTRATADO:\s*\n?\s*(.+?)\s*\n\s*CNPJ:", re.IGNORECASE)
RE_CNPJ_RAMR = re.compile(r"CNPJ:\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.IGNORECASE)
RE_PERIODO_RAMR = re.compile(r"Per[íi]odo de Avalia[çc][ãa]o:\s*\n?\s*(\d{2}/\d{2}/\d{4})\s*A\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)

def _extrair_competencia_ramr(texto):
    # mesma ideia de extrair_competencia_documento(), mas rótulo próprio ("COMPETÊNCIA:" com dois
    # pontos, mês e ano em linhas separadas) - devolve (bruto_pra_exibir, normalizado) ou (None, None)
    m = RE_COMPETENCIA_RAMR.search(texto)
    if not m:
        return None, None
    numero_mes = _mes_para_numero(m.group(1))
    if not numero_mes:
        return None, None
    return f"{m.group(1)}/{m.group(2)}", f"{_MESES_NOME[numero_mes]}/{m.group(2)}"

def processar_relatorio_avaliacao_medicao(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer):
    # esse relatório costuma ser corrigido/refeito (o gestor do contrato pede um novo quando acha
    # erro no anterior, ex: vigência errada) - cada nova versão repete o mesmo título no PDF, então
    # usa a ÚLTIMA ocorrência (a mais recente sempre corrige/substitui as anteriores), não a
    # primeira, diferente dos outros 3 documentos (esses não costumam ser refeitos assim)
    indices = [i for i, t in enumerate(paginas) if "RELATÓRIO DE AVALIAÇÃO E MEDIÇÃO DOS RESULTADOS" in t]
    if not indices:
        return None
    indice = indices[-1]
    # a tabela de pontuação sempre estoura pra página seguinte (confirmado no euro.pdf - "Pontuação
    # Total do Serviço" nunca cabe na mesma página do título) - por isso junta as duas antes de
    # aplicar os regexes; os outros 3 documentos não precisam disso porque seus campos cabem numa página só
    texto_bruto = paginas[indice] + ("\n" + paginas[indice + 1] if indice + 1 < len(paginas) else "")
    texto = remover_duplicatas_consecutivas(texto_bruto)

    linhas = []

    doc_processo = ns.RE_PROCESSO.search(texto)
    doc_processo_valor = doc_processo.group() if doc_processo else None
    linhas.append(linha_tabela(
        "Processo",
        f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1),
        doc_processo_valor or "não encontrado no documento", bool(doc_processo_valor),
        comparar_textos(processo_p1, doc_processo_valor) if processo_p1 and doc_processo_valor else None,
    ))

    m_contratada = RE_CONTRATADO_RAMR.search(texto)
    doc_contratada = limpar_espacos(m_contratada.group(1)) if m_contratada else None
    fonte_contratada = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Contratada", f"{fonte_contratada} (BD)" if fonte_contratada else "contrato não encontrado no banco", bool(fonte_contratada),
        doc_contratada or "não encontrado no documento", bool(doc_contratada),
        comparar_textos(fonte_contratada, doc_contratada) if fonte_contratada and doc_contratada else None,
    ))

    m_cnpj = RE_CNPJ_RAMR.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else None
    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_contrato_doc = RE_CONTRATO_RAMR.search(texto)
    doc_contrato = m_contrato_doc.group(1) if m_contrato_doc else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    fonte_vigencia = None
    if contrato and contrato.get("vigencia_inicio") and contrato.get("vigencia_fim"):
        fonte_vigencia = f"{_formatar_data_iso(contrato['vigencia_inicio'])} A {_formatar_data_iso(contrato['vigencia_fim'])}"
    m_vigencia = RE_VIGENCIA_RAMR.search(texto)
    doc_vigencia = f"{m_vigencia.group(1)} A {m_vigencia.group(2)}" if m_vigencia else None
    linhas.append(linha_tabela(
        "Vigência", f"{fonte_vigencia} (BD)" if fonte_vigencia else "contrato não encontrado no banco", bool(fonte_vigencia),
        doc_vigencia or "não encontrada no documento", bool(doc_vigencia),
        (fonte_vigencia == doc_vigencia) if fonte_vigencia and doc_vigencia else None,
    ))

    (fonte_comp_texto, fonte_comp_disp), competencia_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "competencia")
    doc_comp_bruto, doc_comp = _extrair_competencia_ramr(texto)
    linhas.append(linha_tabela(
        "Competência", fonte_comp_texto, fonte_comp_disp,
        exibir_competencia(doc_comp_bruto, doc_comp) or "não encontrada no documento", bool(doc_comp_bruto),
        (competencia_ref == doc_comp) if competencia_ref and doc_comp else None,
    ))

    fonte_periodo = calcular_periodo(competencia_ref) if competencia_ref else None
    m_periodo = RE_PERIODO_RAMR.search(texto)
    doc_periodo = f"{m_periodo.group(1)} a {m_periodo.group(2)}" if m_periodo else None
    linhas.append(linha_tabela(
        "Período",
        f"{fonte_periodo} (calculado c/ base na competência)" if fonte_periodo else "depende da competência", bool(fonte_periodo),
        doc_periodo or "não encontrado no documento", bool(doc_periodo),
        (fonte_periodo == doc_periodo) if fonte_periodo and doc_periodo else None,
    ))

    return montar_tabela(nome_arquivo, "Relatório de Avaliação e Medição dos Resultados", indice + 1, linhas)

# ------- Documento 5: Termo Circunstanciado do Gestor do Contrato -------

# mesmo estilo de rótulo em CAIXA ALTA do RAMR (linha própria por valor), mas SEM a duplicação de
# linha da camada de acessibilidade (confirmado no euro.pdf - só remover_duplicatas_consecutivas()
# não muda nada aqui, mas não faz mal chamar mesmo assim, por consistência com os outros)
RE_PERIODO_TERMO = re.compile(r"PER[ÍI]ODO DE AVALIA[ÇC][ÃA]O:\s*\n?\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_CONTRATADA_TERMO = re.compile(r"CONTRATADA:\s*\n?\s*(.+?)\s*\n\s*CNPJ:", re.IGNORECASE)
RE_CNPJ_TERMO = re.compile(r"CNPJ:\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.IGNORECASE)
RE_OBJETO_TERMO = re.compile(r"OBJETO:\s*\n?\s*(.+?)\s*\n\s*N[ºo°] DO CONTRATO:", re.IGNORECASE)
RE_CONTRATO_TERMO = re.compile(r"N[ºo°] DO CONTRATO:\s*\n?\s*([\d./]+)", re.IGNORECASE)
RE_VIGENCIA_TERMO = re.compile(r"VIG[ÊE]NCIA:\s*\n?\s*(\d{2}/\d{2}/\d{4})\s*a\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_PROCESSO_EMPENHO_TERMO = re.compile(r"PROCESSO ANUAL DE EMPENHO:\s*\n?\s*([\d.\-]+)", re.IGNORECASE)
# campos do PARECER DO GESTOR DO CONTRATO, no fim do documento - únicas menções a valor/competência
RE_VALOR_TERMO = re.compile(r"no valor de\s*R\$\s*([\d.,]+)", re.IGNORECASE)
RE_COMPETENCIA_TERMO = re.compile(r"referente\s*[àa]\s*competência\s*([A-Za-zçÇãÃéÉêÊúÚ]+)\s*/\s*(\d{4})", re.IGNORECASE)

def processar_termo_circunstanciado(nome_arquivo, paginas, contrato):
    # mesmo raciocínio de "usar a última ocorrência" dos outros 4 processadores (ver
    # processar_relatorio_avaliacao_medicao) - esse termo também pode ser refeito
    indices = [i for i, t in enumerate(paginas) if "TERMO CIRCUNSTANCIADO DO GESTOR DO CONTRATO" in t]
    if not indices:
        return None
    indice = indices[-1]
    # PARECER DO GESTOR DO CONTRATO (valor/competência) fica na página seguinte à do título
    texto_bruto = paginas[indice] + ("\n" + paginas[indice + 1] if indice + 1 < len(paginas) else "")
    texto = remover_duplicatas_consecutivas(texto_bruto)

    linhas = []

    m_contratada = RE_CONTRATADA_TERMO.search(texto)
    doc_contratada = limpar_espacos(m_contratada.group(1)) if m_contratada else None
    fonte_contratada = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Contratada", f"{fonte_contratada} (BD)" if fonte_contratada else "contrato não encontrado no banco", bool(fonte_contratada),
        doc_contratada or "não encontrado no documento", bool(doc_contratada),
        comparar_textos(fonte_contratada, doc_contratada) if fonte_contratada and doc_contratada else None,
    ))

    m_cnpj = RE_CNPJ_TERMO.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else None
    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_objeto = RE_OBJETO_TERMO.search(texto)
    doc_objeto = limpar_espacos(m_objeto.group(1)) if m_objeto else None
    fonte_objeto = contrato["objeto"] if contrato else None
    linhas.append(linha_tabela(
        "Objeto",
        f"{fonte_objeto} (BD)" if fonte_objeto else ("objeto não cadastrado no banco" if contrato else "contrato não encontrado no banco"),
        bool(fonte_objeto),
        doc_objeto or "não encontrado no documento", bool(doc_objeto),
        comparar_textos(fonte_objeto, doc_objeto) if fonte_objeto and doc_objeto else None,
    ))

    m_contrato_doc = RE_CONTRATO_TERMO.search(texto)
    doc_contrato = m_contrato_doc.group(1) if m_contrato_doc else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    fonte_vigencia = None
    if contrato and contrato.get("vigencia_inicio") and contrato.get("vigencia_fim"):
        fonte_vigencia = f"{_formatar_data_iso(contrato['vigencia_inicio'])} a {_formatar_data_iso(contrato['vigencia_fim'])}"
    m_vigencia = RE_VIGENCIA_TERMO.search(texto)
    doc_vigencia = f"{m_vigencia.group(1)} a {m_vigencia.group(2)}" if m_vigencia else None
    linhas.append(linha_tabela(
        "Vigência", f"{fonte_vigencia} (BD)" if fonte_vigencia else "contrato não encontrado no banco", bool(fonte_vigencia),
        doc_vigencia or "não encontrada no documento", bool(doc_vigencia),
        (fonte_vigencia == doc_vigencia) if fonte_vigencia and doc_vigencia else None,
    ))

    m_processo_empenho = RE_PROCESSO_EMPENHO_TERMO.search(texto)
    doc_processo_empenho = m_processo_empenho.group(1) if m_processo_empenho else None
    fonte_processo_empenho = contrato["processo_empenho_anual"] if contrato else None
    linhas.append(linha_tabela(
        "Processo Anual de Empenho",
        f"{fonte_processo_empenho} (BD)" if fonte_processo_empenho else ("não cadastrado no banco" if contrato else "contrato não encontrado no banco"),
        bool(fonte_processo_empenho),
        doc_processo_empenho or "não encontrado no documento", bool(doc_processo_empenho),
        comparar_textos(fonte_processo_empenho, doc_processo_empenho) if fonte_processo_empenho and doc_processo_empenho else None,
    ))

    # Competência e Valor NÃO entram na conferência aqui - não há fonte segura pra eles: é o
    # próprio PARECER DO GESTOR DO CONTRATO (dentro deste documento) que os determina, e ele passa
    # a ser a fonte segura usada pelos OUTROS documentos do processo (ver obter_dados_parecer /
    # _fonte_autorizacao_ou_nf) e pela conferência da própria NF (ver processar_nota_fiscal).
    # Mostrados aqui só como observação, em destaque, pra quem olha esse bloco já ver os valores
    # que valem pro resto do processo.
    m_competencia = RE_COMPETENCIA_TERMO.search(texto)
    doc_comp = f"{_MESES_NOME[_mes_para_numero(m_competencia.group(1))]}/{m_competencia.group(2)}" \
        if m_competencia and _mes_para_numero(m_competencia.group(1)) else None
    m_valor_doc = RE_VALOR_TERMO.search(texto)
    doc_valor = m_valor_doc.group(1) if m_valor_doc else None
    observacao = f"Competência: {doc_comp or 'não encontrada no documento'} | Valor: {doc_valor or 'não encontrado no documento'}"

    fonte_periodo = calcular_periodo(doc_comp) if doc_comp else None
    m_periodo = RE_PERIODO_TERMO.search(texto)
    doc_periodo = f"{m_periodo.group(1)} a {m_periodo.group(2)}" if m_periodo else None
    linhas.append(linha_tabela(
        "Período",
        f"{fonte_periodo} (calculado c/ base na competência)" if fonte_periodo else "depende da competência", bool(fonte_periodo),
        doc_periodo or "não encontrado no documento", bool(doc_periodo),
        (fonte_periodo == doc_periodo) if fonte_periodo and doc_periodo else None,
    ))

    return montar_tabela(nome_arquivo, "Termo Circunstanciado do Gestor do Contrato", indice + 1, linhas, observacao)

# ------- Documento 6: Nota Fiscal (conferida contra o PARECER do Termo Circunstanciado) -------

def obter_dados_parecer(paginas):
    # o PARECER DO GESTOR DO CONTRATO (dentro do Termo Circunstanciado - ver
    # processar_termo_circunstanciado) é a autorização pra emissão da própria NF: quando existe,
    # Valor e Competência nele valem MAIS que a NF nesses 2 campos em qualquer documento do
    # processo (a NF que deve bater com o que foi autorizado, não o contrário) - ver
    # _fonte_autorizacao_ou_nf(). Devolve None se não há Termo Circunstanciado no processo, ou se
    # não achou nem valor nem competência nele.
    indices = [i for i, t in enumerate(paginas) if "TERMO CIRCUNSTANCIADO DO GESTOR DO CONTRATO" in t]
    if not indices:
        return None
    indice = indices[-1]
    # PARECER fica sempre na página seguinte à do título do Termo (mesma particularidade do RAMR)
    texto_bruto = paginas[indice] + ("\n" + paginas[indice + 1] if indice + 1 < len(paginas) else "")
    texto = remover_duplicatas_consecutivas(texto_bruto)

    m_valor = RE_VALOR_TERMO.search(texto)
    m_comp = RE_COMPETENCIA_TERMO.search(texto)
    if not m_valor and not m_comp:
        return None

    competencia = None
    if m_comp:
        numero_mes = _mes_para_numero(m_comp.group(1))
        if numero_mes:
            competencia = f"{_MESES_NOME[numero_mes]}/{m_comp.group(2)}"

    return {
        "pagina": indice + 2,
        "valor": m_valor.group(1) if m_valor else None,
        "competencia": competencia,
    }

def processar_nota_fiscal(nome_arquivo, dados_nf, dados_parecer, contrato):
    # só existe um bloco de conferência da NF quando há Termo Circunstanciado no processo - sem
    # ele, a própria NF já É a fonte segura usada por todos os outros documentos (não faz sentido
    # comparar a NF contra ela mesma)
    if not dados_parecer:
        return None

    pagina = dados_nf["paginas"][0] if dados_nf and dados_nf["paginas"] else dados_parecer["pagina"]

    linhas = []

    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    doc_cnpj = dados_nf.get("cnpj") if dados_nf else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado na NF", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    fonte_comp = dados_parecer.get("competencia")
    doc_comp = dados_nf["competencia"] if dados_nf else None
    linhas.append(linha_tabela(
        "Competência",
        f"{fonte_comp} (Gestor do Contrato pág. {dados_parecer['pagina']})" if fonte_comp else "não encontrada no Termo", bool(fonte_comp),
        doc_comp or "não encontrada na NF", bool(doc_comp),
        (fonte_comp == doc_comp) if fonte_comp and doc_comp else None,
    ))

    fonte_valor = dados_parecer.get("valor")
    doc_valor = dados_nf["valor"] if dados_nf else None
    linhas.append(linha_tabela(
        "Valor Bruto",
        f"{fonte_valor} (Gestor do Contrato pág. {dados_parecer['pagina']})" if fonte_valor else "não encontrado no Termo", bool(fonte_valor),
        doc_valor or "não encontrado na NF", bool(doc_valor),
        comparar_numeros(fonte_valor, doc_valor) if fonte_valor and doc_valor else None,
    ))

    # Contrato e Domicílio Bancário só entram na conferência quando a própria NF cita esses dados -
    # nem todo modelo de NFS-e traz isso na "Descrição do Serviço" (ver RE_CONTRATO_NF/RE_DADOS_BANCARIOS_NF)
    doc_contrato_nf = dados_nf.get("contrato") if dados_nf else None
    if doc_contrato_nf:
        fonte_contrato = contrato["numero_contrato"] if contrato else None
        linhas.append(linha_tabela(
            "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
            doc_contrato_nf, True,
            comparar_numeros(fonte_contrato, doc_contrato_nf) if fonte_contrato else None,
        ))

    doc_agencia_nf = dados_nf.get("agencia") if dados_nf else None
    doc_conta_nf = dados_nf.get("conta") if dados_nf else None
    if doc_agencia_nf and doc_conta_nf:
        fonte_banco = contrato["banco"] if contrato else None
        fonte_agencia = contrato["agencia"] if contrato else None
        fonte_conta = contrato["conta"] if contrato else None
        tem_fonte = bool(fonte_agencia and fonte_conta)
        fonte_texto = (f"{fonte_banco or '-'} | Ag {fonte_agencia} | C/c {fonte_conta} (BD)" if tem_fonte
                       else ("dados bancários não cadastrados no banco" if contrato else "contrato não encontrado no banco"))
        doc_texto = f"{dados_nf.get('banco') or '-'} | Ag {doc_agencia_nf} | C/c {doc_conta_nf}"
        bate_domicilio = (comparar_numeros(fonte_agencia, doc_agencia_nf) and comparar_numeros(fonte_conta, doc_conta_nf)) if tem_fonte else None
        linhas.append(linha_tabela("Domicílio Bancário", fonte_texto, tem_fonte, doc_texto, True, bate_domicilio))

    valor_liquido = dados_nf.get("valor_liquido") if dados_nf else None
    observacao = f"Valor líquido: {valor_liquido or 'não encontrado na NF'}"

    return montar_tabela(nome_arquivo, "Nota Fiscal", pagina, linhas, observacao)

# ------- helpers pequenos -------

def _fonte_nf(dados_nf, campo):
    # devolve (texto, disponivel) pro padrão de linha_tabela, já com o rótulo "NF pág. X"
    if not dados_nf or not dados_nf.get(campo):
        return ("NF não identificada", False)
    return (f"{dados_nf[campo]} ({pagina_nf_str(dados_nf)})", True)

def _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, campo):
    # campo: "valor" ou "competencia" - o PARECER do Termo Circunstanciado (ver
    # obter_dados_parecer) tem prioridade sobre a NF nesses 2 campos, quando disponível no
    # processo. Devolve ((fonte_texto, fonte_disponivel), valor_bruto) - valor_bruto é o que entra
    # nas comparações de igualdade contra o que está escrito em cada documento
    if dados_parecer and dados_parecer.get(campo):
        valor = dados_parecer[campo]
        return (f"{valor} (Gestor do Contrato pág. {dados_parecer['pagina']})", True), valor
    fonte_texto, disponivel = _fonte_nf(dados_nf, campo)
    valor = dados_nf[campo] if dados_nf and dados_nf.get(campo) else None
    return (fonte_texto, disponivel), valor

def _formatar_cnpj(digitos):
    digitos = re.sub(r"\D", "", digitos or "")
    if len(digitos) != 14:
        return digitos
    return f"{digitos[0:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:14]}"

def _formatar_data_iso(iso):
    # "2024-04-30" (formato salvo no banco, vindo do <input type="date"> de cadastrar_contrato.py) -> "30/04/2024"
    partes = (iso or "").split("-")
    if len(partes) != 3:
        return None
    ano, mes, dia = partes
    return f"{dia}/{mes}/{ano}"

# ------- ponto de entrada -------

def gerar_conformidade(nome_arquivo, paginas):
    if not paginas or "Processo Eletrônico" not in paginas[0]:
        return []  # não é um PDF de andamento de processo

    match_processo_p1 = ns.RE_PROCESSO.search(paginas[0])
    processo_p1 = match_processo_p1.group() if match_processo_p1 else None

    contrato = localizar_contrato(paginas)
    dados_nf = obter_dados_nf(paginas)
    dados_parecer = obter_dados_parecer(paginas)

    tabelas = []
    for funcao in (
        lambda: processar_relatorio_avaliacao_medicao(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer),
        lambda: processar_termo_circunstanciado(nome_arquivo, paginas, contrato),
        lambda: processar_nota_fiscal(nome_arquivo, dados_nf, dados_parecer, contrato),
        lambda: processar_relatorio_circunstanciado(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer),
        lambda: processar_despacho_ateste(nome_arquivo, paginas, contrato, dados_nf, dados_parecer),
        lambda: processar_instrumento_cobranca(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer),
    ):
        tabela = funcao()
        if tabela:
            tabelas.append(tabela)
    return tabelas

def coletar_fontes_pdf():
    # mesma dupla fonte que preencher_planilha_ro.py/ns.py usam: abas do Chrome (se disponível) +
    # PDFs abertos/baixados localmente - devolve [(nome_exibicao, [texto_por_pagina, ...]), ...],
    # junto com um aviso (string ou None) se o Chrome não estava disponível
    aviso = None
    try:
        options = webdriver.ChromeOptions()
        options.debugger_address = "127.0.0.1:9222"
        navegador = webdriver.Chrome(options=options)
    except WebDriverException:
        aviso = "Abas do Chrome com PDFs abertos não encontradas, analisados só os PDFs baixados."
        navegador = None

    fontes = []

    if navegador is not None:
        aba_original = navegador.current_window_handle
        for janela in navegador.window_handles:
            navegador.switch_to.window(janela)
            url = navegador.current_url
            if "djtools/process_progress2" not in url:
                continue
            leitor = PdfReader(ns.baixar_pdf_da_aba(navegador, url))
            fontes.append((url, [p.extract_text() or "" for p in leitor.pages]))
        navegador.switch_to.window(aba_original)

    caminhos_pdf = list(dict.fromkeys(
        pdf_aberto_windows.listar_pdfs_abertos() + pdf_aberto_windows.listar_pdfs_recentes()
    ))
    for caminho in caminhos_pdf:
        with open(caminho, "rb") as arquivo:
            leitor = PdfReader(arquivo)
            fontes.append((caminho, [p.extract_text() or "" for p in leitor.pages]))

    return fontes, aviso

def rodar_conferencia():
    # usado tanto pelo main()/CLI quanto pela janela (ApiConformidade.rodar_conferencia) - devolve
    # a lista de blocos (dicts) de todos os documentos encontrados em todos os PDFs disponíveis
    contratos_db.inicializar_db()
    fontes, aviso = coletar_fontes_pdf()
    blocos = []
    for nome_exibicao, paginas in fontes:
        blocos.extend(gerar_conformidade(nome_exibicao, paginas))
    return blocos, aviso


# ------- janela de resultado -------

class ApiConformidade:
    # ponte Python <-> JS desta janela - os dados já vêm prontos (calculados antes da janela
    # abrir, dentro de main()), então só existe pra entregar esse resultado fixo pro JS ler
    def __init__(self, blocos, aviso):
        self.blocos = blocos
        self.aviso = aviso

    def obter_resultado(self):
        return {"blocos": self.blocos, "aviso": self.aviso}

# mesmos tokens de cor de cadastrar_contrato.py, pra manter a identidade visual do projeto
HTML_CONFORMIDADE = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>CCRGCI - Resultado da Conformidade</title>
<style>
  :root {
    --mist: #f2f5f3;
    --cloud: #ffffff;
    --hairline: #dde4e0;
    --ink: #16201b;
    --ink-soft: #56625b;
    --ink-faint: #8a958e;
    --pine: #178c4e;
    --pine-deep: #0f6b3b;
    --pine-tint: #e2f5ea;
    --pine-tint-strong: #c3ecd6;
    --status-error: #d1453d;
    --status-error-tint: #fbe9e8;
    --shadow-1: 0 1px 2px rgba(20,32,27,0.07), 0 1px 1px rgba(20,32,27,0.05);
  }

  * { box-sizing: border-box; }

  html, body {
    margin: 0; height: 100%; background: var(--mist); color: var(--ink);
    font-family: "Segoe UI Variable Display", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
    font-size: 13.5px;
  }

  .pagina { max-width: 980px; margin: 0 auto; padding: 22px 26px 40px; }

  .cabecalho { display: flex; align-items: center; gap: 10px; margin-bottom: 16px; }
  h1 { margin: 0; font-size: 19px; font-weight: 600; letter-spacing: -0.01em; }
  .subtitulo { margin: 2px 0 0; color: var(--ink-soft); font-size: 12.5px; }

  .marca-icone {
    width: 34px; height: 34px; border-radius: 8px;
    background: var(--pine);
    color: #fff;
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-1); flex: 0 0 auto;
  }
  .marca-icone svg { width: 19px; height: 19px; }

  .aviso {
    display: flex; align-items: center; gap: 8px;
    background: var(--status-error-tint); border: 1px solid #f3c9c6; color: var(--status-error);
    border-radius: 8px; padding: 8px 12px; font-size: 12.5px; margin-bottom: 14px;
  }
  .aviso svg { flex: 0 0 auto; }

  .vazio { padding: 26px; text-align: center; color: var(--ink-faint); }

  .arquivo { margin-bottom: 22px; }
  .arquivo__titulo {
    display: flex; align-items: center; gap: 6px;
    font-size: 12.5px; font-weight: 600; color: var(--ink-soft);
    margin: 0 0 8px; word-break: break-all;
  }
  .arquivo__titulo svg { flex: 0 0 auto; color: var(--pine-deep); }

  .painel {
    border: 1px solid var(--hairline); border-radius: 10px; background: var(--pine-tint);
    box-shadow: var(--shadow-1); padding: 14px 16px; margin-bottom: 14px;
  }
  .painel h2 { margin: 0 0 10px; font-size: 13px; font-weight: 600; color: var(--pine-deep); display: flex; align-items: baseline; gap: 6px; }
  .painel h2 .pagina-doc { font-weight: 400; color: var(--ink-faint); font-size: 12px; }
  .observacao { margin: 10px 0 0; padding-top: 10px; border-top: 1px solid var(--pine-tint-strong); font-size: 12.5px; font-weight: 600; color: var(--status-error); }

  table { width: 100%; border-collapse: collapse; }
  th, td { text-align: left; padding: 8px 10px; font-size: 12.5px; border-bottom: 1px solid var(--hairline); vertical-align: top; }
  th { color: var(--ink-soft); font-weight: 700; font-size: 11px; text-transform: uppercase; letter-spacing: 0.02em; }
  tr:last-child td { border-bottom: none; }
  td.campo { font-weight: 400; white-space: nowrap; }
  td.resultado { text-align: center; width: 40px; }

  .valor { font-family: "Cascadia Code", Consolas, monospace; font-size: 12px; }
  .valor--indisponivel { font-family: inherit; font-style: italic; color: var(--ink-faint); }

  .badge { display: inline-flex; align-items: center; justify-content: center; }
  .badge--ok { color: var(--pine-deep); }
  .badge--nao { color: var(--status-error); }
  .badge--indefinido { color: var(--ink-faint); }
</style>
</head>
<body>

<div class="pagina">
  <div class="cabecalho">
    <div class="marca-icone"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.5 2.5 2.5 4.5-5"/></svg></div>
    <div>
      <h1>Resultado da Conformidade</h1>
      <p class="subtitulo">Documentos preenchidos x Fontes seguras (NF e BD).</p>
    </div>
  </div>
  <div id="conteudo"></div>
</div>

<script>
  const ICONE_ARQUIVO = '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M7 3h7l4 4v14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M14 3v4h4"/></svg>';
  const ICONE_AVISO = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><circle cx="12" cy="16.5" r="0.5" fill="currentColor"/><path d="M10.3 4.6 2.9 18a1.5 1.5 0 0 0 1.3 2.2h15.6a1.5 1.5 0 0 0 1.3-2.2L13.7 4.6a1.6 1.6 0 0 0-2.8 0Z"/></svg>';
  const BADGES = {
    ok: '<span class="badge badge--ok" title="Confere"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m8.5 12.5 2.5 2.5 4.5-5"/></svg></span>',
    nao: '<span class="badge badge--nao" title="Não confere"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6M15 9l-6 6"/></svg></span>',
    indefinido: '<span class="badge badge--indefinido" title="Não foi possível comparar"><svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 12h10"/></svg></span>',
  };

  function celulaValor(texto, disponivel) {
    const classe = disponivel ? "valor" : "valor valor--indisponivel";
    const spanTexto = document.createElement("span");
    spanTexto.className = classe;
    spanTexto.textContent = texto;
    return spanTexto.outerHTML;
  }

  function montarTabelaLinhas(linhas) {
    const corpo = linhas.map((linha) => `
      <tr>
        <td class="campo">${linha.campo}</td>
        <td>${celulaValor(linha.fonte, linha.fonte_disponivel)}</td>
        <td>${celulaValor(linha.documento, linha.documento_disponivel)}</td>
        <td class="resultado">${BADGES[linha.resultado]}</td>
      </tr>
    `).join("");
    return `
      <table>
        <thead><tr><th>Campo</th><th>Fonte segura</th><th>Documento</th><th>Resultado</th></tr></thead>
        <tbody>${corpo}</tbody>
      </table>
    `;
  }

  function montarPainelDocumento(bloco) {
    const observacao = bloco.observacao ? `<p class="observacao">${bloco.observacao}</p>` : "";
    return `
      <div class="painel">
        <h2>${bloco.documento} <span class="pagina-doc">pág. ${bloco.pagina}</span></h2>
        ${montarTabelaLinhas(bloco.linhas)}
        ${observacao}
      </div>
    `;
  }

  function agruparPorArquivo(blocos) {
    const grupos = new Map();
    blocos.forEach((bloco) => {
      if (!grupos.has(bloco.arquivo)) grupos.set(bloco.arquivo, []);
      grupos.get(bloco.arquivo).push(bloco);
    });
    return grupos;
  }

  function renderizar(resultado) {
    const alvo = document.getElementById("conteudo");
    let html = "";

    if (resultado.aviso) {
      html += `<div class="aviso">${ICONE_AVISO}${resultado.aviso}</div>`;
    }

    if (!resultado.blocos.length) {
      html += '<p class="vazio">Nenhum documento de conformidade encontrado nos PDFs disponíveis.</p>';
      alvo.innerHTML = html;
      return;
    }

    const grupos = agruparPorArquivo(resultado.blocos);
    grupos.forEach((blocosDoArquivo, arquivo) => {
      html += `
        <div class="arquivo">
          <p class="arquivo__titulo">${ICONE_ARQUIVO}${arquivo}</p>
          ${blocosDoArquivo.map(montarPainelDocumento).join("")}
        </div>
      `;
    });

    alvo.innerHTML = html;
  }

  window.addEventListener("pywebviewready", async () => {
    const resultado = await window.pywebview.api.obter_resultado();
    renderizar(resultado);
  });
</script>
</body>
</html>
"""

def abrir_janela(blocos, aviso):
    # cria a janela em cima da instância de webview já em execução (a principal do gui.py) -
    # os dados já foram calculados antes de chamar essa função, a janela só renderiza
    webview.create_window(
        "CCRGCI - Resultado da Conformidade (NS)", html=HTML_CONFORMIDADE,
        js_api=ApiConformidade(blocos, aviso), width=1040, height=780,
    )

def main(nome_planilha=None):
    # roda pelo card "Conferir Conformidade" do gui.py (background thread) - ao terminar,
    # abre a janela de resultado automaticamente, sem precisar de botão dedicado
    blocos, aviso = rodar_conferencia()
    if aviso:
        print(aviso)
    print("Resultado da Conformidade (NS):")
    if not blocos:
        print("Nenhum documento de para conferência encontrado nos PDFs disponíveis.")
    else:
        print(f"{len(blocos)} documento(s) conferido(s) - Abrindo janela com o resultado...")
    abrir_janela(blocos, aviso)

if __name__ == "__main__":
    main()
    webview.start()
