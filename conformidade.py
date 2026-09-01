import calendar
import io
import os
import re
import threading

import pytesseract
import webview  # mesma lib do gui.py - abre a janela de resultado em cima da instância já em execução
from PIL import Image
from pypdf import PdfReader
from selenium import webdriver
from selenium.common.exceptions import WebDriverException

import contratos_db
import janela_windows
import pdf_aberto_windows
import preencher_planilha_ns as ns  # reaproveita localizar_texto_nf, RE_NUMERO_NF, RE_EMPRESA,
                                     # extrair_data_emissao_nf, extrair_competencia_nf, juntar_com_e -
                                     # já testados/validados nesse módulo, não duplica aqui

# alguns documentos digitalizados dentro do processo (ex: Consulta Optante pelo Simples Nacional -
# print de tela da Receita Federal) não têm camada de texto, só a imagem - o caminho padrão do
# instalador do Tesseract no Windows; se não existir aqui, mantém o padrão do pytesseract (assume
# que está no PATH)
_TESSERACT_PADRAO = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(_TESSERACT_PADRAO):
    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_PADRAO

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

def _mes_digitado_para_numero(chave):
    # aceita o mês digitado à mão como nome completo (com/sem acento, qualquer caixa) OU abreviação
    # de >= 3 letras ("jun", "mar", "dez") - prefixo de 3 é sempre único entre os 12 meses. Número
    # ("06") é tratado por quem chama, não aqui.
    import unicodedata
    def sem_acento(v):
        return unicodedata.normalize("NFKD", v or "").encode("ascii", "ignore").decode().lower().strip()
    alvo = sem_acento(chave)
    if not alvo:
        return None
    for numero, nome in enumerate(_MESES_NOME):
        if not nome:
            continue
        nome_norm = sem_acento(nome)
        if nome_norm == alvo or (len(alvo) >= 3 and nome_norm.startswith(alvo)):
            return numero
    return None

def _normalizar_competencia_digitada(texto):
    # o valor digitado à mão na janela da NF (ex: "junho/2026", "jun/2026", "Junho / 2026",
    # "06/2026") tem que virar o MESMO formato que extrair_competencia_documento produz
    # ("Junho/2026") - senão a comparação (== puro) com o texto dos documentos falha só por
    # caixa/abreviação/mês numérico
    if not texto:
        return texto
    match = re.match(r"\s*([A-Za-zçÇãÃéÉêÊúÚ]+|\d{1,2})\s*/\s*(\d{4})\s*$", texto)
    if not match:
        return texto.strip()
    chave = match.group(1)
    numero_mes = int(chave) if chave.isdigit() else _mes_digitado_para_numero(chave)
    if not numero_mes or not (1 <= numero_mes <= 12):
        return texto.strip()
    return f"{_MESES_NOME[numero_mes]}/{match.group(2)}"

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

# janela de digitação manual (ver solicitar_dados_manuais_nf) - só os campos passados em
# {campos_html} aparecem; mesmos tokens de cor do resto do projeto
HTML_DIGITACAO_NF = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>CCRGCI - Dados da NF não identificados</title>
<style>
  :root {{
    --mist: #f2f5f3; --cloud: #ffffff; --hairline: #dde4e0;
    --ink: #16201b; --ink-soft: #56625b; --ink-faint: #8a958e;
    --pine: #178c4e; --pine-deep: #0f6b3b; --pine-tint: #e2f5ea; --pine-tint-strong: #c3ecd6;
    --status-error: #d1453d; --status-error-tint: #fbe9e8;
    --shadow-1: 0 1px 2px rgba(20,32,27,0.07), 0 1px 1px rgba(20,32,27,0.05);
  }}
  * {{ box-sizing: border-box; }}
  html, body {{
    margin: 0; height: 100%; background: var(--mist); color: var(--ink);
    font-family: "Segoe UI Variable Display", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
    font-size: 13.5px;
  }}
  .pagina {{ padding: 20px 22px; }}
  .cabecalho {{ display: flex; align-items: center; gap: 10px; margin-bottom: 14px; }}
  .marca-icone {{
    width: 34px; height: 34px; border-radius: 8px; background: var(--status-error); color: #fff;
    display: flex; align-items: center; justify-content: center; box-shadow: var(--shadow-1); flex: 0 0 auto;
  }}
  h1 {{ margin: 0; font-size: 16px; font-weight: 600; letter-spacing: -0.01em; }}
  .subtitulo {{ margin: 3px 0 0; color: var(--ink-soft); font-size: 12px; line-height: 1.5; }}
  .subtitulo .pagina-ref {{ color: var(--ink-faint); white-space: nowrap; }}
  .campo {{ display: flex; flex-direction: column; gap: 4px; margin-bottom: 12px; }}
  label {{ font-size: 11.5px; font-weight: 600; color: var(--ink-soft); }}
  input {{
    font-family: inherit; font-size: 13px; padding: 7px 9px;
    border: 1px solid var(--hairline); border-radius: 6px; background: var(--cloud); color: var(--ink);
  }}
  input:focus-visible {{ outline: none; border-color: var(--pine); box-shadow: 0 0 0 3px var(--pine-tint-strong); }}
  .rodape {{ display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }}
  .btn {{
    font-family: inherit; font-size: 12.5px; font-weight: 600; border: 1px solid transparent;
    border-radius: 6px; padding: 7px 16px; cursor: pointer; white-space: nowrap;
  }}
  .btn--outline {{ background: var(--cloud); border-color: var(--hairline); color: var(--ink); }}
  .btn--outline:hover {{ border-color: var(--pine); color: var(--pine-deep); }}
  .btn--acento {{ background: var(--pine); color: #fff; }}
  .btn--acento:hover {{ background: var(--pine-deep); }}
</style>
</head>
<body>
<div class="pagina">
  <div class="cabecalho">
    <div class="marca-icone"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4"/><circle cx="12" cy="16.5" r="0.5" fill="currentColor"/><path d="M10.3 4.6 2.9 18a1.5 1.5 0 0 0 1.3 2.2h15.6a1.5 1.5 0 0 0 1.3-2.2L13.7 4.6a1.6 1.6 0 0 0-2.8 0Z"/></svg></div>
    <div>
      <h1>Dados da NF não identificados</h1>
      <p class="subtitulo">{mensagem}</p>
    </div>
  </div>
  <form id="form-nf">
    {campos_html}
    <div class="rodape">
      <button type="button" class="btn btn--outline" id="botao-pular">Continuar sem preencher</button>
      <button type="submit" class="btn btn--acento">Salvar e continuar</button>
    </div>
  </form>
</div>
<script>
  function formatarDataBr(iso) {{
    if (!iso) return "";
    const [ano, mes, dia] = iso.split("-");
    return `${{dia}}/${{mes}}/${{ano}}`;
  }}

  async function concluir(dados) {{
    await window.pywebview.api.enviar(dados);
  }}

  document.getElementById("form-nf").addEventListener("submit", (evento) => {{
    evento.preventDefault();
    const dados = {{}};
    const campoEmissao = document.getElementById("campo-emissao");
    const campoCompetencia = document.getElementById("campo-competencia");
    const campoValor = document.getElementById("campo-valor");
    if (campoEmissao) dados.emissao = formatarDataBr(campoEmissao.value);
    if (campoCompetencia) dados.competencia = campoCompetencia.value.trim();
    if (campoValor) dados.valor = campoValor.value.trim();
    concluir(dados);
  }});

  document.getElementById("botao-pular").addEventListener("click", () => concluir({{}}));
</script>
</body>
</html>
"""

def _abrir_pdf_na_pagina(caminho_pdf, pagina):
    # só funciona quando a fonte é um PDF baixado localmente (nome_arquivo é um caminho de
    # arquivo de verdade) - quando a fonte é uma aba do Chrome (nome_arquivo é a URL do Suap), não
    # existe um arquivo separado pra abrir, então nem tenta. O fragmento "#page=N" é respeitado
    # pelo visualizador de PDF embutido do Edge/Chrome (o padrão de PDF no Windows)
    if not caminho_pdf or not os.path.isfile(caminho_pdf):
        return False
    caminho_url = caminho_pdf.replace("\\", "/")
    try:
        os.startfile(f"file:///{caminho_url}#page={pagina}")
        return True
    except OSError:
        return False

def solicitar_dados_manuais_nf(nome_arquivo, dados_nf):
    # a NF foi localizada (dados_nf existe - ver obter_dados_nf) mas Emissão/Competência/Valor
    # podem ter vindo vazios - acontece quando o modelo de NF é diferente do testado (rótulos
    # diferentes, layout em colunas que atrapalha o OCR - ver _extrair_texto_ocr, caso real do
    # prevelar.pdf) ou o texto simplesmente não deu pra extrair. Em vez de deixar "indefinido" no
    # resto da conferência, abre uma janela ANTES de fechar o processamento pedindo pra digitar
    # manualmente só esses 3 campos (não pede CNPJ - pedido explícito do usuário 2026-08-24),
    # tentando abrir o PDF já na página certa (ver _abrir_pdf_na_pagina). Bloqueia (thread em
    # segundo plano de rodar_conferencia() espera nessa função) até o usuário decidir por um dos
    # botões da própria janela - fechar pelo X não conta como decisão (ver _ao_fechar).
    if not dados_nf:
        return
    campos_faltando = [c for c in ("emissao", "competencia", "valor") if not dados_nf.get(c)]
    if not campos_faltando:
        return

    campos_form = {
        "emissao": '<div class="campo"><label>Data de Emissão</label><input type="date" id="campo-emissao"></div>',
        "competencia": '<div class="campo"><label>Competência</label><input id="campo-competencia" placeholder="Ex: Junho/2026"></div>',
        "valor": '<div class="campo"><label>Valor do Serviço (R$)</label><input id="campo-valor" placeholder="Ex: 7.130,22"></div>',
    }
    campos_html = "\n    ".join(campos_form[campo] for campo in campos_faltando)

    pagina_numero = dados_nf["paginas"][0] if dados_nf["paginas"] else None
    abriu_pdf = _abrir_pdf_na_pagina(nome_arquivo, pagina_numero) if pagina_numero else False
    # instrução primeiro, depois a página em cinza (mesmo tom --ink-faint dos resultados)
    if dados_nf["paginas"]:
        paginas_texto = ns.juntar_com_e(dados_nf["paginas"])
        pagina_ref = f' <span class="pagina-ref">pág. {paginas_texto}</span>'
    else:
        pagina_ref = ' <span class="pagina-ref">página não identificada</span>'
    instrucao = "Preencha os campos abaixo, conforme a NF." if abriu_pdf else "Abra o documento e preencha os campos abaixo."
    mensagem = f"{instrucao}{pagina_ref}"

    html = HTML_DIGITACAO_NF.format(mensagem=mensagem, campos_html=campos_html)

    evento = threading.Event()
    resultado = {}

    class ApiDigitacaoNF:
        def enviar(self, dados):
            resultado.update(dados)
            evento.set()

    janela = webview.create_window(
        "CCRGCI - Dados da NF não identificados", html=html, js_api=ApiDigitacaoNF(),
        width=600, height=400, on_top=True,
    )

    def _ao_fechar():
        # cancela o fechamento pelo X enquanto não decidiu por um dos botões da própria janela
        # (retornar False no evento "closing" do pywebview veta o fechamento)
        if not evento.is_set():
            return False

    janela.events.closing += _ao_fechar
    evento.wait()
    janela.destroy()  # decidiu (por um botão da própria janela) - fecha sozinha, não precisa mais dela

    digitados = dados_nf.setdefault("_campos_digitados", set())
    for campo in campos_faltando:
        valor = resultado.get(campo)
        if valor:
            if campo == "competencia":
                valor = _normalizar_competencia_digitada(valor)  # "junho/2026" -> "Junho/2026"
            dados_nf[campo] = valor
            digitados.add(campo)

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

def naturezas_despesa_registradas(contrato):
    # natureza_despesa é um campo por EMPENHO (não do contrato em si) - achata e tira duplicatas,
    # preservando ordem, igual empenhos_registrados()
    naturezas = [e.get("natureza_despesa") for e in contrato.get("empenhos", []) if e.get("natureza_despesa")]
    for processo in contrato.get("processos_empenho", []):
        naturezas.extend(e.get("natureza_despesa") for e in processo.get("empenhos", []) if e.get("natureza_despesa"))
    return list(dict.fromkeys(naturezas))

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
    pagina = f" (pág. {bloco['pagina']})" if bloco["pagina"] else ""
    sigla = f" ({bloco['sigla']})" if bloco.get("sigla") else ""
    cabecalho = f"**{bloco['arquivo']} — {bloco['documento']}{sigla}{pagina}**"
    linhas_md = []
    for linha in bloco["linhas"]:
        fonte = f"`{linha['fonte']}`" if linha["fonte_disponivel"] else f"*({linha['fonte']})*"
        doc = f"`{linha['documento']}`" if linha["documento_disponivel"] else f"*({linha['documento']})*"
        resultado = {"ok": "✅", "nao": "❌", "indefinido": "➖"}[linha["resultado"]]
        linhas_md.append(f"| {linha['campo']} | {fonte} | {doc} | {resultado} |")
    observacao = f"\n\n⚠️ {bloco['observacao']}" if bloco.get("observacao") else ""
    if not bloco["linhas"]:  # bloco só de observação (ex: Documentos não localizados)
        return f"{cabecalho}{observacao}"
    corpo = "\n".join(["| Campo | Fonte segura | Documento | Resultado |", "|---|---|---|---|", *linhas_md])
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
    fonte_texto, fonte_disponivel, bate_objeto = _conferir_objeto(contrato, doc_tipo)
    linhas.append(linha_tabela(
        "Tipo de serviço", fonte_texto, fonte_disponivel,
        doc_tipo or "não encontrado no documento", bool(doc_tipo),
        bate_objeto,
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
RE_IC_OPTANTE = re.compile(r"Optante pelo Simples:\s*(\w+)", re.IGNORECASE)

def _linha_optante_simples(texto, dados_optante):
    # "Optante pelo Simples: <Sim|Não>" do Instrumento de Cobrança conferido contra a tela da
    # Receita Federal (Consulta Optante) - a Receita é a fonte segura; divergência = erro DA IC.
    # Vale pros dois fluxos (serviço e almoxarifado). dados_optante vem de obter_dados_optante.
    m = RE_IC_OPTANTE.search(texto)
    doc_opt = None
    if m:
        doc_opt = "Optante" if m.group(1).strip().lower().startswith("s") else "Não optante"
    fonte_opt = (dados_optante or {}).get("situacao")
    return linha_tabela(
        "Optante pelo Simples",
        f"{fonte_opt} (Consulta Optante pág. {dados_optante['pagina']})" if fonte_opt else "Consulta Optante não localizada no processo",
        bool(fonte_opt),
        doc_opt or "não encontrado no documento", bool(doc_opt),
        (fonte_opt == doc_opt) if fonte_opt and doc_opt else None,
    )

def processar_instrumento_cobranca(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer, dados_optante=None):
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

    linhas.append(_linha_optante_simples(texto, dados_optante))

    return montar_tabela(nome_arquivo, "Instrumentos de Cobrança", indice + 1, linhas)

# ------- Documento 7: Consulta Optante pelo Simples Nacional (Receita Federal) -------

# print de tela digitalizado (sem camada de texto) - o texto usado aqui já vem do fallback de OCR
# em _extrair_texto_ocr() (ver coletar_fontes_pdf), não do extract_text() normal do PDF
RE_CNPJ_OPTANTE = re.compile(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.IGNORECASE)
RE_SITUACAO_SIMPLES = re.compile(r"Situa[çc][ãa]o no Simples Nacional:\s*(N[ÃA]O\s+optante|Optante)\s+pelo Simples Nacional", re.IGNORECASE)

def obter_dados_optante(paginas):
    # a tela da Receita Federal "Consulta Optante pelo Simples Nacional" (imagem/OCR) é a FONTE
    # SEGURA da situação no Simples - outros documentos que citam isso (ex: o Instrumento de
    # Cobrança) têm que bater com ela. Devolve {"pagina", "situacao": "Optante"|"Não optante"} ou None.
    indices = [i for i, t in enumerate(paginas) if "SIMPLES NACIONAL" in t.upper() and "SIMEI" in t.upper()]
    if not indices:
        return None
    indice = indices[-1]
    m = RE_SITUACAO_SIMPLES.search(paginas[indice])
    if not m:
        return None
    return {
        "pagina": indice + 1,
        "situacao": "Não optante" if m.group(1).strip().upper().startswith("N") else "Optante",
    }

def processar_consulta_optante(nome_arquivo, paginas, contrato):
    # identifica a página pelo conteúdo (não tem um título fixo pra buscar, é só o print da tela da
    # Receita Federal) - "SIMPLES NACIONAL" + "SIMEI" juntos são específicos o bastante dessa consulta
    indices = [i for i, t in enumerate(paginas) if "SIMPLES NACIONAL" in t.upper() and "SIMEI" in t.upper()]
    if not indices:
        return None
    indice = indices[-1]
    texto = paginas[indice]

    linhas = []

    m_cnpj = RE_CNPJ_OPTANTE.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else None
    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_situacao = RE_SITUACAO_SIMPLES.search(texto)
    situacao = ("Não optante" if m_situacao.group(1).strip().upper().startswith("N") else "Optante") if m_situacao else None
    observacao = situacao or "situação no Simples Nacional não encontrada no documento"

    return montar_tabela(nome_arquivo, "Consulta Optante pelo Simples Nacional", indice + 1, linhas, observacao)

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
    fonte_texto, fonte_disponivel, bate_objeto = _conferir_objeto(contrato, doc_objeto)
    linhas.append(linha_tabela(
        "Objeto", fonte_texto, fonte_disponivel,
        doc_objeto or "não encontrado no documento", bool(doc_objeto),
        bate_objeto,
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

def processar_nota_fiscal(nome_arquivo, dados_nf, dados_parecer, contrato, sem_mao_de_obra=False):
    # COM PARECER (contrato de serviço com mão de obra): confere Competência e Valor Bruto contra o
    # PARECER do Gestor, além de CNPJ/Contrato/Domicílio contra o BD.
    # SEM PARECER: só há bloco quando é contrato de serviço SEM mão de obra - aí confere só o que tem
    # fonte no BD (CNPJ sempre; Contrato e Domicílio quando a NF os cita) e o Valor Bruto contra o
    # valor mensal fixo do contrato, SE houver (ver _valor_mensal_fixo); sem valor fixo cadastrado, o
    # Valor Bruto não é conferido, só listado na observação junto do Valor Líquido. Competência não
    # entra (sem PARECER e sem fonte no BD, não há contra o que comparar).
    # SEM PARECER e sem ser contrato sem mão de obra (ex: contrato não identificado no BD): não há
    # bloco - a NF já é a fonte segura dos outros documentos, comparar a NF contra ela mesma não serve.
    if not dados_parecer and not sem_mao_de_obra:
        return None

    pagina = (dados_nf["paginas"][0] if dados_nf and dados_nf["paginas"]
              else (dados_parecer["pagina"] if dados_parecer else None))

    linhas = []
    observacao_partes = []

    fonte_cnpj = contrato["cnpj"] if contrato else None
    fonte_cnpj_fmt = _formatar_cnpj(fonte_cnpj) if fonte_cnpj else None
    doc_cnpj = dados_nf.get("cnpj") if dados_nf else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado na NF", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    if dados_parecer:
        fonte_comp = dados_parecer.get("competencia")
        doc_comp = dados_nf["competencia"] if dados_nf else None
        linhas.append(linha_tabela(
            "Competência",
            f"{fonte_comp} (Gestor do Contrato pág. {dados_parecer['pagina']})" if fonte_comp else "não encontrada no Termo", bool(fonte_comp),
            doc_comp or "não encontrada na NF", bool(doc_comp),
            (fonte_comp == doc_comp) if fonte_comp and doc_comp else None,
        ))

    # Valor Bruto: contra o PARECER (com mão de obra) ou contra o valor mensal fixo do BD (sem mão
    # de obra); sem nenhuma das duas fontes, cai na observação, do mesmo jeito que o Valor Líquido
    doc_valor = dados_nf.get("valor") if dados_nf else None
    if dados_parecer:
        fonte_valor = dados_parecer.get("valor")
        linhas.append(linha_tabela(
            "Valor Bruto",
            f"{fonte_valor} (Gestor do Contrato pág. {dados_parecer['pagina']})" if fonte_valor else "não encontrado no Termo", bool(fonte_valor),
            doc_valor or "não encontrado na NF", bool(doc_valor),
            comparar_numeros(fonte_valor, doc_valor) if fonte_valor and doc_valor else None,
        ))
    else:
        valor_fixo = _valor_mensal_fixo(contrato)
        if valor_fixo is not None:
            nf_float = _valor_para_float(doc_valor)
            linhas.append(linha_tabela(
                "Valor Bruto", f"{_float_para_valor_br(valor_fixo)} (valor mensal fixo, BD)", True,
                doc_valor or "não encontrado na NF", bool(doc_valor),
                (abs(valor_fixo - nf_float) < 0.005) if nf_float is not None else None,
            ))
        else:
            observacao_partes.append(f"Valor bruto: {doc_valor or 'não encontrado na NF'}")

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
    observacao_partes.append(f"Valor líquido: {valor_liquido or 'não encontrado na NF'}")
    observacao = " | ".join(observacao_partes)

    return montar_tabela(nome_arquivo, "Nota Fiscal", pagina, linhas, observacao)

# ------- helpers pequenos -------

def _fonte_nf(dados_nf, campo):
    # devolve (texto, disponivel) pro padrão de linha_tabela, já com o rótulo "NF pág. X" - quando
    # o campo veio de digitação manual (ver solicitar_dados_manuais_nf), acrescenta "- digitado"
    # nesse rótulo, pra deixar claro que não foi extraído automaticamente do PDF
    if not dados_nf or not dados_nf.get(campo):
        return ("NF não identificada", False)
    sufixo = " - digitado" if campo in dados_nf.get("_campos_digitados", ()) else ""
    return (f"{dados_nf[campo]} ({pagina_nf_str(dados_nf)}{sufixo})", True)

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

def _conferir_objeto(contrato, doc_texto):
    # o objeto do contrato virou 2 campos (objeto_resumido/objeto_detalhado - ver
    # [[project-contratos-db-schema]]) - cada documento cita uma versão diferente (o Despacho usa a
    # descrição detalhada, o Termo Circunstanciado usa a resumida, por exemplo) - bate se o texto
    # do documento conferir com QUALQUER um dos dois, não precisa ser sempre o mesmo
    candidatos = []  # (rótulo, valor) - só os cadastrados
    if contrato and contrato.get("objeto_resumido"):
        candidatos.append(("Resumido", contrato["objeto_resumido"]))
    if contrato and contrato.get("objeto_detalhado"):
        candidatos.append(("Detalhado", contrato["objeto_detalhado"]))

    if not candidatos:
        return ("objeto não cadastrado no banco" if contrato else "contrato não encontrado no banco"), False, None

    correspondente = next(
        ((rotulo, valor) for rotulo, valor in candidatos if comparar_textos(valor, doc_texto)), None
    ) if doc_texto else None

    if correspondente:
        # bateu com um dos dois - mostra só esse (pedido do usuário), não os dois
        return f"{correspondente[0]}: {correspondente[1]} (BD)", True, True
    # não bateu nenhum (ou não há texto no documento) - mostra os dois, pra ver o que foi conferido
    fonte_texto = " | ".join(f"{rotulo}: {valor}" for rotulo, valor in candidatos) + " (BD)"
    return fonte_texto, True, (False if doc_texto else None)

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

def _valor_mensal_fixo(contrato):
    # soma dos itens de "Valores Mensais" do contrato no BD - alguns contratos de serviço SEM mão de
    # obra têm um valor mensal fixo cadastrado, e ele é a fonte segura do Valor Bruto da NF (não há
    # PARECER do Gestor nesses). None quando não há nenhum valor mensal cadastrado (campo em branco
    # no BD) - nesse caso o Valor Bruto vira só informativo (ver processar_nota_fiscal).
    if not contrato:
        return None
    total = sum(item.get("valor") or 0 for item in (contrato.get("valores_mensais") or []))
    return total if total > 0 else None

def _valor_para_float(texto):
    # "7.130,22" / "7130,22" / "7130.22" -> 7130.22 ; None se não parseável
    if texto in (None, ""):
        return None
    limpo = re.sub(r"[^\d,.]", "", str(texto))
    if not limpo:
        return None
    if "," in limpo:                 # formato BR: "," decimal, "." milhar
        limpo = limpo.replace(".", "").replace(",", ".")
    elif limpo.count(".") > 1:       # "1.234.567" sem decimais -> só milhares
        limpo = limpo.replace(".", "")
    try:
        return float(limpo)
    except ValueError:
        return None

def _float_para_valor_br(numero):
    # 7130.22 -> "7.130,22" (mesmo formato dos valores extraídos da NF)
    return f"{numero:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")

# ------- Consistência entre documentos (mesmo dado, comparado entre os próprios documentos) -------

# rótulo canônico -> rótulos usados nas tabelas dos outros documentos que representam o mesmo dado
# (alguns chamam a mesma coisa por nomes diferentes - ex: a NF chama "Valor Bruto" o que os outros
# chamam "Valor"; o Instrumento de Cobrança tem "Valor Faturado" e "Valor Líquido", que nesse
# sistema são sempre iguais ao "Valor do Serviço" da própria NF - ver Documento 1)
_CAMPOS_CONSISTENCIA = {
    "CNPJ": ["CNPJ"],
    "Processo": ["Processo"],
    "Nota Fiscal": ["Nota Fiscal"],
    "Competência": ["Competência"],
    "Período": ["Período"],
    "Valor": ["Valor", "Valor Bruto", "Valor Faturado", "Valor Líquido"],
    "Empenhos": ["Empenhos"],
}

def _comparar_conjuntos(a, b):
    # "Empenhos" é uma lista separada por vírgula (pode ter mais de um) - compara os números em si,
    # não a string inteira (ordem/espaçamento não deveriam importar)
    return {v.strip() for v in a.split(",") if v.strip()} == {v.strip() for v in b.split(",") if v.strip()}

def _valores_monetarios_batem(a, b):
    # "Valor" comparado pelo número em si (centavos), tolerando formatos diferentes ("7.130,22" x
    # "7130,22" x "7130.22") - comparar_numeros quebra com separador de milhar junto do decimal
    fa, fb = _valor_para_float(a), _valor_para_float(b)
    return fa is not None and fb is not None and abs(fa - fb) < 0.005

def _valor_ate(a, b):
    # a <= b (valores numéricos/monetários), com tolerância de meio centavo - usado quando a
    # entrega pode ser MENOR que o autorizado (NF <= OF)
    fa, fb = _valor_para_float(a), _valor_para_float(b)
    return fa is not None and fb is not None and fa <= fb + 0.005

_COMPARADOR_CONSISTENCIA = {
    "CNPJ": comparar_cnpjs,
    "Processo": comparar_textos,
    "Competência": lambda a, b: a == b,  # já vem normalizado (ver _valor_comparavel)
    "Período": lambda a, b: a == b,
    "Valor": _valores_monetarios_batem,
    "Nota Fiscal": comparar_numeros,
    "Empenhos": _comparar_conjuntos,
}

def _valor_comparavel(texto_documento):
    # tira o prefixo "bruto → " quando a exibição precisou de interpretação (ex: "Pagamento de
    # junho de 2026 → Junho/2026") - pra comparação entre documentos só importa o valor normalizado
    return texto_documento.split(" → ")[-1] if " → " in texto_documento else texto_documento

# nome completo do bloco (ver montar_tabela) -> como aparece na coluna "Documento" da consistência -
# só nessa tabela, pra não ficar gigante juntando vários nomes numa célula só (pedido do usuário)
_NOMES_CURTOS_DOCUMENTO = {
    "Relatório de Avaliação e Medição dos Resultados": "Gestor Contrato",
    "Termo Circunstanciado do Gestor do Contrato": "Termo Circunstanciado",
    "Nota Fiscal": "Nota Fiscal",
    "Relatório Circunstanciado de Recebimento Provisório": "Relatório Circunstanciado",
    "Despacho de Ateste de Nota Fiscal de Serviço": "Despacho",
    "Instrumentos de Cobrança": "IC",
    "Consulta Optante pelo Simples Nacional": "Consulta Optante",
}

def processar_consistencia_documentos(nome_arquivo, blocos, processo_p1, contrato, dados_nf, dados_parecer):
    # roda DEPOIS dos outros processadores (precisa da lista de blocos já pronta pra vasculhar) -
    # por lógica de negócio, a mesma informação conferida em documentos distintos do mesmo processo
    # precisa ser igual entre si, não só bater cada uma isoladamente contra a fonte segura (pedido
    # do usuário 2026-08-24). Só entra na comparação quem tem pelo menos 2 documentos com o dado
    # disponível - com só 1, não há o que comparar.
    (fonte_comp_texto, fonte_comp_disp), competencia_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "competencia")
    (fonte_valor_texto, fonte_valor_disp), valor_ref = _fonte_autorizacao_ou_nf(dados_parecer, dados_nf, "valor")
    # sem PARECER e com valor mensal fixo cadastrado (contrato de serviço sem mão de obra), o valor
    # fixo do BD é a referência de Valor - não a NF (mesma fonte usada no bloco da própria NF)
    valor_fixo = None if (dados_parecer and dados_parecer.get("valor")) else _valor_mensal_fixo(contrato)
    if valor_fixo is not None:
        valor_fixo_br = _float_para_valor_br(valor_fixo)
        fonte_valor_texto, fonte_valor_disp, valor_ref = f"{valor_fixo_br} (valor mensal fixo, BD)", True, valor_fixo_br
    fonte_periodo = calcular_periodo(competencia_ref) if competencia_ref else None
    fonte_nf_texto, fonte_nf_disp = _fonte_nf(dados_nf, "nf")
    nf_ref = dados_nf["nf"] if dados_nf and dados_nf.get("nf") else None
    fonte_cnpj_fmt = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    fonte_empenhos_lista = empenhos_registrados(contrato) if contrato else []
    fonte_empenhos_ref = ", ".join(fonte_empenhos_lista) if fonte_empenhos_lista else None

    fontes = {
        "CNPJ": (f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt), fonte_cnpj_fmt),
        "Processo": (f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1), processo_p1),
        "Competência": (fonte_comp_texto, fonte_comp_disp, competencia_ref),
        "Período": (f"{fonte_periodo} (calculado c/ base na competência)" if fonte_periodo else "depende da competência", bool(fonte_periodo), fonte_periodo),
        "Valor": (fonte_valor_texto, fonte_valor_disp, valor_ref),
        "Nota Fiscal": (fonte_nf_texto, fonte_nf_disp, nf_ref),
        "Empenhos": (f"{fonte_empenhos_ref} (BD)" if fonte_empenhos_ref else ("nenhum empenho cadastrado nesse contrato" if contrato else "contrato não encontrado no banco"), bool(fonte_empenhos_ref), fonte_empenhos_ref),
    }

    linhas = []
    for campo, rotulos in _CAMPOS_CONSISTENCIA.items():
        ocorrencias = []
        for bloco in blocos:
            for linha in bloco["linhas"]:
                if linha["campo"] in rotulos and linha["documento_disponivel"]:
                    nome_curto = _NOMES_CURTOS_DOCUMENTO.get(bloco["documento"], bloco["documento"])
                    ocorrencias.append((nome_curto, _valor_comparavel(linha["documento"])))

        if not ocorrencias:
            continue

        fonte_texto, fonte_disponivel, fonte_valor = fontes[campo]
        comparador = _COMPARADOR_CONSISTENCIA[campo]
        if fonte_valor:
            # tem fonte segura - confere cada documento contra ela, mesmo que só 1 documento tenha
            # o campo (ex: Empenhos hoje só aparece no Instrumento de Cobrança)
            referencia = fonte_valor
            bate = all(comparador(referencia, valor) for _, valor in ocorrencias)
        elif len(ocorrencias) >= 2:
            # sem fonte segura, mas ≥ 2 documentos - confere pelo menos entre eles
            referencia = ocorrencias[0][1]
            bate = all(comparador(referencia, valor) for _, valor in ocorrencias)
        else:
            bate = None  # só 1 documento e sem fonte segura - não há nada pra confrontar de verdade

        # quando bate, mostra só a informação em si (já é a mesma em todos os documentos, não
        # precisa repetir - isso já está detalhado nos blocos de cada documento, acima) - só
        # discrimina documento por documento quando NÃO bate, pra ajudar a achar a divergência
        doc_texto = ocorrencias[0][1] if bate else " | ".join(f"{nome}: {valor}" for nome, valor in ocorrencias)
        linhas.append(linha_tabela(campo, fonte_texto, fonte_disponivel, doc_texto, True, bate))

    # Processo Anual de Empenho e Natureza de Despesa não vêm de nenhum documento (só do banco) -
    # não têm o que comparar entre documentos, mas o usuário quer eles em destaque aqui mesmo assim
    partes_observacao = []
    if contrato and contrato.get("processo_empenho_anual"):
        partes_observacao.append(f"Processo Anual de Empenho: {contrato['processo_empenho_anual']}")
    naturezas = naturezas_despesa_registradas(contrato) if contrato else []
    if naturezas:
        partes_observacao.append(f"ND: {', '.join(naturezas)}")
    # ambos vêm só do banco - "(BD)" no fim da linha deixa a fonte explícita, igual às células da tabela
    observacao = f"{' | '.join(partes_observacao)} (BD)" if partes_observacao else None

    if not linhas and not observacao:
        return None
    return montar_tabela(nome_arquivo, "Consistência entre Documentos", None, linhas, observacao)

# ------- Documento esperado que não gerou conferência -------

def _bloco_ausente(nome_arquivo, nome_documento, dados_parecer, motivo=None):
    # placeholder pra um dos documentos esperados que não gerou bloco de conferência - aparece na
    # MESMA posição em que apareceria se tivesse sido encontrado, só que sem a tabela de campos: no
    # lugar dela, o motivo em vermelho (ver .observacao--solta no HTML).
    # "não detectado" abrange tanto a ausência real do documento no processo quanto uma falha na
    # detecção do título (OCR ruim, grafia diferente da esperada) - não dá pra distinguir aqui.
    # A Nota Fiscal (de serviço) só cai aqui quando não há PARECER E não é contrato de serviço sem
    # mão de obra (ex: contrato não identificado no BD) - contrato sem mão de obra gera bloco real
    # de NF, ver processar_nota_fiscal. `motivo` explícito pula essa regra (usado pelo almoxarifado).
    if motivo is None:
        if nome_documento == "Nota Fiscal" and not dados_parecer:
            motivo = "NF não conferida — processo sem PARECER do Gestor do Contrato (informativo)"
        else:
            motivo = "Documento não detectado no processo"
    return montar_tabela(nome_arquivo, nome_documento, None, [], motivo)

# ======= CONTRATO DE ALMOXARIFADO =======
# conjunto de documentos diferente do de serviço (sem IMR / Termo Circunstanciado / PARECER do
# Gestor). Implementado documento a documento, conforme definido com o usuário.

def _mesmos_digitos(a, b):
    # compara só os dígitos, ignorando pontuação (".", "/", "-", espaços) - usado onde a mesma
    # informação vem com formatação diferente entre o BD e o documento: nº de processo (BD
    # "23323.001485.2024-97" x doc ".../..."), agência/conta bancária ("69453-3" x "69.453-3"), etc.
    return bool(a) and bool(b) and re.sub(r"\D", "", a) == re.sub(r"\D", "", b)

_CNPJ_IFF_RAIZ = "10779511"  # raiz do CNPJ do IFFluminense (contratante) - só pra separar do fornecedor

def _cnpj_fornecedor(cnpjs):
    # entre os CNPJs achados num documento, o do fornecedor é o que NÃO é o do IFF (contratante).
    # NÃO escolhe "o que bate com o contrato" de propósito: um CNPJ de fornecedor ERRADO tem que
    # ser pego e comparado (→ ❌), não ignorado silenciosamente.
    return next((c for c in cnpjs if not re.sub(r"\D", "", c).startswith(_CNPJ_IFF_RAIZ)),
                cnpjs[0] if cnpjs else None)

def _linha_empenho(doc_empenhos, contrato, dados_of, rotulo_ausente="não citado no documento"):
    # empenho citado num documento do processo de almoxarifado (NF, encaminhamento, termo de
    # recebimento, IC...): QUANDO presente, tem que constar no BD E ser o mesmo da OF. Sem empenho
    # citado -> ➖ (não é falha). Devolve o dict pronto pro linha_tabela.
    empenhos_bd = empenhos_registrados(contrato) if contrato else []
    empenhos_of = list((dados_of or {}).get("empenhos") or [])
    if empenhos_of:
        fonte_texto, fonte_disp = f"{', '.join(empenhos_of)} (OF pág. {dados_of['pagina']}; BD)", True
    elif empenhos_bd:
        fonte_texto, fonte_disp = f"{', '.join(empenhos_bd)} (BD)", True
    else:
        fonte_texto, fonte_disp = ("OF não localizada no processo" if contrato else "contrato não encontrado no banco"), False

    if not doc_empenhos:
        bate = None
    elif empenhos_of:
        bate = (set(doc_empenhos) == set(empenhos_of)
                and bool(empenhos_bd) and all(e in empenhos_bd for e in doc_empenhos))
    elif empenhos_bd:
        bate = all(e in empenhos_bd for e in doc_empenhos)  # sem OF localizada, ao menos confere o BD
    else:
        bate = None
    return linha_tabela("Empenho", fonte_texto, fonte_disp,
                        ", ".join(doc_empenhos) if doc_empenhos else rotulo_ausente, bool(doc_empenhos), bate)

# ------- Almoxarifado / Documento 1: Ordem de Serviço / Fornecimento -------

RE_OF_NUMERO = re.compile(r"Ordem de Servi[çc]o\s*/\s*Fornecimento\s*n[ºo°]\s*(\d+/\d+)", re.IGNORECASE)
# a OF lista Fornecedor E Contratante no mesmo formato "CNPJ - NOME" (nome pode quebrar em 2 linhas)
RE_OF_ENTIDADE = re.compile(
    r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\s*-\s*(.+?)(?=\n(?:Contratante:|Amparo Legal:|\d{2}/\d{2}/\d{4}))",
    re.DOTALL)
RE_OF_VIGENCIA = re.compile(r"Vig[êe]ncia Inicial:.*?(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})", re.DOTALL)
RE_OF_PROCESSO_CONTRATACAO = re.compile(r"\d{5}\.\d{6}/\d{4}-\d{2}")
RE_OF_OBJETO = re.compile(
    r"Objeto:\s*\n(.+?)\n(?:Contrato n[ºo°]|2 - INFORMA|Powered by)", re.DOTALL | re.IGNORECASE)
RE_OF_EMPENHOS_BLOCO = re.compile(
    r"Empenhos:\s*(.*?)\n\s*(?:Locais de Execu|3 - ITENS)", re.DOTALL | re.IGNORECASE)
RE_OF_NUMERO_NE = re.compile(r"\d{4}NE\d{6}")
RE_OF_VALOR_TOTAL = re.compile(r"Valor Total da presente Ordem.*?R\$\s*([\d.,]+)", re.DOTALL | re.IGNORECASE)
RE_OF_EXECUCAO = re.compile(
    r"Data de assinatura:.*?(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})",
    re.DOTALL | re.IGNORECASE)
RE_OF_SECAO_ITENS = re.compile(
    r"3 - ITENS DA AUTORIZA[ÇC][ÃA]O DE EXECU[ÇC][ÃA]O\s*(.*?)(?:\n\s*O Valor Total da presente|\n\s*4 - INFORMA|\Z)",
    re.DOTALL | re.IGNORECASE)
# item da OF: "Material <num> <desc+unidade, pode quebrar em várias linhas> <qtd> <parcela>
# <qtd.solic> R$ <unit> \n R$ \n <total>". DOTALL pra descrição multi-linha ("POLPA\nDE\nFRUTA").
RE_OF_ITEM = re.compile(
    r"(?:Material|Servi[çc]o)\s+(\d{3,6})\s+(.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+"
    r"R\$\s*([\d.,]+)\s*\n\s*R\$\s*\n?\s*([\d.,]+)", re.IGNORECASE | re.DOTALL)

def _split_desc_unidade(bruto):
    # "POLPA DE FRUTA QUILOGRAMA" -> ("POLPA DE FRUTA", "QUILOGRAMA"); "PÃO" -> ("PÃO", "")
    d = limpar_espacos(bruto)
    desc, _, unid = d.rpartition(" ")
    return (desc, unid) if desc else (d, "")

def _localizar_ordem_fornecimento(paginas):
    # a OF ocupa 3 páginas no PDF ("1/3".."3/3"): "1 - INFORMAÇÕES DO CONTRATO", "2 - INFORMAÇÕES DA
    # ORDEM DE SERVIÇO" (+ itens + valor total) e complementares/autorização. Devolve (indice_da_1ª,
    # texto_das_3_juntas) ou (None, None). Última ocorrência, caso a OF tenha sido refeita.
    indices = [i for i, t in enumerate(paginas)
               if "1 - INFORMAÇÕES DO CONTRATO" in t and "Ordem de Serviço / Fornecimento" in t]
    if not indices:
        return None, None
    indice = indices[-1]
    return indice, remover_duplicatas_consecutivas("\n".join(paginas[indice:indice + 3]))

def obter_dados_of(paginas):
    # dados da OF que NÃO têm fonte segura no BD e que os próximos documentos do processo de
    # almoxarifado (NF, Termo de Recebimento, IC) vão conferir contra: nº da OF, valor total,
    # período de execução e a lista de itens autorizados (não exibida no bloco da OF, guardada aqui
    # pra uso posterior). None se a OF não está no processo.
    indice, texto = _localizar_ordem_fornecimento(paginas)
    if texto is None:
        return None

    m_num = RE_OF_NUMERO.search(texto)
    m_valor = RE_OF_VALOR_TOTAL.search(texto)
    m_exec = RE_OF_EXECUCAO.search(texto)
    m_secao = RE_OF_SECAO_ITENS.search(texto)
    itens_texto = m_secao.group(1).strip() if m_secao else ""

    m_bloco_emp = RE_OF_EMPENHOS_BLOCO.search(texto)
    empenhos = []
    for ne in (RE_OF_NUMERO_NE.findall(m_bloco_emp.group(1)) if m_bloco_emp else []):
        if ne not in empenhos:
            empenhos.append(ne)

    itens = []
    for pedaco in re.split(r"(?=(?:Material|Servi[çc]o)\s+\d{3,6}\b)", itens_texto):
        m = RE_OF_ITEM.match(pedaco)
        if not m:
            continue
        desc, unid = _split_desc_unidade(m.group(2))
        itens.append({
            "num_item": m.group(1), "descricao": desc, "unidade": unid,
            "quantidade": m.group(3), "parcela": m.group(4), "quant_solicitada": m.group(5),
            "valor_unitario": m.group(6), "valor_total": m.group(7),
        })

    return {
        "pagina": indice + 1,
        "numero": m_num.group(1) if m_num else None,
        "valor_total": m_valor.group(1) if m_valor else None,
        "assinatura": m_exec.group(1) if m_exec else None,
        "execucao_inicio": m_exec.group(2) if m_exec else None,
        "execucao_fim": m_exec.group(3) if m_exec else None,
        "empenhos": empenhos,
        "itens_texto": itens_texto,   # seção 3 crua, fallback caso o parse de linha falhe num modelo diferente
        "itens": itens,
    }

def _linhas_cabecalho_contrato(texto, contrato):
    # "INFORMAÇÕES DO CONTRATO" no padrão de 2 colunas (rótulos numa linha, valores interleaved) -
    # aparece igual na Ordem de Serviço/Fornecimento e no Termo de Recebimento Definitivo. Devolve
    # 6 linhas contra o BD: Contrato, Fornecedor, CNPJ, Objeto, Vigência, Processo de contratação.
    linhas = []

    m_contrato = RE_CONTRATO_GENERICO.search(texto)
    doc_contrato = m_contrato.group(1) if m_contrato else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    # entre Fornecedor e Contratante (mesmo formato "CNPJ - NOME"), pega o que NÃO é o IFF - assim
    # um CNPJ de fornecedor errado ainda é pego e comparado (→ ❌), não some
    entidades = [(c, limpar_espacos(n)) for c, n in RE_OF_ENTIDADE.findall(texto)]
    fornecedor = next(((c, n) for c, n in entidades if not re.sub(r"\D", "", c).startswith(_CNPJ_IFF_RAIZ)),
                      entidades[0] if entidades else None)
    doc_cnpj = fornecedor[0] if fornecedor else None
    doc_fornecedor = fornecedor[1] if fornecedor else None

    fonte_fornecedor = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Fornecedor", f"{fonte_fornecedor} (BD)" if fonte_fornecedor else "contrato não encontrado no banco", bool(fonte_fornecedor),
        doc_fornecedor or "não encontrado no documento", bool(doc_fornecedor),
        comparar_textos(fonte_fornecedor, doc_fornecedor) if fonte_fornecedor and doc_fornecedor else None,
    ))

    fonte_cnpj_fmt = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_objeto = RE_OF_OBJETO.search(texto)
    doc_objeto = limpar_espacos(m_objeto.group(1)) if m_objeto else None
    fonte_obj_texto, fonte_obj_disp, bate_obj = _conferir_objeto(contrato, doc_objeto)
    linhas.append(linha_tabela(
        "Objeto", fonte_obj_texto, fonte_obj_disp,
        doc_objeto or "não encontrado no documento", bool(doc_objeto),
        bate_obj,
    ))

    fonte_vigencia = None
    if contrato and contrato.get("vigencia_inicio") and contrato.get("vigencia_fim"):
        fonte_vigencia = f"{_formatar_data_iso(contrato['vigencia_inicio'])} a {_formatar_data_iso(contrato['vigencia_fim'])}"
    m_vig = RE_OF_VIGENCIA.search(texto)
    doc_vigencia = f"{m_vig.group(1)} a {m_vig.group(2)}" if m_vig else None
    linhas.append(linha_tabela(
        "Vigência", f"{fonte_vigencia} (BD)" if fonte_vigencia else "contrato não encontrado no banco", bool(fonte_vigencia),
        doc_vigencia or "não encontrada no documento", bool(doc_vigencia),
        (fonte_vigencia == doc_vigencia) if fonte_vigencia and doc_vigencia else None,
    ))

    m_proc = RE_OF_PROCESSO_CONTRATACAO.search(texto)
    doc_proc = m_proc.group() if m_proc else None
    fonte_proc = contrato["processo_contratacao"] if contrato else None
    linhas.append(linha_tabela(
        "Processo de contratação", f"{fonte_proc} (BD)" if fonte_proc else "contrato não encontrado no banco", bool(fonte_proc),
        doc_proc or "não encontrado no documento", bool(doc_proc),
        _mesmos_digitos(fonte_proc, doc_proc) if fonte_proc and doc_proc else None,
    ))
    return linhas

def processar_ordem_fornecimento(nome_arquivo, paginas, contrato, dados_of):
    # Documento 1 do processo de pagamento de almoxarifado. Os itens da OF NÃO entram na tabela
    # deste bloco (ficam em dados_of["itens"], pra conferir nos documentos seguintes) - aqui só a
    # conferência do cabeçalho contra o BD + observação com nº da OF / valor total / execução.
    indice, texto = _localizar_ordem_fornecimento(paginas)
    if texto is None:
        return None
    dados_of = dados_of or {}

    linhas = _linhas_cabecalho_contrato(texto, contrato)

    # fonte segura = só o(s) empenho(s) que a OF está usando, confirmado(s) no BD - não a lista
    # inteira de empenhos do contrato
    doc_empenhos = list(dados_of.get("empenhos") or [])
    empenhos_bd = empenhos_registrados(contrato) if contrato else []
    validos_bd = [e for e in doc_empenhos if e in empenhos_bd]
    if validos_bd:
        fonte_emp_texto, fonte_emp_disp = f"{', '.join(validos_bd)} (BD)", True
    elif empenhos_bd:
        fonte_emp_texto, fonte_emp_disp = f"{', '.join(empenhos_bd)} (BD)", True
    else:
        fonte_emp_texto, fonte_emp_disp = ("nenhum empenho cadastrado nesse contrato" if contrato else "contrato não encontrado no banco"), False
    bate_emp = all(e in empenhos_bd for e in doc_empenhos) if doc_empenhos and empenhos_bd else None
    linhas.append(linha_tabela(
        "Empenhos", fonte_emp_texto, fonte_emp_disp,
        ", ".join(doc_empenhos) if doc_empenhos else "não encontrado no documento", bool(doc_empenhos),
        bate_emp,
    ))

    # Nº da OF e Valor Total não têm fonte segura no BD (são determinados pela própria OF) - viram
    # observação e são a referência dos próximos documentos do processo (NF, Termo de Recebimento,
    # IC). Os itens (dados_of["itens"]) ficam guardados, não entram aqui.
    observacao = (f"Nº da OF: {dados_of.get('numero') or 'não encontrado'} | "
                  f"Valor total: {dados_of.get('valor_total') or 'não encontrado'}")

    return montar_tabela(nome_arquivo, "Ordem de Serviço / Fornecimento", indice + 1, linhas, observacao)

# ------- Almoxarifado / Documento 2: Nota Fiscal (DANFE - NF-e de material) -------

# a NF de almoxarifado é uma DANFE (venda de mercadoria), layout totalmente diferente da NFS-e de
# serviço - regexes dedicados. Página identificada por "DANFE" + "CHAVE DE ACESSO".
RE_DANFE_NUMERO = re.compile(r"N[ºo°]\.\s*(\d[\d.]+\d)")
RE_DANFE_CONTRATO = re.compile(r"CONTRATO\s*n?[ºo°]?\s*([\d./]+)", re.IGNORECASE)
RE_DANFE_OF = re.compile(r"ORDEM DE FORNECIMENTO\s*n?[ºo°]?\s*([\d./]+)", re.IGNORECASE)
RE_DANFE_EMPENHO = re.compile(r"\d{4}NE\d{6}")
RE_DANFE_VALOR_NOTA = re.compile(r"V\.?\s*TOTAL DA NOTA\s*\n?\s*([\d.,]+)", re.IGNORECASE)
RE_DANFE_EMISSAO = re.compile(r"DATA DA EMISS[ÃA]O\s*\n?\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_DANFE_BANCO = re.compile(
    r"BANCO\s+(?:([A-Za-zÀ-Úà-ú]+)\s+)?(\d{3,4})\s+AG\.?\s*(\d+)\s+C/C\s*([\d.]+(?:-\d+)?)", re.IGNORECASE)
# item da DANFE, dentro da seção "DADOS DOS PRODUTOS / SERVIÇOS": CÓDIGO DESCRIÇÃO(pode quebrar de
# linha) NCM(6-8díg) O/CSOSN CFOP UN QUANT VALOR_UNIT VALOR_TOTAL 0,00 ... - código pode ter 1
# dígito ("7"); UN até ~12 letras ("LITRO"). A NCM (6-8 dígitos seguidos) é a âncora do fim da descrição.
RE_DANFE_ITEM = re.compile(
    r"\s*(\d{1,6})\s+(.+?)\s+\d{6,8}\s+\d{2,4}\s+\d{4}\s+([A-Za-zÀ-Úà-ú]{1,12})\s+"
    r"([\d.]+,\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d{2})\b", re.DOTALL)

def _localizar_danfe(paginas):
    indices = [i for i, t in enumerate(paginas) if "DANFE" in t and "CHAVE DE ACESSO" in t]
    if not indices:
        return None, None
    indice = indices[-1]
    return indice, remover_duplicatas_consecutivas(paginas[indice])

def _num_nf(bruto):
    # "000.001.019" -> "1019"
    if not bruto:
        return None
    digitos = re.sub(r"\D", "", bruto).lstrip("0")
    return digitos or "0"

def _fmt_item(item):
    prefixo = f"[{item['num_item']}] " if item.get("num_item") else ""
    return f"{prefixo}{item['descricao']}: {item['quantidade']} × {item['valor_unitario']} = {item['valor_total']}"

def _cruzar_por_valor_unitario(itens_of, itens_nf, itens_trd):
    # alinha os itens dos 3 docs pelo valor unitário (preço contratual, igual nos 3) - devolve
    # [(chave, item_of|None, item_nf|None, item_trd|None), ...] na ordem em que aparecem
    def chave(it):
        v = _valor_para_float((it or {}).get("valor_unitario"))
        return round(v, 2) if v is not None else None
    chaves = []
    for grupo in (itens_of or [], itens_nf or [], itens_trd or []):
        for it in grupo:
            k = chave(it)
            if k is not None and k not in chaves:
                chaves.append(k)
    return [(k,
             next((i for i in (itens_of or []) if chave(i) == k), None),
             next((i for i in (itens_nf or []) if chave(i) == k), None),
             next((i for i in (itens_trd or []) if chave(i) == k), None))
            for k in chaves]

def _somar_itens(itens):
    return _float_para_valor_br(sum(_valor_para_float(i["valor_total"]) or 0 for i in (itens or [])))

def _bloco_cruzamento_itens(nome_arquivo, dados_of, dados_nf, dados_trd, pendentes=None):
    # cruza os itens dos documentos de almoxarifado:
    #   LIMITE = quantidade que ainda pode ser paga - a OF inteira na 1ª NF, ou o que ficou
    #            PENDENTE (registrado na Observação do contrato) quando a mesma OF volta num
    #            processo novo. Entrega/pagamento parcial é permitido: NF <= LIMITE.
    #   NF     = o que foi efetivamente entregue/faturado agora
    #   TRD    = recebimento definitivo (tem que ser EXATO com a NF)
    itens_nf  = (dados_nf  or {}).get("itens") or []
    itens_trd = (dados_trd or {}).get("itens") or []
    itens_lim = list(pendentes) if pendentes else ((dados_of or {}).get("itens") or [])
    rotulo_lim = "Pendente" if pendentes else "OF"
    tot_lim = _somar_itens(pendentes) if pendentes else (dados_of or {}).get("valor_total")
    if not (itens_lim or itens_nf or itens_trd):
        return None

    cruzados = _cruzar_por_valor_unitario(itens_lim, itens_nf, itens_trd)
    linhas = []
    for k, lim, nf, trd in cruzados:
        rotulo = ((lim or trd or {}).get("num_item")
                  or (limpar_espacos((nf or {}).get("descricao", "")) or f"R$ {k:.2f}/un")[:44])
        col_doc = []
        if lim:
            col_doc.append(f"{rotulo_lim}: {lim['quantidade']} × {lim['valor_unitario']} = {lim['valor_total']}")
        if trd:
            col_doc.append(f"Receb.: {trd['quantidade']} × {trd['valor_unitario']} = {trd['valor_total']}")
        fonte_texto = (f"NF: {nf['quantidade']} × {nf['valor_unitario']} = {nf['valor_total']}"
                       if nf else f"não consta na NF (segue pendente)" if pendentes else "não consta na NF (item não entregue)")

        problemas = []
        if nf and not lim:
            problemas.append(f"faturado item fora do {'pendente' if pendentes else 'autorizado na OF'}")
        if nf and not trd:
            problemas.append("entregue mas fora do recebimento")
        if trd and not nf:
            problemas.append("recebido item não faturado")
        if nf and trd and not (_valores_monetarios_batem(nf["quantidade"], trd["quantidade"])
                               and _valores_monetarios_batem(nf["valor_total"], trd["valor_total"])):
            problemas.append("NF ≠ recebimento (têm que ser exatos)")
        if nf and lim and not (_valor_ate(nf["quantidade"], lim["quantidade"])
                               and _valor_ate(nf["valor_total"], lim["valor_total"])):
            problemas.append(f"NF acima do {'pendente' if pendentes else 'autorizado na OF'}")

        bate = None if (nf is None and trd is None) else (not problemas)
        col_texto = " | ".join(col_doc) if col_doc else "-"
        if problemas:
            col_texto += f"   ⚠ {'; '.join(problemas)}"
        linhas.append(linha_tabela(f"Item {rotulo}", fonte_texto, bool(nf), col_texto, bool(col_doc), bate))

    tot_nf  = (dados_nf  or {}).get("valor_total")
    tot_trd = (dados_trd or {}).get("valor_total")
    prob_tot = []
    if tot_nf and tot_trd and not _valores_monetarios_batem(tot_nf, tot_trd):
        prob_tot.append("NF ≠ recebimento")
    if tot_nf and tot_lim and not _valor_ate(tot_nf, tot_lim):
        prob_tot.append(f"NF acima do {'pendente' if pendentes else 'autorizado'}")
    col_tot = " | ".join(p for p in (f"{rotulo_lim}: {tot_lim}" if tot_lim else None,
                                     f"Receb.: {tot_trd}" if tot_trd else None) if p)
    if prob_tot:
        col_tot += f"   ⚠ {'; '.join(prob_tot)}"
    linha_total = linha_tabela(
        "Total", f"NF: {tot_nf}" if tot_nf else "NF não localizada", bool(tot_nf),
        col_tot or "-", bool(tot_lim or tot_trd),
        (not prob_tot) if (tot_nf and (tot_lim or tot_trd)) else None,
    )
    linha_total["destaque"] = True  # linha dos totais - realçada (negrito) na janela
    linhas.append(linha_total)

    # reconciliação: a soma dos itens NÃO pagos nesta NF tem que fechar com (LIMITE total − NF paga)
    nao_pagos = [lim for _, lim, nf, _ in cruzados if lim and not nf]
    f_lim, f_nf = _valor_para_float(tot_lim), _valor_para_float(tot_nf)
    if nao_pagos and f_lim is not None and f_nf is not None:
        soma = sum(_valor_para_float(o["valor_total"]) or 0 for o in nao_pagos)
        dif = f_lim - f_nf
        rot = "Itens ainda pendentes" if pendentes else "Itens não entregues"
        linhas.append(linha_tabela(
            rot,
            f"{tot_lim} − {tot_nf} = {_float_para_valor_br(dif)} ({rotulo_lim} − NF paga)", True,
            f"{' + '.join(o['valor_total'] for o in nao_pagos)} = {_float_para_valor_br(soma)}", True,
            abs(soma - dif) < 0.005,
        ))

    return montar_tabela(nome_arquivo, "Cruzamento de Itens (OF × NF × Recebimento)", None, linhas)

def _num_curto(numero):
    # "00004/2026" -> "4/2026" ; "00049" -> "49"
    return "/".join(re.sub(r"^0+(?=\d)", "", p) for p in (numero or "").split("/"))

# linha "Itens não entregues (OF 4/2026) --> 49) 100x8,40 + 46) 100x9,40 = 940,00" na Observação
RE_OBS_NAO_ENTREGUES = re.compile(
    r"^Itens (?:não entregues|ainda pendentes) \(OF ([\d/]+)\)\s*-->\s*(.+?)\s*=\s*[\d.,]+\s*$", re.MULTILINE)
RE_OBS_ITEM = re.compile(r"(\d+)\)\s*([\d.,]+)\s*x\s*([\d.,]+)")

def _parse_itens_pendentes(observacao, of_num_curto):
    # lê de volta a linha "Itens não entregues (OF X) --> ..." da Observação -> lista de itens
    # {num_item, quantidade, valor_unitario, valor_total}. None se não há linha pra essa OF.
    for m in RE_OBS_NAO_ENTREGUES.finditer(observacao or ""):
        if m.group(1) != of_num_curto:
            continue
        itens = []
        for ni, q, vu in RE_OBS_ITEM.findall(m.group(2)):
            fq, fvu = _valor_para_float(q) or 0, _valor_para_float(vu) or 0
            itens.append({"num_item": ni, "quantidade": q, "valor_unitario": vu,
                          "valor_total": _float_para_valor_br(fq * fvu)})
        return itens or None
    return None

def _registrar_itens_nao_entregues(contrato, dados_of, dados_nf, pendentes=None):
    # entrega parcial de almoxarifado: mantém na Observação do contrato a linha "Itens não entregues
    # (OF X) --> ..." com o que AINDA falta pagar daquela OF. 1ª NF: OF inteira menos o que a NF
    # pagou. NF seguinte da mesma OF: o que estava pendente menos o que esta NF pagou. Quando não
    # sobra nada, a linha é removida. Só mexe quando a reconciliação fecha (soma == LIMITE − NF).
    if not (contrato and contrato.get("id") and dados_of and dados_nf):
        return
    itens_lim = list(pendentes) if pendentes else (dados_of.get("itens") or [])
    itens_nf = dados_nf.get("itens") or []
    if not (itens_lim and itens_nf):
        return
    cruzados = _cruzar_por_valor_unitario(itens_lim, itens_nf, [])
    faltam = [lim for _, lim, nf, _ in cruzados if lim and not nf]  # itens não pagos nesta NF

    tot_lim = _somar_itens(pendentes) if pendentes else dados_of.get("valor_total")
    f_lim, f_nf = _valor_para_float(tot_lim), _valor_para_float(dados_nf.get("valor_total"))
    soma = sum(_valor_para_float(o["valor_total"]) or 0 for o in faltam)
    if f_lim is None or f_nf is None or abs(soma - (f_lim - f_nf)) >= 0.005:
        return  # reconciliação não fechou - não mexe no BD

    of_num = _num_curto(dados_of.get("numero"))
    prefixo = f"Itens não entregues (OF {of_num})"
    if faltam:
        partes = " + ".join(f"{_num_curto(o['num_item'])}) {o['quantidade']}x{o['valor_unitario']}" for o in faltam)
        contratos_db.acrescentar_observacao_linha(
            contrato["id"], prefixo, f"{prefixo} --> {partes} = {_float_para_valor_br(soma)}")
    elif pendentes:  # zerou a pendência daquela OF - tira a linha
        contratos_db.acrescentar_observacao_linha(contrato["id"], prefixo, None)

# ------- Almoxarifado: Consistência entre Documentos -------
# mesma info conferida em documentos distintos do processo tem que ser igual entre si (não só bater
# isoladamente contra a fonte segura). Varre os blocos já montados.
_CAMPOS_CONSISTENCIA_ALMOX = {
    "CNPJ": ["CNPJ"],
    "Contrato": ["Contrato"],
    "Fornecedor": ["Fornecedor", "Interessado"],
    "Objeto": ["Objeto"],
    "Vigência": ["Vigência"],
    "Processo de contratação": ["Processo de contratação"],
    "Processo": ["Processo"],                       # do pagamento (pág. 1)
    "Nota Fiscal": ["Nota Fiscal"],
    "Ordem de Fornecimento": ["Ordem de Fornecimento"],
    "Empenho": ["Empenho", "Empenhos"],
    "Valor": ["Valor", "Valor Faturado", "Valor Líquido"],
}
_COMPARADOR_CONSISTENCIA_ALMOX = {
    "CNPJ": comparar_cnpjs,
    "Contrato": comparar_numeros,
    "Fornecedor": comparar_textos,
    "Objeto": comparar_textos,
    "Vigência": lambda a, b: a == b,
    "Processo de contratação": _mesmos_digitos,
    "Processo": _mesmos_digitos,
    "Nota Fiscal": comparar_numeros,
    "Ordem de Fornecimento": comparar_numeros,
    "Empenho": _comparar_conjuntos,
    "Valor": _valores_monetarios_batem,
}
_NOMES_CURTOS_DOC_ALMOX = {
    "Ordem de Serviço / Fornecimento": "OF",
    "Nota Fiscal": "NF",
    "Encaminhamento de Material": "Encaminhamento",
    "Despacho de Ateste de Nota Fiscal de Material": "Despacho Ateste",
    "Capa de Pagamento": "Capa",
    "Termo de Recebimento Definitivo": "Recebimento",
    "Instrumentos de Cobrança": "IC",
    "Consulta Optante pelo Simples Nacional": "Consulta Optante",
}

def _bloco_consistencia_almoxarifado(nome_arquivo, tabelas, contrato, dados_of, dados_nf, processo_p1):
    dados_of, dados_nf = dados_of or {}, dados_nf or {}
    fonte_cnpj = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    fonte_vig = None
    if contrato and contrato.get("vigencia_inicio") and contrato.get("vigencia_fim"):
        fonte_vig = f"{_formatar_data_iso(contrato['vigencia_inicio'])} a {_formatar_data_iso(contrato['vigencia_fim'])}"
    fonte_obj = (contrato.get("objeto_detalhado") or contrato.get("objeto_resumido")) if contrato else None
    empenhos_of = ", ".join(dados_of.get("empenhos") or []) or None

    fontes = {
        "CNPJ": (f"{fonte_cnpj} (BD)" if fonte_cnpj else "—", bool(fonte_cnpj), fonte_cnpj),
        "Contrato": (f"{contrato['numero_contrato']} (BD)" if contrato and contrato.get("numero_contrato") else "—",
                     bool(contrato and contrato.get("numero_contrato")), contrato.get("numero_contrato") if contrato else None),
        "Fornecedor": (f"{contrato['nome_contratada']} (BD)" if contrato and contrato.get("nome_contratada") else "—",
                       bool(contrato and contrato.get("nome_contratada")), contrato.get("nome_contratada") if contrato else None),
        "Objeto": (f"{fonte_obj} (BD)" if fonte_obj else "—", bool(fonte_obj), fonte_obj),
        "Vigência": (f"{fonte_vig} (BD)" if fonte_vig else "—", bool(fonte_vig), fonte_vig),
        "Processo de contratação": (f"{contrato['processo_contratacao']} (BD)" if contrato and contrato.get("processo_contratacao") else "—",
                                    bool(contrato and contrato.get("processo_contratacao")),
                                    contrato.get("processo_contratacao") if contrato else None),
        "Processo": (f"{processo_p1} (pág. 1)" if processo_p1 else "—", bool(processo_p1), processo_p1),
        "Nota Fiscal": (f"{dados_nf.get('numero')} (NF)" if dados_nf.get("numero") else "—",
                        bool(dados_nf.get("numero")), dados_nf.get("numero")),
        "Ordem de Fornecimento": (f"{dados_of.get('numero')} (OF)" if dados_of.get("numero") else "—",
                                  bool(dados_of.get("numero")), dados_of.get("numero")),
        "Empenho": (f"{empenhos_of} (OF)" if empenhos_of else "—", bool(empenhos_of), empenhos_of),
        "Valor": (f"{dados_nf.get('valor_total')} (NF)" if dados_nf.get("valor_total") else "—",
                  bool(dados_nf.get("valor_total")), dados_nf.get("valor_total")),
    }

    linhas = []
    for campo, rotulos in _CAMPOS_CONSISTENCIA_ALMOX.items():
        ocorrencias = []
        for bloco in tabelas:
            nome_curto = _NOMES_CURTOS_DOC_ALMOX.get(bloco["documento"], bloco["documento"])
            for linha in bloco["linhas"]:
                if linha["campo"] in rotulos and linha["documento_disponivel"]:
                    ocorrencias.append((nome_curto, _valor_comparavel(linha["documento"])))
        if not ocorrencias:
            continue
        fonte_texto, fonte_disp, fonte_valor = fontes[campo]
        comparador = _COMPARADOR_CONSISTENCIA_ALMOX[campo]
        referencia = fonte_valor or (ocorrencias[0][1] if len(ocorrencias) >= 2 else None)
        if referencia is None:
            bate, doc_texto = None, ocorrencias[0][1]
        else:
            divergentes = [(n, v) for n, v in ocorrencias if not comparador(referencia, v)]
            bate = not divergentes
            # quando bate, só o valor (já detalhado nos blocos acima); quando não, só o(s)
            # documento(s) onde está errado
            doc_texto = ocorrencias[0][1] if bate else " | ".join(f"{n}: {v}" for n, v in divergentes)
        linhas.append(linha_tabela(campo, fonte_texto, fonte_disp, doc_texto, True, bate))

    if not linhas:
        return None
    return montar_tabela(nome_arquivo, "Consistência entre Documentos", None, linhas)

def obter_dados_nf_almoxarifado(paginas):
    # dados da DANFE que os documentos seguintes (Instrumentos de Cobrança, SIAFI NS) conferem
    # contra: nº da NF, valor total, emissão e itens. None se não há DANFE no processo.
    indice, texto = _localizar_danfe(paginas)
    if texto is None:
        return None
    m_num = RE_DANFE_NUMERO.search(texto)
    m_valor = RE_DANFE_VALOR_NOTA.search(texto)
    m_emissao = RE_DANFE_EMISSAO.search(texto)
    # só a seção "DADOS DOS PRODUTOS" (senão o "0 - ENTRADA / 1 - SAÍDA" do topo da DANFE entra
    # como item). Cada item começa numa linha "<código> <letra...>" (a linha da NCM começa com
    # 6-8 dígitos + espaço + dígito, então não é confundida).
    m_prod = re.search(
        r"DADOS DOS PRODUTOS\s*/?\s*SERVI[ÇC]OS(.*?)(?:DADOS ADICIONAIS|INFORMA[ÇC][ÕO]ES COMPLEMENTARES|RESERVADO AO FISCO|\Z)",
        texto, re.DOTALL | re.IGNORECASE)
    secao = m_prod.group(1) if m_prod else texto
    itens = []
    for chunk in re.split(r"(?=\n\d{1,6}\s+\D)", "\n" + secao):
        m = RE_DANFE_ITEM.match(chunk)
        if m:
            itens.append({
                "codigo": m.group(1),
                "descricao": limpar_espacos(m.group(2)),
                "unidade": m.group(3),
                "quantidade": m.group(4),
                "valor_unitario": m.group(5),
                "valor_total": m.group(6),
            })
    return {
        "pagina": indice + 1,
        "numero": _num_nf(m_num.group(1)) if m_num else None,
        "valor_total": m_valor.group(1) if m_valor else None,
        "emissao": m_emissao.group(1) if m_emissao else None,
        "itens": itens,
    }

def processar_nota_fiscal_almoxarifado(nome_arquivo, paginas, contrato, dados_of, dados_nf):
    indice, texto = _localizar_danfe(paginas)
    if texto is None:
        return None
    dados_of = dados_of or {}
    dados_nf = dados_nf or {}

    linhas = []

    # CNPJ do emitente = o que NÃO é o do IFF/destinatário (a DANFE traz os dois) - assim um CNPJ
    # de emitente errado ainda é comparado (→ ❌)
    doc_cnpj = _cnpj_fornecedor(re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto))
    fonte_cnpj_fmt = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado na NF", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_contrato = RE_DANFE_CONTRATO.search(texto)
    doc_contrato = m_contrato.group(1) if m_contrato else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    # Contrato e Ordem de Fornecimento ficam nas "Informações Complementares" da DANFE - nem toda
    # NF traz; sem eles é ➖ "não citado", não é falha
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não citado na NF", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    m_of = RE_DANFE_OF.search(texto)
    doc_of = m_of.group(1) if m_of else None
    fonte_of = dados_of.get("numero")
    linhas.append(linha_tabela(
        "Ordem de Fornecimento", f"{fonte_of} (OF pág. {dados_of['pagina']})" if fonte_of else "OF não localizada no processo", bool(fonte_of),
        doc_of or "não citada na NF", bool(doc_of),
        comparar_numeros(fonte_of, doc_of) if fonte_of and doc_of else None,
    ))

    doc_empenhos = []
    for ne in RE_DANFE_EMPENHO.findall(texto):
        if ne not in doc_empenhos:
            doc_empenhos.append(ne)
    linhas.append(_linha_empenho(doc_empenhos, contrato, dados_of, rotulo_ausente="não citado na NF"))

    m_banco = RE_DANFE_BANCO.search(texto)
    if m_banco:
        doc_ag, doc_conta = m_banco.group(3), m_banco.group(4)
        fonte_ag = contrato["agencia"] if contrato else None
        fonte_conta = contrato["conta"] if contrato else None
        tem_fonte = bool(fonte_ag and fonte_conta)
        fonte_texto = (f"{contrato.get('banco') or '-'} | Ag {fonte_ag} | C/c {fonte_conta} (BD)" if tem_fonte
                       else ("dados bancários não cadastrados no banco" if contrato else "contrato não encontrado no banco"))
        doc_texto = f"{m_banco.group(1) or '-'} {m_banco.group(2)} | Ag {doc_ag} | C/c {doc_conta}"
        bate_dom = (_mesmos_digitos(fonte_ag, doc_ag) and _mesmos_digitos(fonte_conta, doc_conta)) if tem_fonte else None
        linhas.append(linha_tabela("Domicílio Bancário", fonte_texto, tem_fonte, doc_texto, True, bate_dom))

    # itens e valor total da NF vão pro bloco "Cruzamento de Itens" (OF × NF × Recebimento), no fim
    observacao = (f"Nº da NF: {dados_nf.get('numero') or 'não encontrado'} | "
                  f"Emissão: {dados_nf.get('emissao') or 'não encontrada'} | "
                  f"Valor total: {dados_nf.get('valor_total') or 'não encontrado'}")

    return montar_tabela(nome_arquivo, "Nota Fiscal", indice + 1, linhas, observacao)

# ------- Almoxarifado / Documentos 3 e 4: Encaminhamento de Material e Despacho de Ateste -------

# os dois são ofícios/despachos curtos com a MESMA conferência (NF / empenho / CNPJ / fornecedor
# contra as fontes seguras) - só muda a página que os identifica. A fonte usa a ligatura "ﬁ"
# (U+FB01) em "nota ﬁscal" - normalizada pra "fi" antes dos regexes.
RE_ALMOX_NF = re.compile(r"Nota\s*Fiscal\s*n[ºo°]\s*(\d+)", re.IGNORECASE)
RE_ALMOX_NE = re.compile(r"\d{4}NE\d{6}")
RE_ALMOX_CNPJ_FORNECEDOR = re.compile(r"fornecedor\s*\n?\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.IGNORECASE)
RE_ALMOX_FORNECEDOR = re.compile(
    r"fornecedor\s*\n?\s*\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*\n?\s*-\s*\n?\s*(.+?)\s*\.?\s*\n",
    re.IGNORECASE | re.DOTALL)

def _linhas_ateste_almoxarifado(texto, contrato, dados_of, dados_nf):
    # 4 linhas comuns ao Encaminhamento de Material e ao Despacho de Ateste: NF, Empenho, CNPJ e
    # Fornecedor conferidos contra dados_nf / BD / OF.
    dados_nf = dados_nf or {}
    linhas = []

    m_nf = RE_ALMOX_NF.search(texto)
    doc_nf = m_nf.group(1) if m_nf else None
    fonte_nf = dados_nf.get("numero")
    linhas.append(linha_tabela(
        "Nota Fiscal", f"{fonte_nf} (NF pág. {dados_nf['pagina']})" if fonte_nf else "NF não localizada no processo", bool(fonte_nf),
        doc_nf or "não encontrada no documento", bool(doc_nf),
        comparar_numeros(fonte_nf, doc_nf) if fonte_nf and doc_nf else None,
    ))

    doc_empenhos = []
    for ne in RE_ALMOX_NE.findall(texto):
        if ne not in doc_empenhos:
            doc_empenhos.append(ne)
    linhas.append(_linha_empenho(doc_empenhos, contrato, dados_of))

    # CNPJ do fornecedor: o que vem logo depois de "fornecedor" - assim um CNPJ errado (ex:
    # "53.194.350" no lugar de "53.094.350") ainda é pego e comparado (→ ❌), não some
    m_cnpj = RE_ALMOX_CNPJ_FORNECEDOR.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else _cnpj_fornecedor(re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto))
    fonte_cnpj_fmt = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_forn = RE_ALMOX_FORNECEDOR.search(texto)
    doc_forn = limpar_espacos(m_forn.group(1)) if m_forn else None
    fonte_forn = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Fornecedor", f"{fonte_forn} (BD)" if fonte_forn else "contrato não encontrado no banco", bool(fonte_forn),
        doc_forn or "não encontrado no documento", bool(doc_forn),
        comparar_textos(fonte_forn, doc_forn) if fonte_forn and doc_forn else None,
    ))
    return linhas

def _texto_pagina_almox(paginas, indice):
    return remover_duplicatas_consecutivas(paginas[indice]).replace("ﬁ", "fi").replace("ﬂ", "fl")

def processar_encaminhamento_material(nome_arquivo, paginas, contrato, dados_of, dados_nf):
    indices = [i for i, t in enumerate(paginas) if "Encaminhamento de material" in t and "OFÍCIO" in t]
    if not indices:
        return None
    indice = indices[-1]
    linhas = _linhas_ateste_almoxarifado(_texto_pagina_almox(paginas, indice), contrato, dados_of, dados_nf)
    return montar_tabela(nome_arquivo, "Encaminhamento de Material", indice + 1, linhas)

def processar_despacho_ateste_material(nome_arquivo, paginas, contrato, dados_of, dados_nf):
    # "Assunto: Despacho de Ateste de Nota Fiscal de Material" também aparece no Termo de Recebimento
    # Definitivo (que referencia esse despacho) - exige "ATESTO" (o verbo do ato) pra pegar a página certa
    indices = [i for i, t in enumerate(paginas)
               if "Despacho de Ateste de Nota Fiscal de Material" in t and "ATESTO" in t]
    if not indices:
        return None
    indice = indices[-1]
    linhas = _linhas_ateste_almoxarifado(_texto_pagina_almox(paginas, indice), contrato, dados_of, dados_nf)
    return montar_tabela(nome_arquivo, "Despacho de Ateste de Nota Fiscal de Material", indice + 1, linhas)

# ------- Almoxarifado / Documento 5: Capa de Pagamento -------

RE_CP_PROCESSO = re.compile(r"processo de compra protocolado sob o n[úu]mero:\s*([\d./-]+)", re.IGNORECASE)
RE_CP_CNPJ = re.compile(r"CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.IGNORECASE)
RE_CP_INTERESSADO = re.compile(r"Interessado:\s*(.+)")
RE_CP_NF = re.compile(r"N[úu]mero Nota Fiscal:\s*(\d+)", re.IGNORECASE)
RE_CP_VALOR = re.compile(r"Valor:\s*([\d.,]+)")
RE_CP_ELEMENTO = re.compile(r"Elemento de Despesa\s+Valor\s*\n\s*(\d{6}[.\-]\d{2})", re.IGNORECASE)

def _natureza_do_empenho(contrato, numero_empenho):
    # natureza_despesa (ND) cadastrada pro empenho específico - "" (não informada) vira None
    if not (contrato and numero_empenho):
        return None
    todos = list(contrato.get("empenhos", []))
    for p in contrato.get("processos_empenho", []):
        todos.extend(p.get("empenhos", []))
    for e in todos:
        if e.get("numero_empenho") == numero_empenho:
            return (e.get("natureza_despesa") or "").strip() or None
    return None

def _processo_do_empenho(contrato, numero_empenho):
    # nº do processo de empenho (almoxarifado) que contém o empenho dado - None se não achar
    if not (contrato and numero_empenho):
        return None
    for p in contrato.get("processos_empenho", []):
        if any(e.get("numero_empenho") == numero_empenho for e in p.get("empenhos", [])):
            return p.get("numero_processo")
    return None

def processar_capa_pagamento(nome_arquivo, paginas, contrato, dados_of, dados_nf):
    # Documento 5: a "Capa de Pagamento" da Coordenação de Almoxarifado. p13 é a versão digitalizada
    # ("CAPA PAGAMENTO" em caixa alta) - a de conteúdo é a que tem "Capa de Pagamento" + a tabela
    # "Total por elemento de despesa".
    indices = [i for i, t in enumerate(paginas)
               if "Capa de Pagamento" in t and "Total por elemento de despesa" in t]
    if not indices:
        return None
    indice = indices[-1]
    texto = paginas[indice]
    dados_of = dados_of or {}
    dados_nf = dados_nf or {}

    doc_empenhos = []
    for ne in RE_ALMOX_NE.findall(texto):
        if ne not in doc_empenhos:
            doc_empenhos.append(ne)
    empenho_usado = doc_empenhos[0] if doc_empenhos else None

    linhas = []

    # fonte segura = o processo de empenho a que o empenho DESTE pagamento está vinculado no BD -
    # não a lista inteira de processos do contrato
    m_proc = RE_CP_PROCESSO.search(texto)
    doc_proc = m_proc.group(1) if m_proc else None
    proc_bd = _processo_do_empenho(contrato, empenho_usado)
    procs_bd = [p["numero_processo"] for p in contrato.get("processos_empenho", [])] if contrato else []
    if proc_bd:
        fonte_proc_texto, fonte_proc_disp = f"{proc_bd} (BD)", True  # o processo vinculado ao empenho usado
        bate_proc = _mesmos_digitos(proc_bd, doc_proc) if doc_proc else None
    else:
        fonte_proc_texto = f"{', '.join(procs_bd)} (BD)" if procs_bd else ("nenhum processo de empenho cadastrado" if contrato else "contrato não encontrado no banco")
        fonte_proc_disp = bool(procs_bd)
        bate_proc = any(_mesmos_digitos(p, doc_proc) for p in procs_bd) if procs_bd and doc_proc else None
    linhas.append(linha_tabela(
        "Processo de empenho", fonte_proc_texto, fonte_proc_disp,
        doc_proc or "não encontrado no documento", bool(doc_proc),
        bate_proc,
    ))

    m_cnpj = RE_CP_CNPJ.search(texto)
    doc_cnpj = m_cnpj.group(1) if m_cnpj else None
    fonte_cnpj_fmt = _formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else None
    linhas.append(linha_tabela(
        "CNPJ", f"{fonte_cnpj_fmt} (BD)" if fonte_cnpj_fmt else "contrato não encontrado no banco", bool(fonte_cnpj_fmt),
        doc_cnpj or "não encontrado no documento", bool(doc_cnpj),
        comparar_cnpjs(fonte_cnpj_fmt, doc_cnpj) if fonte_cnpj_fmt and doc_cnpj else None,
    ))

    m_int = RE_CP_INTERESSADO.search(texto)
    doc_int = limpar_espacos(m_int.group(1)) if m_int else None
    fonte_forn = contrato["nome_contratada"] if contrato else None
    linhas.append(linha_tabela(
        "Interessado", f"{fonte_forn} (BD)" if fonte_forn else "contrato não encontrado no banco", bool(fonte_forn),
        doc_int or "não encontrado no documento", bool(doc_int),
        comparar_textos(fonte_forn, doc_int) if fonte_forn and doc_int else None,
    ))

    m_nf = RE_CP_NF.search(texto)
    doc_nf = _num_nf(m_nf.group(1)) if m_nf else None
    fonte_nf = dados_nf.get("numero")
    linhas.append(linha_tabela(
        "Nota Fiscal", f"{fonte_nf} (NF pág. {dados_nf['pagina']})" if fonte_nf else "NF não localizada no processo", bool(fonte_nf),
        doc_nf or "não encontrada no documento", bool(doc_nf),
        comparar_numeros(fonte_nf, doc_nf) if fonte_nf and doc_nf else None,
    ))

    linhas.append(_linha_empenho(doc_empenhos, contrato, dados_of))

    # o Valor da capa é o valor efetivamente pago = o da NF (não o autorizado na OF, que pode ser maior)
    m_valor = RE_CP_VALOR.search(texto)
    doc_valor = m_valor.group(1) if m_valor else None
    fonte_valor = dados_nf.get("valor_total")
    linhas.append(linha_tabela(
        "Valor", f"{fonte_valor} (NF pág. {dados_nf['pagina']})" if fonte_valor else "NF não localizada no processo", bool(fonte_valor),
        doc_valor or "não encontrado no documento", bool(doc_valor),
        _valores_monetarios_batem(fonte_valor, doc_valor) if fonte_valor and doc_valor else None,
    ))

    m_elem = RE_CP_ELEMENTO.search(texto)
    doc_nat = m_elem.group(1) if m_elem else None
    fonte_nat = _natureza_do_empenho(contrato, doc_empenhos[0]) if doc_empenhos else None
    naturezas_bd = naturezas_despesa_registradas(contrato) if contrato else []
    if fonte_nat:
        fonte_nat_texto, fonte_nat_disp = f"{fonte_nat} (BD)", True
        bate_nat = _mesmos_digitos(fonte_nat, doc_nat) if doc_nat else None
    elif naturezas_bd:
        fonte_nat_texto, fonte_nat_disp = f"{', '.join(naturezas_bd)} (BD)", True
        bate_nat = any(_mesmos_digitos(n, doc_nat) for n in naturezas_bd) if doc_nat else None
    else:
        fonte_nat_texto, fonte_nat_disp = ("nenhuma ND cadastrada" if contrato else "contrato não encontrado no banco"), False
        bate_nat = None
    linhas.append(linha_tabela(
        "Natureza de Despesa", fonte_nat_texto, fonte_nat_disp,
        doc_nat or "não encontrada no documento", bool(doc_nat),
        bate_nat,
    ))

    return montar_tabela(nome_arquivo, "Capa de Pagamento", indice + 1, linhas)

# ------- Almoxarifado / Documento 6: Termo de Recebimento Definitivo -------

RE_TRD_OSF = re.compile(r"(\d{5}/\d{4})\s+Material\s+\d{4,6}", re.IGNORECASE)
RE_TRD_VALOR_TOTAL_DEF = re.compile(
    r"valor total.*?termo de recebimento provis[óo]rio.*?R\$\s*([\d.,]+)", re.DOTALL | re.IGNORECASE)

def _localizar_trd(paginas):
    indices = [i for i, t in enumerate(paginas)
               if "Termo de Recebimento Definitivo nº" in t and "1 - TERMO DE RECEBIMENTO DEFINITIVO" in t]
    if not indices:
        return None, None
    indice = indices[-1]
    return indice, remover_duplicatas_consecutivas("\n".join(paginas[indice:indice + 4]))

def obter_dados_trd(paginas):
    # itens do Termo de Recebimento Definitivo (o que foi DEFINITIVAMENTE recebido - tem que ser
    # exato com a NF). Cada item na seção "3 - TERMOS RECEBIMENTO PROVISÓRIO": começa com
    # "<OS/F> Material <num> <desc> <unid>" e a quantidade que vale é "Quantidade Informada".
    indice, texto = _localizar_trd(paginas)
    if texto is None:
        return None
    m_sec = re.search(r"3 - TERMOS RECEBIMENTO PROVIS[ÓO]RIO(.*?)(?:\n\s*[4-6] - |\Z)", texto, re.DOTALL | re.IGNORECASE)
    secao = m_sec.group(1) if m_sec else texto
    itens = []
    for pedaco in re.split(r"(?=\d{5}/\d{4}\s+Material\s+\d{4,6}\b)", secao):
        m_id = re.match(r"\d{5}/\d{4}\s+Material\s+(\d{4,6})\s+(.+?)\s+([A-Za-zÀ-Úà-ú]+)\s*(?:\n|$)", pedaco, re.DOTALL)
        m_qi = re.search(r"Quantidade Informada:\s*([\d.,]+)", pedaco, re.IGNORECASE)
        m_vu = re.search(r"Valor Unit[áa]rio:\s*R\$\s*([\d.,]+)", pedaco, re.IGNORECASE)
        m_vt = re.search(r"Valor Total:\s*R\$\s*([\d.,]+)", pedaco, re.IGNORECASE)
        if m_id and (m_qi or m_vu or m_vt):
            desc, _u = _split_desc_unidade(f"{m_id.group(2)} {m_id.group(3)}")
            itens.append({
                "num_item": m_id.group(1), "descricao": desc,
                "quantidade": m_qi.group(1) if m_qi else None,
                "valor_unitario": m_vu.group(1) if m_vu else None,
                "valor_total": m_vt.group(1) if m_vt else None,
            })
    m_tot = RE_TRD_VALOR_TOTAL_DEF.search(texto)
    return {"pagina": indice + 1, "valor_total": m_tot.group(1) if m_tot else None, "itens": itens}

def processar_termo_recebimento_definitivo(nome_arquivo, paginas, contrato, dados_of, processo_p1):
    # Documento 6. Ocupa 4 páginas. Cabeçalho igual ao da OF; os ITENS e o VALOR vão pro bloco
    # "Cruzamento de Itens" (OF × NF × Recebimento), no fim - aqui só cabeçalho + Processo + OF.
    indice, texto = _localizar_trd(paginas)
    if texto is None:
        return None
    dados_of = dados_of or {}

    linhas = _linhas_cabecalho_contrato(texto, contrato)

    # Processo (de pagamento) - o TRD é o único doc de almoxarifado que cita o processo do PDF
    m_proc_pag = ns.RE_PROCESSO.search(texto)
    doc_proc_pag = m_proc_pag.group() if m_proc_pag else None
    linhas.append(linha_tabela(
        "Processo", f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1),
        doc_proc_pag or "não encontrado no documento", bool(doc_proc_pag),
        comparar_textos(processo_p1, doc_proc_pag) if processo_p1 and doc_proc_pag else None,
    ))

    m_osf = RE_TRD_OSF.search(texto)
    doc_osf = m_osf.group(1) if m_osf else None
    fonte_of = dados_of.get("numero")
    linhas.append(linha_tabela(
        "Ordem de Fornecimento", f"{fonte_of} (OF pág. {dados_of['pagina']})" if fonte_of else "OF não localizada no processo", bool(fonte_of),
        doc_osf or "não encontrada no documento", bool(doc_osf),
        comparar_numeros(fonte_of, doc_osf) if fonte_of and doc_osf else None,
    ))

    return montar_tabela(nome_arquivo, "Termo de Recebimento Definitivo", indice + 1, linhas)

# ------- Almoxarifado / Documento 7: Instrumentos de Cobrança (contratos.gov.br) -------

RE_IC_CONTRATO = re.compile(r"Contrato:\s*([\d./]+)")
RE_IC_NUMERO = re.compile(r"N[úu]mero:\s*(\d+)", re.IGNORECASE)
RE_IC_EMISSAO = re.compile(r"Dt\.\s*Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
RE_IC_VALOR_FAT = re.compile(r"Valor Faturado:\s*R\$\s*([\d.,]+)", re.IGNORECASE)
RE_IC_VALOR_LIQ = re.compile(r"Valor L[íi]quido:\s*R\$\s*([\d.,]+)", re.IGNORECASE)
# subelemento na tabela "Empenhos:" da IC - "2026NE510407 07 - GENEROS DE ALIMENTACAO R$ 148,40"
RE_IC_SUBELEMENTO = re.compile(r"\d{4}NE\d{6}\s+(\d{1,3})\s*-\s*\w", re.IGNORECASE)
# RE_IC_OPTANTE e _linha_optante_simples ficam na seção de serviço (Documento 3) - compartilhados

def _subelemento_do_empenho(contrato, numero_empenho):
    # últimos dígitos da natureza de despesa do empenho - o que vem DEPOIS do "-" no BD
    # ("339030-07" -> "07"). None se a ND não tem "-" (não dá pra isolar o subelemento).
    nat = _natureza_do_empenho(contrato, numero_empenho)
    if not nat or "-" not in nat:
        return None
    return nat.split("-")[-1].strip()

def _competencia_de_data(data_br):
    # "24/07/2026" -> "Julho/2026"
    m = re.match(r"\d{2}/(\d{2})/(\d{4})", data_br or "")
    if not m or not (1 <= int(m.group(1)) <= 12):
        return None
    return f"{_MESES_NOME[int(m.group(1))]}/{m.group(2)}"

def processar_instrumento_cobranca_almoxarifado(nome_arquivo, paginas, contrato, dados_of, dados_nf, processo_p1, dados_optante):
    # IC do contratos.gov.br (mesma tela que a de serviço, mas sem PARECER como fonte). p21 é a
    # versão digitalizada ("Registro de IC") - a de conteúdo tem "Valor Faturado" + "Dt. Emissão".
    indices = [i for i, t in enumerate(paginas)
               if "Instrumentos de cobrança" in t and "Valor Faturado" in t and "Dt. Emissão" in t]
    if not indices:
        return None
    indice = indices[-1]
    texto = paginas[indice]
    dados_of = dados_of or {}
    dados_nf = dados_nf or {}

    linhas = []

    m_contrato = RE_IC_CONTRATO.search(texto)
    doc_contrato = m_contrato.group(1) if m_contrato else None
    fonte_contrato = contrato["numero_contrato"] if contrato else None
    linhas.append(linha_tabela(
        "Contrato", f"{fonte_contrato} (BD)" if fonte_contrato else "contrato não encontrado no banco", bool(fonte_contrato),
        doc_contrato or "não encontrado no documento", bool(doc_contrato),
        comparar_numeros(fonte_contrato, doc_contrato) if fonte_contrato and doc_contrato else None,
    ))

    m_proc = ns.RE_PROCESSO.search(texto)
    doc_proc = m_proc.group() if m_proc else None
    linhas.append(linha_tabela(
        "Processo", f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1),
        doc_proc or "não encontrado no documento", bool(doc_proc),
        comparar_textos(processo_p1, doc_proc) if processo_p1 and doc_proc else None,
    ))

    m_nf = RE_IC_NUMERO.search(texto)
    doc_nf = _num_nf(m_nf.group(1)) if m_nf else None
    fonte_nf = dados_nf.get("numero")
    linhas.append(linha_tabela(
        "Nota Fiscal", f"{fonte_nf} (NF pág. {dados_nf['pagina']})" if fonte_nf else "NF não localizada no processo", bool(fonte_nf),
        doc_nf or "não encontrada no documento", bool(doc_nf),
        comparar_numeros(fonte_nf, doc_nf) if fonte_nf and doc_nf else None,
    ))

    m_em = RE_IC_EMISSAO.search(texto)
    doc_em = m_em.group(1) if m_em else None
    fonte_em = dados_nf.get("emissao")
    linhas.append(linha_tabela(
        "Dt. Emissão", f"{fonte_em} (NF pág. {dados_nf['pagina']})" if fonte_em else "NF não localizada no processo", bool(fonte_em),
        doc_em or "não encontrada no documento", bool(doc_em),
        (fonte_em == doc_em) if fonte_em and doc_em else None,
    ))

    doc_empenhos = []
    for ne in RE_ALMOX_NE.findall(texto):
        if ne not in doc_empenhos:
            doc_empenhos.append(ne)
    linhas.append(_linha_empenho(doc_empenhos, contrato, dados_of))

    # Subelemento = os dígitos depois do "-" na ND do empenho no BD ("339030-07" -> "07")
    m_sub = RE_IC_SUBELEMENTO.search(texto)
    doc_sub = m_sub.group(1) if m_sub else None
    fonte_sub = _subelemento_do_empenho(contrato, doc_empenhos[0]) if doc_empenhos else None
    linhas.append(linha_tabela(
        "Subelemento",
        f"{fonte_sub} (BD)" if fonte_sub else "subelemento não isolável na ND do banco",
        bool(fonte_sub),
        doc_sub or "não encontrado no documento", bool(doc_sub),
        _mesmos_digitos(fonte_sub, doc_sub) if fonte_sub and doc_sub else None,
    ))

    # Mês/Ano Referência vs mês da emissão da NF (não há PARECER/competência formal no almoxarifado)
    doc_comp_bruto, doc_comp = extrair_competencia_documento(texto, aceitar_numerico=True)
    fonte_comp = _competencia_de_data(dados_nf.get("emissao"))
    linhas.append(linha_tabela(
        "Competência", f"{fonte_comp} (mês da emissão da NF)" if fonte_comp else "NF não localizada no processo", bool(fonte_comp),
        exibir_competencia(doc_comp_bruto, doc_comp) or "não encontrada no documento", bool(doc_comp_bruto),
        (fonte_comp == doc_comp) if fonte_comp and doc_comp else None,
    ))

    # Valor Faturado/Líquido da IC = o efetivamente faturado = o da NF (não o autorizado na OF)
    fonte_valor = dados_nf.get("valor_total")
    for rotulo, m in (("Valor Faturado", RE_IC_VALOR_FAT.search(texto)), ("Valor Líquido", RE_IC_VALOR_LIQ.search(texto))):
        doc_v = m.group(1) if m else None
        linhas.append(linha_tabela(
            rotulo, f"{fonte_valor} (NF pág. {dados_nf['pagina']})" if fonte_valor else "NF não localizada no processo", bool(fonte_valor),
            doc_v or "não encontrado no documento", bool(doc_v),
            _valores_monetarios_batem(fonte_valor, doc_v) if fonte_valor and doc_v else None,
        ))

    linhas.append(_linha_optante_simples(texto, dados_optante))

    return montar_tabela(nome_arquivo, "Instrumentos de Cobrança", indice + 1, linhas)

def _conformidade_almoxarifado(nome_arquivo, paginas, contrato, processo_p1):
    # cada documento produz dados de referência (dados_of, dados_nf, ...) que os seguintes conferem
    # contra. Implementado documento a documento; os demais entram aqui conforme forem definidos.
    dados_of = obter_dados_of(paginas)
    dados_nf = obter_dados_nf_almoxarifado(paginas)
    dados_trd = obter_dados_trd(paginas)
    dados_optante = obter_dados_optante(paginas)

    doc_ausente = lambda nome: _bloco_ausente(nome_arquivo, nome, None, motivo="Documento não detectado no processo")
    tabelas = [
        processar_ordem_fornecimento(nome_arquivo, paginas, contrato, dados_of) or doc_ausente("Ordem de Serviço / Fornecimento"),
        processar_nota_fiscal_almoxarifado(nome_arquivo, paginas, contrato, dados_of, dados_nf) or doc_ausente("Nota Fiscal"),
        processar_encaminhamento_material(nome_arquivo, paginas, contrato, dados_of, dados_nf) or doc_ausente("Encaminhamento de Material"),
        processar_despacho_ateste_material(nome_arquivo, paginas, contrato, dados_of, dados_nf) or doc_ausente("Despacho de Ateste de Nota Fiscal de Material"),
        processar_capa_pagamento(nome_arquivo, paginas, contrato, dados_of, dados_nf) or doc_ausente("Capa de Pagamento"),
        processar_termo_recebimento_definitivo(nome_arquivo, paginas, contrato, dados_of, processo_p1) or doc_ausente("Termo de Recebimento Definitivo"),
        processar_instrumento_cobranca_almoxarifado(nome_arquivo, paginas, contrato, dados_of, dados_nf, processo_p1, dados_optante) or doc_ausente("Instrumentos de Cobrança"),
        # a tela da Receita "Consulta Optante pelo Simples Nacional" (imagem/OCR) é igual à de
        # serviço - reaproveita o mesmo processador
        processar_consulta_optante(nome_arquivo, paginas, contrato) or doc_ausente("Consulta Optante pelo Simples Nacional"),
    ]
    consistencia = _bloco_consistencia_almoxarifado(nome_arquivo, tabelas, contrato, dados_of, dados_nf, processo_p1)
    if consistencia:
        tabelas.append(consistencia)
    # se a Observação do contrato já tem itens pendentes desta MESMA OF (de um processo anterior),
    # é contra eles que a nova NF é cruzada - não contra a OF inteira
    of_num_curto = _num_curto((dados_of or {}).get("numero"))
    pendentes = _parse_itens_pendentes(contrato.get("observacao"), of_num_curto) if (contrato and of_num_curto) else None

    cruzamento = _bloco_cruzamento_itens(nome_arquivo, dados_of, dados_nf, dados_trd, pendentes)
    if cruzamento:
        tabelas.append(cruzamento)
        _registrar_itens_nao_entregues(contrato, dados_of, dados_nf, pendentes)  # atualiza a pendência no BD
    return tabelas

# ------- ponto de entrada -------

def gerar_conformidade(nome_arquivo, paginas):
    if not paginas or "Processo Eletrônico" not in paginas[0]:
        return []  # não é um PDF de andamento de processo

    match_processo_p1 = ns.RE_PROCESSO.search(paginas[0])
    processo_p1 = match_processo_p1.group() if match_processo_p1 else None

    contrato = localizar_contrato(paginas)
    if (contrato or {}).get("tipo_contrato") == "almoxarifado":
        return _conformidade_almoxarifado(nome_arquivo, paginas, contrato, processo_p1)

    dados_nf = obter_dados_nf(paginas)
    solicitar_dados_manuais_nf(nome_arquivo, dados_nf)
    dados_parecer = obter_dados_parecer(paginas)
    dados_optante = obter_dados_optante(paginas)  # tela da Receita = fonte segura do Optante pelo Simples (ver IC)

    # contrato sem mão de obra não tem IMR nem Termo Circunstanciado do Gestor (a conferência
    # começa na NF) - nesses dois o placeholder de "não detectado" é suprimido; se ainda assim o
    # documento aparecer no PDF, o bloco real é mostrado normalmente. tem_mao_de_obra vem do
    # cadastro do contrato (fonte segura): só 0 suprime - 1, não informado ou contrato não
    # encontrado mantêm a cobrança dos 7.
    sem_mao_de_obra = bool(contrato) and contrato.get("tem_mao_de_obra") == 0

    tabelas = []
    for nome_documento, sigla, so_com_mao_de_obra, funcao in (
        ("Relatório de Avaliação e Medição dos Resultados", "IMR", True,
         lambda: processar_relatorio_avaliacao_medicao(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer)),
        ("Termo Circunstanciado do Gestor do Contrato", None, True,
         lambda: processar_termo_circunstanciado(nome_arquivo, paginas, contrato)),
        ("Nota Fiscal", None, False,
         lambda: processar_nota_fiscal(nome_arquivo, dados_nf, dados_parecer, contrato, sem_mao_de_obra)),
        ("Relatório Circunstanciado de Recebimento Provisório", None, False,
         lambda: processar_relatorio_circunstanciado(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer)),
        ("Despacho de Ateste de Nota Fiscal de Serviço", None, False,
         lambda: processar_despacho_ateste(nome_arquivo, paginas, contrato, dados_nf, dados_parecer)),
        ("Instrumentos de Cobrança", None, False,
         lambda: processar_instrumento_cobranca(nome_arquivo, paginas, processo_p1, contrato, dados_nf, dados_parecer, dados_optante)),
        ("Consulta Optante pelo Simples Nacional", None, False,
         lambda: processar_consulta_optante(nome_arquivo, paginas, contrato)),
    ):
        bloco = funcao()
        if not bloco:
            if so_com_mao_de_obra and sem_mao_de_obra:
                continue  # não existe pra contrato sem mão de obra - não cobra
            # documento não localizado entra na MESMA posição, como placeholder (ver _bloco_ausente)
            bloco = _bloco_ausente(nome_arquivo, nome_documento, dados_parecer)
        if sigla:
            bloco["sigla"] = sigla
        tabelas.append(bloco)

    consistencia = processar_consistencia_documentos(nome_arquivo, tabelas, processo_p1, contrato, dados_nf, dados_parecer)
    if consistencia:
        tabelas.append(consistencia)

    return tabelas

def _extrair_texto_ocr(pagina_pdf):
    # fallback pra páginas sem camada de texto (documento digitalizado como imagem pura, ex:
    # Consulta Optante pelo Simples Nacional - print de tela da Receita Federal) - roda OCR na
    # 1ª imagem embutida da página. Ampliar 3x + escala de cinza melhora bastante a leitura desses
    # prints (testado no euro.pdf); mesmo assim o Tesseract confunde "a"/"o" minúsculos com os
    # indicadores ordinais "ª"/"º" nesse tipo de fonte fina, e o CNPJ vem com travessão "—" em vez
    # de hífen - por isso a normalização no final
    for imagem in pagina_pdf.images:
        try:
            figura = Image.open(io.BytesIO(imagem.data)).convert("L")
            largura, altura = figura.size
            figura = figura.resize((largura * 3, altura * 3), Image.LANCZOS)
            texto = pytesseract.image_to_string(figura, lang="por", config="--psm 6")
        except Exception:
            continue
        return texto.replace("ª", "a").replace("º", "o").replace("—", "-")
    return ""

def _extrair_texto_pagina(pagina_pdf):
    texto = pagina_pdf.extract_text() or ""
    return texto if texto.strip() else _extrair_texto_ocr(pagina_pdf)

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
            fontes.append((url, [_extrair_texto_pagina(p) for p in leitor.pages]))
        navegador.switch_to.window(aba_original)

    caminhos_pdf = list(dict.fromkeys(
        pdf_aberto_windows.listar_pdfs_abertos() + pdf_aberto_windows.listar_pdfs_recentes()
    ))
    for caminho in caminhos_pdf:
        with open(caminho, "rb") as arquivo:
            leitor = PdfReader(arquivo)
            fontes.append((caminho, [_extrair_texto_pagina(p) for p in leitor.pages]))

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
    --azul-doc: #1d6fb8;
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
  .painel h2 { margin: 0 0 10px; font-size: 13px; font-weight: 600; color: var(--azul-doc); display: flex; align-items: baseline; gap: 6px; }
  .painel h2 .sigla-doc { font-weight: 400; color: var(--azul-doc); font-size: 12px; }
  .painel h2 .pagina-doc { font-weight: 400; color: var(--ink-faint); font-size: 12px; }
  .observacao { margin: 10px 0 0; padding-top: 10px; border-top: 1px solid var(--pine-tint-strong); font-size: 12.5px; font-weight: 600; color: var(--status-error); }
  .observacao--solta { margin-top: 0; padding-top: 0; border-top: none; }

  table { width: 100%; border-collapse: collapse; }
  th, td { text-align: left; padding: 8px 10px; font-size: 12.5px; border-bottom: 1px solid var(--hairline); vertical-align: top; }
  th { color: var(--ink-soft); font-weight: 700; font-size: 11px; text-transform: uppercase; letter-spacing: 0.02em; }
  tr:last-child td { border-bottom: none; }
  td.campo { font-weight: 400; white-space: nowrap; }
  tr.linha--destaque td.campo { font-weight: 700; }
  tr.linha--destaque .valor { font-weight: 700; }
  td.resultado { text-align: center; width: 40px; }

  .valor { font-family: "Cascadia Code", Consolas, monospace; font-size: 12px; }
  .valor--indisponivel { font-family: inherit; font-style: italic; color: var(--ink-faint); }
  .valor--erro { color: var(--status-error); font-weight: 700; }

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
      <p class="subtitulo">Documentos preenchidos x Fontes seguras (BD, Termo Gestor e NF).</p>
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

  function celulaValor(texto, disponivel, erro) {
    const spanTexto = document.createElement("span");
    spanTexto.className = "valor" + (disponivel ? "" : " valor--indisponivel") + (erro ? " valor--erro" : "");
    spanTexto.textContent = texto;
    return spanTexto.outerHTML;
  }

  function montarTabelaLinhas(linhas) {
    // numa linha que não confere (❌), o valor do lado "Documento" (o dado errado) vai em vermelho
    const corpo = linhas.map((linha) => `
      <tr class="${linha.destaque ? "linha--destaque" : ""}">
        <td class="campo">${linha.campo}</td>
        <td>${celulaValor(linha.fonte, linha.fonte_disponivel)}</td>
        <td>${celulaValor(linha.documento, linha.documento_disponivel, linha.resultado === "nao")}</td>
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
    const pagina = bloco.pagina ? `<span class="pagina-doc">pág. ${bloco.pagina}</span>` : "";
    const sigla = bloco.sigla ? `<span class="sigla-doc">(${bloco.sigla})</span>` : "";
    const tabela = bloco.linhas.length ? montarTabelaLinhas(bloco.linhas) : "";
    // sem tabela (documento não localizado), a observação vira o conteúdo do painel - tira o
    // divisor/respiro que ela ganha quando vem logo abaixo de uma tabela
    const classeObs = tabela ? "observacao" : "observacao observacao--solta";
    const observacao = bloco.observacao ? `<p class="${classeObs}">${bloco.observacao}</p>` : "";
    return `
      <div class="painel">
        <h2>${bloco.documento} ${sigla} ${pagina}</h2>
        ${tabela}
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
    x, y, largura, altura = janela_windows.geometria_para_tela(1040, 780)
    webview.create_window(
        "CCRGCI - Resultado da Conformidade (NS)", html=HTML_CONFORMIDADE,
        js_api=ApiConformidade(blocos, aviso), width=largura, height=altura, x=x, y=y,
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
