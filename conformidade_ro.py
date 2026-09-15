# -*- coding: utf-8 -*-
"""
Conformidade (RO) - etapa de empenho do pipeline (ver preencher_planilha_ro.py).

Mesma ideia da conformidade.py (que confere a etapa de pagamento / NS): confronta,
documento por documento dentro do PDF de "Andamento do processo" de uma Solicitacao
de empenho, os dados escritos em cada documento contra as fontes seguras (banco de
contratos, tela do Siafi / Registro Orcamentario do proprio processo, pagina 1 do
processo) - e tambem uns contra os outros.

O mesmo processo e reusado o ano inteiro (varios empenhos/reforcos): so a etapa mais
recente e conferida - as paginas depois do ultimo "Despacho: Sem Ocorrencia" ou
"Certificado de Conformidade Sem Ocorrencia" (mesma regra do preencher_planilha_ro.py).

Reaproveita a infra ja testada da conformidade.py:
  - comparadores (comparar_numeros / comparar_textos / comparar_cnpjs / _valores_monetarios_batem)
  - coleta de PDFs em dupla fonte (abas do Chrome em modo debug + PDFs baixados/abertos)
  - remover_duplicatas_consecutivas / limpar_espacos / empenhos_registrados
  - montagem de tabela/linha (montar_tabela / linha_tabela) e a janela de resultado

Documentos implementados (validados contra o 23.pdf, contrato PRIME 17/2023):
  1. Solicitacao de Empenho  - tela SIAFI CONSULTA-CONRO da NC cuja OBSERVACAO e a
     "SOLICITACAO DE EMPENHO DO CONTRATO N ... COM A EMPRESA ...". Confere numero do
     contrato e contratada contra o banco; e a FONTE de PTRES / Fonte / ND / PI / Valor
     (bloco vermelho) pros documentos seguintes conferirem contra.
  2. Dotacao Orcamentaria    - DESPACHO "Assunto: Dotacao Orcamentaria". Confere o
     processo contra a pagina 1 e PTRES / Fonte / Natureza de Despesa / PI / Valor
     contra a Solicitacao de Empenho anterior. Extrai o subelemento ("339039-25" -> 25)
     pro bloco vermelho e pros documentos seguintes.
  3. RO da NE (execucao)     - tela SIAFI CONSULTA-CONRO da NE (com FAVORECIDO e
     DOCUMENTO WEB 2026NE...). Confere favorecido (CNPJ/nome) contra o banco, a CEL.
     ORCAMENTARIA (PTRES/Fonte/ND/PI) e a CLAS.ORC contra a Solicitacao/Dotacao, e a
     SOMA dos valores das ROs da NE do ciclo contra o valor total solicitado. Bloco
     vermelho: valor desta RO + se o empenho ja esta cadastrado no banco.
"""

import re

import gspread  # manipula a Planilha de Controle no Drive
import webview  # mesma lib do gui.py - abre a janela de resultado em cima da instancia ja em execucao
from google.oauth2.service_account import Credentials

import contratos_db
import escolher_planilha
import janela_windows
import pintar_celula_planilha
import conformidade as cf  # reaproveita comparadores, coleta de PDFs, montagem de tabela e a janela

RE_PROCESSO = cf.ns.RE_PROCESSO
# variante tolerante do RE_PROCESSO, so usada como fallback na OBSERVACAO da tela CONRO (texto
# livre de largura fixa, que pode quebrar o processo em qualquer ponto, ex: "2026-\n91") - o
# resultado sai com espaco/quebra de linha embutido, por isso quem usa isso sempre normaliza com
# re.sub(r"\s+", "", ...) antes de comparar
RE_CONRO_PROCESSO_LIVRE = re.compile(r"\d{5}\s*\.\s*\d{6}\s*\.\s*\d{4}\s*-\s*\d{2}")

# amarelo claro 1 - mesmo amarelo padrao usado no relacionar_valor_op.py / baixar_ob.py / etc.
AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255)

# a extracao de texto desse tipo de PDF costuma trocar acento por "?" (U+FFFD) - os
# marcadores textuais usam "." curinga no lugar da letra acentuada
RE_PAG1_PROCESSO_ELETRONICO = re.compile(r"Processo\s+Eletr.nico")
RE_PAG1_SOLIC_EMPENHO = re.compile(r"Solicita..o de empenho")
# marco que fecha um ciclo de empenho: o "Despacho: Sem ocorrencia" da conformidade OU o
# "Certificado de Conformidade Sem Ocorrencia" (alguns processos trazem o certificado no lugar
# do despacho) - o recorte da etapa mais recente comeca depois da ultima ocorrencia de qualquer um.
# O despacho as vezes vem com o setor entre colchetes antes ("[COFCCI] - Sem ocorrência."), que
# esse "\[.+?\]\s*-\s*" opcional cobre - confirmado ao vivo 2026-09-15 (processo 23322.000020.2026-91,
# despacho #1261872, pág. 55): sem isso era o UNICO marco do PDF inteiro (71 paginas) e nao batia,
# entao `corte` ficava em 0 e misturava todos os ciclos anteriores com o atual
RE_MARCO_SEM_OCORRENCIA = re.compile(
    r"(?:Despacho:\s*(?:\[.+?\]\s*-\s*)?|Certificado\s+de\s+Conformidade\s+)Sem\s+ocorr.ncia",
    re.IGNORECASE,
)

RE_CONRO_TITULO = re.compile(r"CONSULTA-CONRO|CONSULTA REGISTRO ORCAMENTARIO")
RE_CONRO_RO_NUMERO = re.compile(r"NUMERO\s*:\s*(2026RO\d+)")


def _somar_valores(valores):
    total = sum(cf._valor_para_float(v) for v in valores)
    return cf._float_para_valor_br(total)


# ======= Documento 1: Solicitacao de Empenho (tela SIAFI CONSULTA-CONRO da NC) =======

# "SOLICITACAO DE EMPENHO DO CONTRATO N 17/2023 COM A EMPRESA PRIME ..., REFERENTE ..." ou, quando
# o empenho e de uma contratacao direta (ex: fornecimento de energia), "SOLICITACAO DE EMPENHO DA
# CONTRATACAO 90037/2025 COM A EMPRESA ..." (sem o "N", nº com 5 digitos = pregao/contratacao)
RE_CONRO_SOLIC_CONTRATO = re.compile(
    r"SOLICITA..O DE EMPENHO D[AO]\s+(?:CONTRATO|CONTRATA..O)\s+(?:N.?\s*)?(\d+/\d{4})"
)
RE_CONRO_DOC_WEB_NC = re.compile(r"DOCUMENTO WEB\s*:\s*(2026NC\d+)")
# "COM A EMPRESA X, REFERENTE" (contrato) OU "DA EMPRESA X, REFERENTE" (contratação direta) -
# para no 1º "," em vez de exigir "REFERENTE" completo porque a extração do PDF pode quebrar
# essa palavra no meio ("REFER\nENTE", quebra de linha da tela de largura fixa do SIAFI) -
# confirmado ao vivo pelo usuário 2026-09-14 (empresa BONJE GAS LTDA não era achada)
RE_CONRO_EMPRESA = re.compile(r"(?:COM\s+A|DA)\s+EMPRESA\s+(.+?)\s*,", re.DOTALL)
# linha de evento da tela de eventos: "001 301202 ... 14.222,41" (evento + valor no fim)
RE_CONRO_EVENTO_VALOR = re.compile(r"^\s*\d{3}\s+(\d{6})\b.*?(\d{1,3}(?:\.\d{3})*,\d{2})\s*$", re.M)
# linha da celula orcamentaria (logo abaixo da linha de evento):
# "            1   231634 1000000000 339039        L20RLP01RTN"  (ESF PTRES FONTE ND [UGR] [PI])
RE_CONRO_CELULA = re.compile(r"^\s+\d\s+(\d{4,6})\s+(\d{8,10})\s+(\d{6})\s*(?:\s+(\S+))?\s*$", re.M)


def _celula_orcamentaria(texto):
    # devolve (ptres, fonte, [ND das linhas que tem PI], [PIs distintos]) - as linhas COM
    # PI sao o detalhamento por natureza de despesa (evento 301202); a linha sem PI (evento
    # 301201) so traz a ND generica 339000, que nao interessa como classificacao
    ptres = fonte = ""
    nds, pis = [], []
    for m in RE_CONRO_CELULA.finditer(texto):
        p, f, nd, pi = m.group(1), m.group(2), m.group(3), (m.group(4) or "").strip()
        if not ptres:
            ptres, fonte = p, f
        if pi:
            if nd not in nds:
                nds.append(nd)
            if pi not in pis:
                pis.append(pi)
    return ptres, fonte, nds, pis


def _valor_total_conro(texto):
    # o valor total do empenho e o da linha do evento 301201 (as 301202 sao o detalhamento
    # por ND e somadas dao o mesmo total); se nao achar, cai pro ultimo valor da tela
    valores = RE_CONRO_EVENTO_VALOR.findall(texto)  # [(evento, valor), ...]
    for evento, valor in valores:
        if evento == "301201":
            return valor
    return valores[-1][1] if valores else ""


# ======= Documento 0: Requerimento de Empenho (formulario nativo do Suap, ANTES da tela SIAFI) =======
# "REQUERIMENTO nn/aaaa ... SOLICITAÇÃO DE EMPENHO DE CONTRATO" - preenchido pelo requisitante do
# setor, normalmente a 1ª pagina do ciclo (antes da Solicitacao de Empenho/tela SIAFI sequer
# existir). Texto limpo (sem a corrupcao de acento das telas verdes SIAFI), mas quebra numeros em
# varias linhas por causa da formatacao em negrito/tabela do Suap - todo regex usa \s*\n?\s* entre
# os grupos por causa disso.
RE_REQ_NUMERO = re.compile(r"REQUERIMENTO\s+(\d+/\d{4})")
RE_REQ_CONTRATO = re.compile(r"Solicita..o de empenho do contrato\s*\n?\s*(\d+/\d{4})", re.IGNORECASE)
RE_REQ_OBJETO = re.compile(r"OBJETO:\s*\n?(.+?)\n\s*VALOR:", re.DOTALL | re.IGNORECASE)
# o campo OBJETO inteiro é um preâmbulo padrão ("Solicitação de empenho do contrato X/aaaa,
# referente ao exercício de aaaa, para o campus NOME referente a <objeto real>.") - a redação do
# preâmbulo varia (às vezes tem "campus NOME" antes do "referente a" final, às vezes não, e pode
# ter um "referente a"/"referente ao" mais cedo também, ex: "referente ao exercício de") -
# comparar_textos só aceita prefixo/início igual, então isola o trecho depois da ÚLTIMA ocorrência
# de "referente a[s]", que é sempre o objeto real, não importa o que veio antes
RE_REQ_REFERENTE_A = re.compile(r"referente\s*\n?\s*a[s]?\s+", re.IGNORECASE)


def _objeto_real_requerimento(texto_objeto):
    partes = RE_REQ_REFERENTE_A.split(texto_objeto)
    return partes[-1] if len(partes) > 1 else texto_objeto
RE_REQ_VALOR_TOTAL = re.compile(r"VALOR:\s*\n?\s*R\$\s*\n?\s*([\d.]+,\d{2})", re.IGNORECASE)
RE_REQ_EMPRESA = re.compile(r"EMPRESA:\s*\n?\s*(.+?)\s*\n", re.IGNORECASE)
RE_REQ_CNPJ = re.compile(r"CNPJ:\s*\n?\s*([\d./-]+)", re.IGNORECASE)
# recorte só da tabela de itens (do cabeçalho "Valor total" até "EMPRESA:") - importante NÃO
# buscar "R$ x,xx" na página inteira, porque o próprio "VALOR: R$ x,xx" do requerimento (linha
# ANTES da tabela) também bateria e desalinharia a paridade unitário/total dos itens
RE_REQ_TABELA_ITENS = re.compile(r"Valor total\s*\n(.+?)\n\s*EMPRESA:", re.DOTALL | re.IGNORECASE)
# valores "R$ x,xx" dentro do recorte acima (Item/Descrição/Quantidade/Valor unitário/Valor
# total) - confirmado ao vivo que a extração alterna unitário/total por item, nessa ordem
# ("02 P45 16 R$ 425,92 R$ 6.814,72"), entao os de indice IMPAR (1, 3, 5...) sao os totais
RE_REQ_VALOR_ITEM = re.compile(r"R\$\s*([\d.]+,\d{2})")


def processar_requerimento_empenho(nome_arquivo, paginas, contrato, corte):
    # devolve (blocos, requerimentos) - requerimentos e a lista (em ordem de pagina) dos valores
    # que cada Requerimento declara, pra Solicitacao de Empenho (NC) conferir contra - o
    # Requerimento vem ANTES da tela SIAFI no processo, entao ele e a fonte segura do Valor da NC
    # (nao o contrario): pedido do usuario 2026-09-14, "o requerimento passa a ser a fonte segura
    # para a NC"
    blocos, requerimentos = [], []

    # esse formulario (com a tabela de itens) so existe em contratos de ALMOXARIFADO - contrato de
    # SERVICO tem um "REQUERIMENTO nn/aaaa" com layout diferente (sem essa tabela/campos), que os
    # regex abaixo nao reconhecem - sem esse filtro, um requerimento de servico virava um bloco
    # inteiro "nao encontrado" em vez de simplesmente nao aparecer (pedido do usuario 2026-09-14)
    if contrato and contrato.get("tipo_contrato") != "almoxarifado":
        return blocos, requerimentos

    for idx in range(corte, len(paginas)):
        texto = paginas[idx]
        m_numero = RE_REQ_NUMERO.search(texto)
        if not m_numero:
            continue  # nao e a pagina do Requerimento

        linhas = []

        m_contrato = RE_REQ_CONTRATO.search(texto)
        doc_contrato = m_contrato.group(1) if m_contrato else ""
        contrato_bd = contrato["numero_contrato"] if contrato else ""
        linhas.append(cf.linha_tabela(
            "Contrato",
            f"{contrato_bd} (BD)" if contrato_bd else "contrato não encontrado no banco", bool(contrato_bd),
            doc_contrato or "não encontrado", bool(doc_contrato),
            cf.comparar_numeros(contrato_bd, doc_contrato) if (contrato_bd and doc_contrato) else None,
        ))

        m_objeto = RE_REQ_OBJETO.search(texto)
        doc_objeto = re.sub(r"\s+", " ", _objeto_real_requerimento(m_objeto.group(1))).strip() if m_objeto else ""
        fonte_objeto_texto, fonte_objeto_disp, bate_objeto = cf._conferir_objeto(contrato, doc_objeto)
        linhas.append(cf.linha_tabela(
            "Objeto", fonte_objeto_texto, fonte_objeto_disp, doc_objeto or "não encontrado", bool(doc_objeto), bate_objeto,
        ))

        m_valor = RE_REQ_VALOR_TOTAL.search(texto)
        doc_valor = m_valor.group(1) if m_valor else ""

        # se tem mais de 1 item, confere se a soma dos "Valor total" de cada item bate com o
        # VALOR geral do requerimento - com só 1 item a checagem seria redundante (o próprio
        # "Valor total" do item já É o valor geral)
        m_tabela = RE_REQ_TABELA_ITENS.search(texto)
        valores_tabela = RE_REQ_VALOR_ITEM.findall(m_tabela.group(1)) if m_tabela else []  # [unit1, total1, ...]
        totais_itens = valores_tabela[1::2]
        if len(totais_itens) > 1:
            soma_itens = _somar_valores(totais_itens)
            linhas.append(cf.linha_tabela(
                "Soma dos Itens",
                f"{' + '.join(totais_itens)} = {soma_itens}", True,
                doc_valor or "não encontrado", bool(doc_valor),
                cf._valores_monetarios_batem(soma_itens, doc_valor) if doc_valor else None,
            ))

        m_empresa = RE_REQ_EMPRESA.search(texto)
        doc_empresa = cf.limpar_espacos(m_empresa.group(1)) if m_empresa else ""
        contratada_bd = contrato["nome_contratada"] if contrato else ""
        linhas.append(cf.linha_tabela(
            "Empresa",
            f"{contratada_bd} (BD)" if contratada_bd else "contrato não encontrado no banco", bool(contratada_bd),
            doc_empresa or "não encontrada", bool(doc_empresa),
            cf.comparar_textos(contratada_bd, doc_empresa) if (contratada_bd and doc_empresa) else None,
        ))

        m_cnpj = RE_REQ_CNPJ.search(texto)
        doc_cnpj = re.sub(r"\D", "", m_cnpj.group(1)) if m_cnpj else ""
        cnpj_bd = contrato["cnpj"] if contrato else ""
        linhas.append(cf.linha_tabela(
            "CNPJ",
            f"{cf._formatar_cnpj(cnpj_bd)} (BD)" if cnpj_bd else "contrato não encontrado no banco", bool(cnpj_bd),
            cf._formatar_cnpj(doc_cnpj) if doc_cnpj else "não encontrado", bool(doc_cnpj),
            cf.comparar_cnpjs(cnpj_bd, doc_cnpj) if (cnpj_bd and doc_cnpj) else None,
        ))

        titulo = f"Requerimento de Empenho {m_numero.group(1)}"
        observacao = f"Valor: {doc_valor}" if doc_valor else None
        blocos.append(cf.montar_tabela(nome_arquivo, titulo, idx + 1, linhas, observacao=observacao))
        requerimentos.append({"pagina": idx + 1, "valor": doc_valor})

    return blocos, requerimentos


def _requerimento_anterior(requerimentos, pagina):
    # o ultimo Requerimento de Empenho que aparece ANTES desta pagina - mesmo raciocinio de
    # _referencia_anterior, mas pro Requerimento (que vem antes da tela SIAFI no processo)
    anterior = None
    for r in requerimentos:
        if r["pagina"] < pagina:
            anterior = r
    return anterior


def processar_solicitacao_empenho(nome_arquivo, paginas, contrato, requerimentos, processo_p1, corte):
    # devolve (blocos, referencias) - referencias e a lista (em ordem de pagina) dos dados
    # orcamentarios que cada Solicitacao determina, pros documentos seguintes conferirem contra
    blocos, referencias = [], []

    for idx in range(corte, len(paginas)):
        texto_pag = paginas[idx]
        if not RE_CONRO_TITULO.search(texto_pag):
            continue
        m_contrato = RE_CONRO_SOLIC_CONTRATO.search(texto_pag)
        if not m_contrato:
            continue  # e uma tela CONSULTA-CONRO, mas nao a da solicitacao de empenho (ex: RO da NE)

        m_ro = RE_CONRO_RO_NUMERO.search(texto_pag)
        ro_numero = m_ro.group(1) if m_ro else ""
        m_nc = RE_CONRO_DOC_WEB_NC.search(texto_pag)
        nc_numero = m_nc.group(1) if m_nc else ""

        # o detalhamento orcamentario (PTRES/Fonte/ND/PI/Valor) fica na tela de eventos,
        # que e a pagina seguinte do mesmo RO - junta as duas
        texto = texto_pag
        if idx + 1 < len(paginas) and ro_numero and ro_numero in paginas[idx + 1]:
            texto = texto_pag + "\n" + paginas[idx + 1]

        linhas = []

        doc_contrato = m_contrato.group(1)
        contrato_bd = contrato["numero_contrato"] if contrato else ""
        linhas.append(cf.linha_tabela(
            "Contrato",
            f"{contrato_bd} (BD)" if contrato_bd else "contrato nao encontrado no banco", bool(contrato_bd),
            doc_contrato, True,
            cf.comparar_numeros(contrato_bd, doc_contrato) if contrato_bd else None,
        ))

        m_empresa = RE_CONRO_EMPRESA.search(texto)
        doc_empresa = cf.limpar_espacos(m_empresa.group(1)) if m_empresa else ""
        contratada_bd = contrato["nome_contratada"] if contrato else ""
        linhas.append(cf.linha_tabela(
            "Contratada",
            f"{contratada_bd} (BD)" if contratada_bd else "contrato nao encontrado no banco", bool(contratada_bd),
            doc_empresa or "nao encontrada", bool(doc_empresa),
            # o SIAFI costuma abreviar o nome (ex: so "PRIME") - comparar_textos aceita
            # quando um lado e prefixo do outro
            cf.comparar_textos(contratada_bd, doc_empresa) if (contratada_bd and doc_empresa) else None,
        ))

        # RE_CONRO_PROCESSO_LIVRE (fallback) porque a OBSERVACAO as vezes quebra o processo bem
        # antes do fim ("2026-\n91") - confirmado ao vivo 2026-09-15 (23322.000020.2026-91)
        m_processo = RE_PROCESSO.search(texto) or RE_CONRO_PROCESSO_LIVRE.search(texto)
        doc_processo = re.sub(r"\s+", "", m_processo.group()) if m_processo else ""
        linhas.append(cf.linha_tabela(
            "Processo",
            f"{processo_p1} (pág. 1)" if processo_p1 else "nao encontrado", bool(processo_p1),
            doc_processo or "nao encontrado", bool(doc_processo),
            cf.comparar_textos(processo_p1, doc_processo) if (processo_p1 and doc_processo) else None,
        ))

        req = _requerimento_anterior(requerimentos, idx + 1)
        req_valor = req["valor"] if req else ""
        ptres, fonte, nds, pis = _celula_orcamentaria(texto)
        valor = _valor_total_conro(texto)
        linhas.append(cf.linha_tabela(
            "Valor",
            f"{req_valor} (Requerimento)" if req_valor else "sem Requerimento anterior", bool(req_valor),
            valor or "nao encontrado", bool(valor),
            cf._valores_monetarios_batem(req_valor, valor) if (req_valor and valor) else None,
        ))

        # PTRES/Fonte/ND/PI continuam extraídos (ptres, fonte, nds, pis) e guardados em
        # `referencias` pra Dotação/RO da NE conferirem contra - só a exibição na nota vermelha
        # deste bloco foi removida (pedido do usuário 2026-09-14: só Contrato/Contratada/Valor)

        # nome no padrao "Registro Orçamentário <RO>-<NC>"
        partes_titulo = "-".join(p for p in (ro_numero, nc_numero) if p)
        titulo = f"Registro Orçamentário {partes_titulo}" if partes_titulo else "Registro Orçamentário"
        blocos.append(cf.montar_tabela(nome_arquivo, titulo, idx + 1, linhas))
        referencias.append({
            "pagina": idx + 1, "ro": ro_numero,
            "ptres": ptres, "fonte": fonte, "nds": nds, "pis": pis, "valor": valor,
            "clas_orc": [],  # preenchido pela Dotacao Orcamentaria (ND + subelemento, ex: "33903925")
        })

    return blocos, referencias


def _referencia_anterior(referencias, pagina):
    # a ultima Solicitacao de Empenho que aparece ANTES desta pagina no PDF ("doc anterior")
    anterior = None
    for r in referencias:
        if r["pagina"] < pagina:
            anterior = r
    return anterior


# ======= Documento 2: Dotacao Orcamentaria (DESPACHO "Assunto: Dotacao Orcamentaria") =======

RE_DOT_ASSUNTO = re.compile(r"Assunto:\s*Dota..o Or.ament.ria", re.IGNORECASE)
RE_DOT_DESPACHO = re.compile(r"DESPACHO\s+(\d+/\d{4})")
# quando um despacho pede a retificação de outro ("... PARA VERIFICAÇÃO/RETIFICAÇÃO DO DESPACHO
# 288/2026.") o despacho retificado deixa de valer - só o que retifica (o mais recente) deve ser
# considerado. Nem todo processo tem isso (pedido do usuário 2026-09-14: "isso pode não ocorrer")
RE_DOT_RETIFICACAO_DE = re.compile(r"RETIFICA..O\s+DO\s+DESPACHO\s+(?:N.?\s*)?(\d+/\d{4})", re.IGNORECASE)

# despacho "Encaminho para dotação orçamentária: Valor: R$ X | Nota de Empenho: 2026NExxxxxx" -
# geralmente em contratos de SERVIÇO (não tem o Requerimento de Empenho, que é do fluxo de
# material/almoxarifado) - quando existe, mostra como nota vermelha (não como coluna) na tabela
# "Consistência Orçamentária entre Documentos", pro usuário conferir visualmente contra a RO da NE.
# Nem todo processo tem esse despacho (pedido do usuário 2026-09-14: "geralmente nos contratos de
# serviço")
RE_DESP_ENCAMINHA_DOTACAO = re.compile(
    r"Encaminho\s+para\s+dota..o\s+or.ament.ria.*?Valor:\s*\n?\s*R\$\s*\n?\s*([\d.]+,\d{2})"
    r".*?Nota\s+de\s+Empenho:\s*\n?\s*(2026NE\d+)",
    re.IGNORECASE | re.DOTALL,
)


def _nota_solicitacao_servico(paginas, corte):
    for pagina in paginas[corte:]:
        m = RE_DESP_ENCAMINHA_DOTACAO.search(pagina)
        if m:
            return f"Solicitação: R$ {m.group(1)} | {m.group(2)}"
    return None
# uma linha de alocacao vem como um bloco de linhas: UG / PTRES / FONTE / ND / UGR / PI / VALOR.
# O valor as vezes vem com "R$" na frente (23.pdf), as vezes sem (26.pdf) - por isso opcional
RE_DOT_ALOCACAO = re.compile(
    r"(\d{6})\s*\n\s*(\d{6})\s*\n\s*(\d{8,10})\s*\n\s*(\d{6})\s*\n\s*(.+?)\s*\n\s*(\S+)\s*\n\s*(?:R\$\s*)?([\d.]+,\d{2})",
    re.M,
)
# subelemento no despacho: "339039-25". A extracao do PDF costuma meter espaco ou quebra
# de linha em volta do hifen ("339039 - 17", "339039-\n17") e as vezes troca o "-" por
# en/em-dash - aceita tudo isso, mas exige algum traco separando (sem ele, "339039" casaria
# com qualquer "17" solto perto na tela)
# o "nd" capturado aqui às vezes sai com 1 dígito a menos do que deveria (ex: "39039-78" em vez
# de "339039-78" no texto descritivo do subelemento, mesmo com a tabela estruturada acima
# mostrando "339039" certo) - confirmado ao vivo 2026-09-15 contra o documento real, não é
# artefato de extração do PDF, o próprio texto do despacho vem assim. Por isso \d{5,6} (tolerante)
# em vez de \d{6} fixo, e quem usa isso reconstrói o ND a partir de `doc_nds` (a tabela
# estruturada, sempre correta), não do nd capturado aqui - só o subelemento (2 dígitos) importa.
# NÃO usar \d{4,6}: 4 dígitos bate com o "aaaa-nn" do número do processo (ex: "2026-91"), gerando
# um subelemento falso - confirmado ao vivo (deu "Subelemento: 91 e 78" com \d{4,6})
RE_DOT_SUBELEMENTO = re.compile(r"\b(\d{5,6})\s*[-‐-―]\s*(\d{2})\b")


def _linha_igualdade(campo, esperado, obtido, igual, rotulo="NC"):
    # esperado (fonte segura) vem da NC (Registro Orcamentario) anterior; obtido, deste documento.
    # rotulo: CLAS.ORC é a exceção - a NC não relata subelemento, quem preenche `clas_orc` no
    # `ref` compartilhado é a Dotação Orçamentária (ver _ORC_DOTACAO_VIA_REF) - "(NC)" sozinho
    # seria enganoso ali, por isso o rotulo "NC + Dotação" nesse caso
    tem_ref = bool(esperado)
    return cf.linha_tabela(
        campo,
        (f"{esperado} ({rotulo})") if tem_ref else f"sem {rotulo} anterior", tem_ref,
        obtido or "nao encontrado", bool(obtido),
        igual if (tem_ref and obtido) else None,
    )


def processar_dotacao_orcamentaria(nome_arquivo, paginas, processo_p1, referencias, corte):
    blocos = []

    # despachos que algum outro despacho do ciclo pediu pra retificar - excluidos abaixo, so o
    # despacho retificador (mais recente) e considerado
    despachos_retificados = {
        m.group(1) for pagina in paginas[corte:] for m in RE_DOT_RETIFICACAO_DE.finditer(pagina)
    }

    # pagina de cada despacho "Assunto: Dotacao Orcamentaria" do ciclo (inclusive os retificados,
    # que nao geram bloco) - so pra anotar "Retificação do X/aaaa (pág. N)" no despacho retificador
    pagina_por_despacho = {}
    for idx in range(corte, len(paginas)):
        if not RE_DOT_ASSUNTO.search(paginas[idx]):
            continue
        m = RE_DOT_DESPACHO.search(cf.remover_duplicatas_consecutivas(paginas[idx]))
        if m:
            pagina_por_despacho[m.group(1)] = idx + 1

    for idx in range(corte, len(paginas)):
        texto_pag = paginas[idx]
        if not RE_DOT_ASSUNTO.search(texto_pag):
            continue
        texto = cf.remover_duplicatas_consecutivas(texto_pag)  # esse despacho vem com cada linha duplicada

        m_desp_atual = RE_DOT_DESPACHO.search(texto)
        if m_desp_atual and m_desp_atual.group(1) in despachos_retificados:
            continue  # despacho retificado por outro - nao entra na conferencia

        # este despacho sobreviveu ao filtro acima - se algum outro foi retificado no ciclo,
        # assume que é este que retifica (caso comum: só sobra 1 despacho quando há retificação)
        retificados_por_este = [n for n in despachos_retificados if n != (m_desp_atual.group(1) if m_desp_atual else None)]
        nota_retificacao = None
        if retificados_por_este:
            partes_retificacao = [
                f"{n} (pág. {pagina_por_despacho[n]})" if n in pagina_por_despacho else n
                for n in retificados_por_este
            ]
            nota_retificacao = f"Retificação do {' e '.join(partes_retificacao)}"

        ref = _referencia_anterior(referencias, idx + 1)

        alocacoes = RE_DOT_ALOCACAO.findall(texto)  # [(ug, ptres, fonte, nd, ugr, pi, valor), ...]
        doc_ptres = alocacoes[0][1] if alocacoes else ""
        doc_fonte = alocacoes[0][2] if alocacoes else ""
        doc_nds = list(dict.fromkeys(a[3] for a in alocacoes))
        doc_pis = list(dict.fromkeys(a[5] for a in alocacoes))
        doc_valor = _somar_valores([a[6] for a in alocacoes]) if alocacoes else ""

        linhas = []

        doc_processo = RE_PROCESSO.search(texto)
        doc_processo = doc_processo.group() if doc_processo else ""
        linhas.append(cf.linha_tabela(
            "Processo",
            f"{processo_p1} (pág. 1)" if processo_p1 else "nao encontrado", bool(processo_p1),
            doc_processo or "nao encontrado", bool(doc_processo),
            cf.comparar_textos(processo_p1, doc_processo) if (processo_p1 and doc_processo) else None,
        ))

        ref_ptres = ref["ptres"] if ref else ""
        ref_fonte = ref["fonte"] if ref else ""
        ref_nds = ref["nds"] if ref else []
        ref_pis = ref["pis"] if ref else []
        ref_valor = ref["valor"] if ref else ""

        linhas.append(_linha_igualdade("PTRES", ref_ptres, doc_ptres, ref_ptres == doc_ptres))
        linhas.append(_linha_igualdade("Fonte", ref_fonte, doc_fonte, ref_fonte == doc_fonte))
        linhas.append(_linha_igualdade(
            "Natureza de Despesa",
            cf.ns.juntar_com_e(ref_nds) if ref_nds else "",
            cf.ns.juntar_com_e(doc_nds) if doc_nds else "",
            set(ref_nds) == set(doc_nds),
        ))
        linhas.append(_linha_igualdade(
            "PI",
            cf.ns.juntar_com_e(ref_pis) if ref_pis else "",
            cf.ns.juntar_com_e(doc_pis) if doc_pis else "",
            set(ref_pis) == set(doc_pis),
        ))
        linhas.append(_linha_igualdade(
            "Valor", ref_valor, doc_valor,
            cf._valores_monetarios_batem(ref_valor, doc_valor) if (ref_valor and doc_valor) else False,
        ))

        # subelemento: "339039-25 TAXA DE ADMINISTRACAO ..." -> subelemento 25. Guarda na
        # referencia (ND + subelemento sem pontuacao, ex: "33903925") pro RO da NE conferir
        # a CLAS.ORC contra, e mostra no bloco vermelho. O ND usado aqui vem de `doc_nds` (a
        # tabela estruturada, sempre com 6 dígitos certos) - NÃO do "nd" capturado por
        # RE_DOT_SUBELEMENTO, que às vezes sai truncado (ver comentário do regex)
        pares_sub = list(dict.fromkeys(RE_DOT_SUBELEMENTO.findall(texto)))  # [(nd_bruto, sub), ...]
        nds_confiaveis = doc_nds or [nd_bruto for nd_bruto, _ in pares_sub]
        if ref is not None:
            for _, sub in pares_sub:
                for nd in nds_confiaveis:
                    if nd + sub not in ref["clas_orc"]:
                        ref["clas_orc"].append(nd + sub)
        # o subelemento e so o que vem depois do traco ("339039-25" -> "25")
        subelementos = list(dict.fromkeys(sub for _, sub in pares_sub))
        partes_observacao = []
        if nota_retificacao:
            partes_observacao.append(nota_retificacao)
        if subelementos:
            partes_observacao.append(f"Subelemento: {cf.ns.juntar_com_e(subelementos)}")
        observacao = " | ".join(partes_observacao) if partes_observacao else None

        titulo = f"Dotação Orçamentária — Despacho {m_desp_atual.group(1)}" if m_desp_atual else "Dotação Orçamentária"
        blocos.append(cf.montar_tabela(nome_arquivo, titulo, idx + 1, linhas, observacao=observacao))

    return blocos


# ======= Documento 3: RO da NE / execucao do empenho (tela SIAFI CONSULTA-CONRO da NE) =======

RE_RONE_DOC_WEB_NE = re.compile(r"DOCUMENTO WEB\s*:\s*(2026NE\d+)")
RE_RONE_CEL_ORCAMENTARIA = re.compile(
    r"CEL\.\s*ORCAMENTARIA\s*:\s*\d\s+(\d{4,6})\s+(\d{8,10})\s+(\d{6})\s+(\S+)"
)
# tela de eventos: "001 401202 ... 33903925" e o valor na linha de baixo
RE_RONE_CLAS_ORC = re.compile(r"CLAS\.CONT\s+CLAS\.ORC.*?\n\s*\d{3}\s+\d{6}\D*?(\d{8})\b", re.S)
RE_RONE_VALOR_EVENTO = re.compile(r"\b\d{8}\b\s*\n\s*(\d{1,3}(?:\.\d{3})*,\d{2})")
# codigo do evento (linha "001 401201"/"001 401202"/"001 401203") - 401201 = empenho original
# (unico que tem Processo/Processo de Contratação na OBSERVACAO); 401202 (reforço) e 401203
# (cancelamento) tem uma OBSERVACAO diferente ("REGISTRO DE ANULACAO/REFORCO/CANCELAMENTO DO
# EMPENHO...", sem processo nenhum) - confirmado pelo usuário 2026-09-15, não é bug de extração
RE_RONE_EVENTO = re.compile(r"^\s*\d{3}\s+(4012\d{2})\b", re.M)
# OBSERVACAO da tela: "... PROCESSO nnnnn.nnnnnn.aaaa-nn (nnnnn.nnnnnn.aaaa-nn)" - o 1º é o
# processo corrente (confere contra a pág. 1), o 2º (entre parênteses) é o processo de
# contratação (confere contra o BD) - mesmo formato usado no CONOB/CONNS, valor diferente.
# \s* entre os grupos porque o nº entre parênteses pode quebrar de linha no meio (a extração do
# PDF segue a quebra de linha visual da tela: "...  (23323\n.001539.2024-14)") - confirmado
# ao vivo pelo usuário 2026-09-14
RE_RONE_PROCESSO_CONTRATACAO = re.compile(r"\((\d{5}\s*\.\s*\d{6}\s*\.\s*\d{4}\s*-\s*\d{2})\)")


def processar_ro_da_ne(nome_arquivo, paginas, contrato, referencias, corte, processo_p1):
    # coleta todas as telas RO da NE do ciclo (com FAVORECIDO + DOCUMENTO WEB 2026NE...);
    # o valor total empenhado costuma ser dividido entre 2+ ROs da NE - a SOMA delas e que
    # tem que fechar com o valor da Solicitacao de Empenho
    telas = []  # [(idx, texto), ...]
    for idx in range(corte, len(paginas)):
        texto = paginas[idx]
        if not RE_CONRO_TITULO.search(texto):
            continue
        if not RE_RONE_DOC_WEB_NE.search(texto):
            continue
        if not cf.ns.RE_EMPRESA.search(texto):
            continue
        telas.append((idx, texto))

    nes = []
    for _, texto in telas:
        m = RE_RONE_DOC_WEB_NE.search(texto)
        if m and m.group(1) not in nes:
            nes.append(m.group(1))

    if not telas:
        return [], nes

    valores_ro = []
    for _, texto in telas:
        m = RE_RONE_VALOR_EVENTO.search(texto)
        if m:
            valores_ro.append(m.group(1))
    soma_valor = _somar_valores(valores_ro) if valores_ro else ""
    soma_texto = " + ".join(valores_ro)
    if len(valores_ro) > 1:
        soma_texto += f" = {soma_valor}"

    ref = _referencia_anterior(referencias, telas[0][0] + 1)
    ref_ptres = ref["ptres"] if ref else ""
    ref_fonte = ref["fonte"] if ref else ""
    ref_nds = ref["nds"] if ref else []
    ref_pis = ref["pis"] if ref else []
    ref_valor = ref["valor"] if ref else ""
    ref_clas_orc = ref["clas_orc"] if ref else []

    cnpj_bd = contrato["cnpj"] if contrato else ""
    contratada_bd = contrato["nome_contratada"] if contrato else ""

    blocos = []
    for idx, texto in telas:
        m_ro = RE_CONRO_RO_NUMERO.search(texto)
        ro_numero = m_ro.group(1) if m_ro else ""
        m_fav = cf.ns.RE_EMPRESA.search(texto)
        doc_cnpj = re.sub(r"\D", "", m_fav.group(1)) if m_fav else ""
        doc_favorecido = cf.limpar_espacos(m_fav.group(2)) if m_fav else ""

        linhas = []
        linhas.append(cf.linha_tabela(
            "CNPJ",
            f"{cf._formatar_cnpj(cnpj_bd)} (BD)" if cnpj_bd else "contrato nao encontrado no banco", bool(cnpj_bd),
            cf._formatar_cnpj(doc_cnpj) if doc_cnpj else "nao encontrado", bool(doc_cnpj),
            cf.comparar_cnpjs(cnpj_bd, doc_cnpj) if (cnpj_bd and doc_cnpj) else None,
        ))
        linhas.append(cf.linha_tabela(
            "Favorecido",
            f"{contratada_bd} (BD)" if contratada_bd else "contrato nao encontrado no banco", bool(contratada_bd),
            doc_favorecido or "nao encontrado", bool(doc_favorecido),
            cf.comparar_textos(contratada_bd, doc_favorecido) if (contratada_bd and doc_favorecido) else None,
        ))

        # Processo/Processo de Contratação só existem na OBSERVACAO do evento 401201 (empenho
        # original) - 401202 (reforço)/401203 (cancelamento) tem outra OBSERVACAO sem processo
        # nenhum, então nem tenta essas 2 linhas nesse caso (evita "não encontrado" enganoso)
        m_evento = RE_RONE_EVENTO.search(texto)
        if not m_evento or m_evento.group(1) == "401201":
            m_proc = cf.RE_SIAFI_PROCESSO_PONTUADO.search(texto)
            doc_processo = m_proc.group(1) if m_proc else ""
            linhas.append(cf.linha_tabela(
                "Processo",
                f"{processo_p1} (pág. 1)" if processo_p1 else "processo da capa não identificado", bool(processo_p1),
                doc_processo or "não encontrado", bool(doc_processo),
                cf._mesmos_digitos(processo_p1, doc_processo) if (processo_p1 and doc_processo) else None,
            ))

            m_proc_contratacao = RE_RONE_PROCESSO_CONTRATACAO.search(texto)
            doc_processo_contratacao = re.sub(r"\s+", "", m_proc_contratacao.group(1)) if m_proc_contratacao else ""
            processo_contratacao_bd = contrato.get("processo_contratacao") if contrato else ""
            linhas.append(cf.linha_tabela(
                "Processo de Contratação",
                f"{processo_contratacao_bd} (BD)" if processo_contratacao_bd else "contrato não encontrado no banco", bool(processo_contratacao_bd),
                doc_processo_contratacao or "não encontrado", bool(doc_processo_contratacao),
                cf._mesmos_digitos(processo_contratacao_bd, doc_processo_contratacao) if (processo_contratacao_bd and doc_processo_contratacao) else None,
            ))

        m_cel = RE_RONE_CEL_ORCAMENTARIA.search(texto)
        doc_ptres = m_cel.group(1) if m_cel else ""
        doc_fonte = m_cel.group(2) if m_cel else ""
        doc_nd = m_cel.group(3) if m_cel else ""
        doc_pi = m_cel.group(4) if m_cel else ""
        linhas.append(_linha_igualdade("PTRES", ref_ptres, doc_ptres, ref_ptres == doc_ptres))
        linhas.append(_linha_igualdade("Fonte", ref_fonte, doc_fonte, ref_fonte == doc_fonte))
        linhas.append(_linha_igualdade(
            "Natureza de Despesa", cf.ns.juntar_com_e(ref_nds) if ref_nds else "",
            doc_nd, doc_nd in ref_nds,
        ))
        linhas.append(_linha_igualdade(
            "PI", cf.ns.juntar_com_e(ref_pis) if ref_pis else "",
            doc_pi, doc_pi in ref_pis,
        ))

        m_clas = RE_RONE_CLAS_ORC.search(texto)
        doc_clas_orc = m_clas.group(1) if m_clas else ""
        linhas.append(_linha_igualdade(
            "CLAS.ORC", cf.ns.juntar_com_e(ref_clas_orc) if ref_clas_orc else "",
            doc_clas_orc, doc_clas_orc in ref_clas_orc,
            rotulo="NC + Dotação",
        ))

        linhas.append(cf.linha_tabela(
            "Valor total empenhado",
            f"{ref_valor} (NC)" if ref_valor else "sem NC anterior", bool(ref_valor),
            soma_texto or "nao encontrado", bool(soma_texto),
            cf._valores_monetarios_batem(ref_valor, soma_valor) if (ref_valor and soma_valor) else None,
        ))

        titulo = f"RO da NE — {ro_numero}" if ro_numero else "RO da NE"
        blocos.append(cf.montar_tabela(nome_arquivo, titulo, idx + 1, linhas))

    return blocos, nes


# ======= Consistencia entre documentos (mesmo dado confrontado entre os proprios documentos) =======
# Mesma ideia do processar_consistencia_documentos da conformidade.py (etapa NS): alem de cada
# documento bater isoladamente contra a fonte segura, a mesma informacao escrita em documentos
# diferentes do mesmo empenho tem que ser igual ENTRE SI. Aqui isso pega, principalmente, uma
# divergencia entre a Dotacao Orcamentaria e o RO da NE quando a Solicitacao de Empenho nao foi
# localizada (sem ela, os dois so eram conferidos cada um contra "sem NC anterior").

# rotulo canonico -> rotulos usados nas tabelas dos documentos que representam o mesmo dado
_CAMPOS_CONSISTENCIA_RO = {
    "Contrato": ["Contrato"],
    "Contratada": ["Contratada", "Favorecido"],
    "CNPJ": ["CNPJ"],
    "Processo": ["Processo"],
    "PTRES": ["PTRES"],
    "Fonte": ["Fonte"],
    "Natureza de Despesa": ["Natureza de Despesa"],
    "PI": ["PI"],
    "Valor": ["Valor", "Valor total empenhado"],
    "CLAS.ORC": ["CLAS.ORC"],
}


def _tokens_orc(texto):
    return {t for t in re.split(r"\s+e\s+|[,;/]\s*|\s+", texto.strip()) if t}


def _conjuntos_orc_compativeis(a, b):
    # ND / PI podem vir com mais de um codigo num documento e so um no outro (ex: a Dotacao lista
    # tambem a natureza generica) - aceita quando um conjunto contem o outro
    ta, tb = _tokens_orc(a), _tokens_orc(b)
    return bool(ta) and bool(tb) and (ta == tb or ta <= tb or tb <= ta)


def _valor_consistencia(texto):
    # "300,00 + 316,67 = 616,67" -> "616,67" (o RO da NE mostra a soma das ROs assim)
    return texto.split(" = ")[-1].split(" → ")[-1].strip()


_COMPARADOR_CONSISTENCIA_RO = {
    "Contrato": cf.comparar_numeros,
    "Contratada": cf.comparar_textos,
    "CNPJ": cf.comparar_cnpjs,
    "Processo": cf.comparar_textos,
    "PTRES": lambda a, b: a.strip() == b.strip(),
    "Fonte": lambda a, b: a.strip() == b.strip(),
    "Natureza de Despesa": _conjuntos_orc_compativeis,
    "PI": _conjuntos_orc_compativeis,
    "Valor": lambda a, b: cf._valores_monetarios_batem(_valor_consistencia(a), _valor_consistencia(b)),
    "CLAS.ORC": cf._mesmos_digitos,
}


def _nome_curto_documento_ro(titulo):
    # o titulo do bloco carrega o numero do RO/despacho no fim - encurta pro nome do documento.
    # a "Solicitacao de Empenho" e a tela SIAFI da NC - rotulada "NC" pra bater com os outros blocos
    if titulo.startswith("Registro Orçamentário"):
        return "NC"
    if titulo.startswith("Dotação Orçamentária"):
        return "Dotação"
    if titulo.startswith("RO da NE"):
        return "RO da NE"
    return titulo


# campos "1 fonte segura x N documentos" (tabela de 2 colunas) x campos orçamentários, que os 3
# documentos (NC/Dotação/RO da NE) relatam cada um com o mesmo peso - por isso viram uma coluna
# por documento em vez de "fonte segura x documento" (pedido do usuário 2026-09-14)
_CAMPOS_DUAS_COLUNAS_RO = ["Contrato", "Contratada", "CNPJ", "Processo"]
_CAMPOS_ORCAMENTARIOS_RO = ["PTRES", "Fonte", "Natureza de Despesa", "PI", "Valor", "CLAS.ORC"]
_COLUNAS_ORCAMENTARIAS_RO = ["RO da NC", "Dotação Orçamentária", "RO da NE"]

# campo orçamentário -> chave em `referencias` (valor lido da própria tela da NC) - CLAS.ORC não
# tem chave própria porque a NC não relata subelemento, só a Dotação (ver _ORC_DOTACAO_VIA_REF)
_ORC_CAMPO_PARA_CHAVE_REF = {
    "PTRES": "ptres", "Fonte": "fonte", "Natureza de Despesa": "nds",
    "PI": "pis", "Valor": "valor", "CLAS.ORC": None,
}
# CLAS.ORC da Dotação não vira "linha" no bloco dela (só entra na observação em vermelho) - o
# valor que ela contribuiu só existe em `referencias` (clas_orc), preenchido por
# processar_dotacao_orcamentaria
_ORC_DOTACAO_VIA_REF = {"CLAS.ORC": "clas_orc"}


def _valor_referencia(ref, chave):
    if not ref or not chave:
        return ""
    valor = ref.get(chave)
    if isinstance(valor, list):
        return cf.ns.juntar_com_e(valor) if valor else ""
    return valor or ""


def _valores_por_documento_ro(campo, rotulos, blocos):
    # nome curto do documento -> lista de valores achados nas linhas dos blocos (pode ter mais de
    # 1 quando há reforço, ex: 2 telas "RO da NE" no mesmo ciclo)
    por_doc = {}
    for bloco in blocos:
        nome = _nome_curto_documento_ro(bloco["documento"])
        for linha in bloco["linhas"]:
            if linha["campo"] in rotulos and linha["documento_disponivel"]:
                valor = linha["documento"]
                if campo == "Valor":
                    valor = _valor_consistencia(valor)
                por_doc.setdefault(nome, []).append(valor)
    return por_doc


def processar_consistencia_documentos_ro(nome_arquivo, blocos, referencias, nota_solicitacao=None):
    # roda DEPOIS dos processadores de documento (precisa dos blocos prontos). So a consistencia
    # orcamentaria (RO da NC x Dotacao Orcamentaria x RO da NE) - Contrato/Contratada/CNPJ/Processo
    # nao entram mais aqui, pedido do usuario 2026-09-14 (ficava redundante com a linha "Processo"
    # que cada documento ja confere individualmente contra a pag. 1 / BD).
    ref = referencias[-1] if referencias else None

    # campos orçamentários: RO da NC x Dotação Orçamentária x RO da NE lado a lado, 1 coluna por
    # documento - nenhum dos 3 é tratado como "fonte segura" dos outros (ver _CAMPOS_ORCAMENTARIOS_RO)
    linhas_orc = []
    for campo in _CAMPOS_ORCAMENTARIOS_RO:
        rotulos = _CAMPOS_CONSISTENCIA_RO[campo]
        por_doc = _valores_por_documento_ro(campo, rotulos, blocos)

        nc_texto = _valor_referencia(ref, _ORC_CAMPO_PARA_CHAVE_REF[campo])
        dotacao_valores = list(dict.fromkeys(por_doc.get("Dotação", [])))
        if not dotacao_valores and campo in _ORC_DOTACAO_VIA_REF:
            dotacao_texto = _valor_referencia(ref, _ORC_DOTACAO_VIA_REF[campo])
        else:
            dotacao_texto = cf.ns.juntar_com_e(dotacao_valores) if dotacao_valores else ""
        rone_valores = list(dict.fromkeys(por_doc.get("RO da NE", [])))
        rone_texto = cf.ns.juntar_com_e(rone_valores) if rone_valores else ""

        valores_colunas = [
            (nc_texto, bool(nc_texto)), (dotacao_texto, bool(dotacao_texto)), (rone_texto, bool(rone_texto)),
        ]
        disponiveis = [texto for texto, disp in valores_colunas if disp]
        if len(disponiveis) < 2:
            continue  # só 1 documento trouxe o campo - nada a confrontar (já aparece no bloco do próprio documento)

        comparador = _COMPARADOR_CONSISTENCIA_RO[campo]
        bate = all(comparador(disponiveis[0], v) for v in disponiveis[1:])
        linhas_orc.append(cf.linha_tabela_multi(campo, valores_colunas, bate))

    blocos_consistencia = []
    if linhas_orc:
        blocos_consistencia.append(cf.montar_tabela(
            nome_arquivo, "Consistência Orçamentária entre Documentos", None, linhas_orc,
            colunas=_COLUNAS_ORCAMENTARIAS_RO, observacao=nota_solicitacao,
        ))
    return blocos_consistencia


# ======= ponto de entrada =======

def gerar_conformidade_ro(nome_arquivo, paginas):
    # roda os processadores de documento pra um PDF de "Andamento do processo" de
    # Solicitacao de empenho. So processa o tipo certo de PDF, pra nao conferir por
    # engano um processo de pagamento (NS), ja coberto pela conformidade.py.
    # Devolve (blocos, resumo); resumo e None quando o PDF nao e do tipo certo, senao
    # {"processo", "nes": [...], "tudo_ok": bool} - usado pra pintar a Planilha de Controle.
    if not paginas or not RE_PAG1_PROCESSO_ELETRONICO.search(paginas[0]):
        return [], None  # nao e um PDF de andamento de processo
    if not RE_PAG1_SOLIC_EMPENHO.search(paginas[0]):
        return [], None  # provavelmente um processo de pagamento (NS) - conferido pela conformidade.py

    # o mesmo processo e reusado o ano todo - so a etapa mais recente e conferida:
    # as paginas depois do ultimo "Despacho: Sem Ocorrencia" / "Certificado de Conformidade Sem Ocorrencia"
    corte = 0
    for i, texto in enumerate(paginas):
        if RE_MARCO_SEM_OCORRENCIA.search(texto):
            corte = i + 1

    # nº do contrato do ciclo, da OBSERVACAO da tela da NC ("SOLICITACAO DE EMPENHO DO CONTRATO
    # N 01/2024 ...") - desambigua a empresa no banco quando ela tem varios contratos (mesmo CNPJ).
    # Ultima ocorrencia = ciclo mais recente
    numero_contrato = ""
    for texto in paginas[corte:]:
        m = RE_CONRO_SOLIC_CONTRATO.search(texto)
        if m:
            numero_contrato = m.group(1)
    contrato = cf.localizar_contrato(paginas, numero_contrato)  # fonte segura: contratada, CNPJ, numero do contrato, empenhos
    m_processo = RE_PROCESSO.search(paginas[0])
    processo_p1 = m_processo.group() if m_processo else None

    req_blocos, requerimentos = processar_requerimento_empenho(nome_arquivo, paginas, contrato, corte)
    solic_blocos, referencias = processar_solicitacao_empenho(nome_arquivo, paginas, contrato, requerimentos, processo_p1, corte)
    dot_blocos = processar_dotacao_orcamentaria(nome_arquivo, paginas, processo_p1, referencias, corte)
    rone_blocos, nes = processar_ro_da_ne(nome_arquivo, paginas, contrato, referencias, corte, processo_p1)

    # exibe na ordem em que os documentos aparecem no processo
    blocos = sorted(req_blocos + solic_blocos + dot_blocos + rone_blocos, key=lambda bloco: bloco["pagina"])

    # confronta a mesma informacao entre os documentos - sempre por ultimo (blocos sem pagina)
    nota_solicitacao = _nota_solicitacao_servico(paginas, corte)
    blocos_consistencia = processar_consistencia_documentos_ro(nome_arquivo, blocos, referencias, nota_solicitacao)
    blocos.extend(blocos_consistencia)

    # "toda conferencia bater" = pelo menos um documento conferido e TODA linha "ok"
    # (nenhuma divergencia e nenhum campo que nao deu pra comparar)
    tudo_ok = bool(blocos) and all(
        linha["resultado"] == "ok" for bloco in blocos for linha in bloco["linhas"]
    )
    resumo = {"processo": processo_p1, "nes": nes, "tudo_ok": tudo_ok}
    return blocos, resumo


def rodar_conferencia_ro():
    # usado tanto pelo main()/CLI quanto pela janela - devolve (blocos, aviso, resumos, num_pdfs):
    # blocos = todos os documentos conferidos; resumos = um por PDF valido, pra pintar a planilha;
    # num_pdfs = quantos PDFs a coleta encontrou (Chrome + abertos + caixa de diálogo), pro print
    contratos_db.inicializar_db()
    fontes, aviso = cf.coletar_fontes_pdf()
    blocos, resumos = [], []
    for nome_exibicao, paginas in fontes:
        blocos_pdf, resumo = gerar_conformidade_ro(nome_exibicao, paginas)
        blocos.extend(blocos_pdf)
        if resumo:
            resumos.append(resumo)
    return blocos, aviso, resumos, len(fontes)


def pintar_empenhos_aprovados(nome_planilha, resumos):
    # quando TODA a conferencia de um processo bate, pinta a celula do empenho (coluna "NE"
    # da aba "RO") de amarelo claro 1 - mesma sinalizacao dos outros scripts do pipeline - e
    # preenche a coluna "DESPACHO" da linha com "Sem ocorrência" (texto usado depois pelo
    # baixar_anexar_ne.py pra tramitar o processo com esse despacho)
    aprovados = [r for r in resumos if r["tudo_ok"] and r["nes"]]
    if not aprovados:
        return

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES)
    gc = gspread.authorize(credenciais)
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    aba = planilha.worksheet("RO")

    cabecalho = aba.row_values(1)
    col_ne = cabecalho.index("NE") + 1 if "NE" in cabecalho else 8  # coluna H
    col_processo = cabecalho.index("PROCESSO") + 1 if "PROCESSO" in cabecalho else 1  # coluna A
    col_despacho = cabecalho.index("DESPACHO") + 1 if "DESPACHO" in cabecalho else None

    valores_ne = aba.col_values(col_ne)
    valores_processo = aba.col_values(col_processo)
    valores_despacho = aba.col_values(col_despacho) if col_despacho else []

    for resumo in aprovados:
        for i, valor in enumerate(valores_ne[1:], start=2):  # pula o cabecalho; i = nº da linha
            if valor.strip() not in resumo["nes"]:
                continue
            processo_linha = valores_processo[i - 1].strip() if i - 1 < len(valores_processo) else ""
            if resumo["processo"] and processo_linha and processo_linha != resumo["processo"]:
                continue  # mesma NE em outro processo - nao pinta
            pintar_celula_planilha.executar(aba, i, col_ne, AMARELO_CLARO_1)

            extra = ""
            if col_despacho:
                despacho_atual = valores_despacho[i - 1].strip() if i - 1 < len(valores_despacho) else ""
                if not despacho_atual:
                    aba.update_cell(i, col_despacho, "Sem ocorrência")
                    extra = ' e DESPACHO preenchido com "Sem ocorrência"'
                elif despacho_atual.lower() != "sem ocorrência":
                    extra = f' (DESPACHO já tinha "{despacho_atual}", mantido)'

            print(f"{resumo['processo'] or valor} - conferencia OK, empenho {valor.strip()} "
                  f"pintado de amarelo na planilha (linha {i}){extra}.")


def abrir_janela(blocos, aviso):
    # reaproveita o HTML e a ApiConformidade da janela de resultado da conformidade.py -
    # so muda o titulo da janela do SO pra deixar claro que e a conferencia de RO
    x, y, largura, altura = janela_windows.geometria_para_tela(1040, 780)
    webview.create_window(
        "CCRGCI - Resultado da Conformidade (RO)", html=cf.HTML_CONFORMIDADE,
        js_api=cf.ApiConformidade(blocos, aviso), width=largura, height=altura, x=x, y=y,
    )


def main(nome_planilha=None):
    # roda pelo card "Fazer Conformidade (RO)" do gui.py (background thread) - ao terminar,
    # abre a janela de resultado automaticamente, mesmo padrao da conformidade.py
    blocos, aviso, resumos, num_pdfs = rodar_conferencia_ro()
    print(f"PDFs baixados e abertos encontrados: {num_pdfs}")
    print("Abrindo janela com o resultado...")

    # pinta a celula do empenho na Planilha de Controle pros processos 100% conferidos
    # (so quando rodando pelo gui.py, que passa a planilha escolhida)
    if nome_planilha:
        try:
            pintar_empenhos_aprovados(nome_planilha, resumos)
        except Exception as e:
            print(f"Nao foi possivel pintar a Planilha de Controle: {e}")

    abrir_janela(blocos, aviso)


if __name__ == "__main__":
    main()
    webview.start()
