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

# amarelo claro 1 - mesmo amarelo padrao usado no relacionar_valor_op.py / baixar_ob.py / etc.
AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255)

# a extracao de texto desse tipo de PDF costuma trocar acento por "?" (U+FFFD) - os
# marcadores textuais usam "." curinga no lugar da letra acentuada
RE_PAG1_PROCESSO_ELETRONICO = re.compile(r"Processo\s+Eletr.nico")
RE_PAG1_SOLIC_EMPENHO = re.compile(r"Solicita..o de empenho")
# marco que fecha um ciclo de empenho: o "Despacho: Sem ocorrencia" da conformidade OU o
# "Certificado de Conformidade Sem Ocorrencia" (alguns processos trazem o certificado no lugar
# do despacho) - o recorte da etapa mais recente comeca depois da ultima ocorrencia de qualquer um
RE_MARCO_SEM_OCORRENCIA = re.compile(
    r"(?:Despacho:\s*|Certificado\s+de\s+Conformidade\s+)Sem\s+ocorr.ncia", re.IGNORECASE
)

RE_CONRO_TITULO = re.compile(r"CONSULTA-CONRO|CONSULTA REGISTRO ORCAMENTARIO")
RE_CONRO_RO_NUMERO = re.compile(r"NUMERO\s*:\s*(2026RO\d+)")


def _somar_valores(valores):
    total = sum(cf._valor_para_float(v) for v in valores)
    return cf._float_para_valor_br(total)


def _natureza_empenho_bd(contrato, numero_empenho):
    # natureza de despesa (ex: "339039-25") cadastrada pro empenho especifico no banco -
    # "" se o contrato/empenho nao esta cadastrado ou o empenho nao tem ND
    if not contrato:
        return ""
    empenhos = list(contrato.get("empenhos", []))
    for processo in contrato.get("processos_empenho", []):
        empenhos.extend(processo.get("empenhos", []))
    for e in empenhos:
        if e.get("numero_empenho") == numero_empenho:
            return e.get("natureza_despesa") or ""
    return ""


# ======= Documento 1: Solicitacao de Empenho (tela SIAFI CONSULTA-CONRO da NC) =======

# "SOLICITACAO DE EMPENHO DO CONTRATO N 17/2023 COM A EMPRESA PRIME ..., REFERENTE ..." ou, quando
# o empenho e de uma contratacao direta (ex: fornecimento de energia), "SOLICITACAO DE EMPENHO DA
# CONTRATACAO 90037/2025 COM A EMPRESA ..." (sem o "N", nº com 5 digitos = pregao/contratacao)
RE_CONRO_SOLIC_CONTRATO = re.compile(
    r"SOLICITA..O DE EMPENHO D[AO]\s+(?:CONTRATO|CONTRATA..O)\s+(?:N.?\s*)?(\d+/\d{4})"
)
RE_CONRO_DOC_WEB_NC = re.compile(r"DOCUMENTO WEB\s*:\s*(2026NC\d+)")
RE_CONRO_EMPRESA = re.compile(r"COM A EMPRESA\s+(.+?)\s*,\s*REFERENTE", re.DOTALL)
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


def processar_solicitacao_empenho(nome_arquivo, paginas, contrato, corte):
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

        ptres, fonte, nds, pis = _celula_orcamentaria(texto)
        valor = _valor_total_conro(texto)

        partes = []
        if ptres:
            partes.append(f"PTRES: {ptres}")
        if fonte:
            partes.append(f"Fonte: {fonte}")
        if nds:
            partes.append(f"ND: {cf.ns.juntar_com_e(nds)}")
        if pis:
            partes.append(f"PI: {cf.ns.juntar_com_e(pis)}")
        if valor:
            partes.append(f"Valor: {valor}")

        # nome no padrao "Registro Orçamentário <RO>-<NC>"
        partes_titulo = "-".join(p for p in (ro_numero, nc_numero) if p)
        titulo = f"Registro Orçamentário {partes_titulo}" if partes_titulo else "Registro Orçamentário"
        blocos.append(cf.montar_tabela(
            nome_arquivo, titulo, idx + 1, linhas,
            observacao=" | ".join(partes) if partes else None,
        ))
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
RE_DOT_SUBELEMENTO = re.compile(r"\b(\d{6})\s*[-‐-―]\s*(\d{2})\b")


def _linha_igualdade(campo, esperado, obtido, igual):
    # esperado (fonte segura) vem da NC (Registro Orcamentario) anterior; obtido, deste documento
    tem_ref = bool(esperado)
    return cf.linha_tabela(
        campo,
        (esperado + " (NC)") if tem_ref else "sem NC anterior", tem_ref,
        obtido or "nao encontrado", bool(obtido),
        igual if (tem_ref and obtido) else None,
    )


def processar_dotacao_orcamentaria(nome_arquivo, paginas, processo_p1, referencias, corte):
    blocos = []

    for idx in range(corte, len(paginas)):
        texto_pag = paginas[idx]
        if not RE_DOT_ASSUNTO.search(texto_pag):
            continue
        texto = cf.remover_duplicatas_consecutivas(texto_pag)  # esse despacho vem com cada linha duplicada

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
        # a CLAS.ORC contra, e mostra no bloco vermelho
        pares_sub = list(dict.fromkeys(RE_DOT_SUBELEMENTO.findall(texto)))  # [(nd, sub), ...]
        if ref is not None:
            for nd, sub in pares_sub:
                if nd + sub not in ref["clas_orc"]:
                    ref["clas_orc"].append(nd + sub)
        # o subelemento e so o que vem depois do traco ("339039-25" -> "25")
        subelementos = list(dict.fromkeys(sub for _, sub in pares_sub))
        observacao = f"Subelemento: {cf.ns.juntar_com_e(subelementos)}" if subelementos else None

        m_desp = RE_DOT_DESPACHO.search(texto)
        titulo = f"Dotação Orçamentária — Despacho {m_desp.group(1)}" if m_desp else "Dotação Orçamentária"
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


def processar_ro_da_ne(nome_arquivo, paginas, contrato, referencias, corte):
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

    empenhos_bd = cf.empenhos_registrados(contrato) if contrato else []
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
        ))

        linhas.append(cf.linha_tabela(
            "Valor total empenhado",
            f"{ref_valor} (NC)" if ref_valor else "sem NC anterior", bool(ref_valor),
            soma_texto or "nao encontrado", bool(soma_texto),
            cf._valores_monetarios_batem(ref_valor, soma_valor) if (ref_valor and soma_valor) else None,
        ))

        m_ne = RE_RONE_DOC_WEB_NE.search(texto)
        ne = m_ne.group(1) if m_ne else ""
        m_valor_ro = RE_RONE_VALOR_EVENTO.search(texto)
        valor_ro = m_valor_ro.group(1) if m_valor_ro else ""
        partes = []
        if valor_ro:
            partes.append(f"Valor desta RO: {valor_ro}")
        if ne:
            if ne in empenhos_bd:
                nd_bd = _natureza_empenho_bd(contrato, ne)
                # ND do documento = CLAS.ORC (natureza + subelemento, ex "33903925"); cai pra
                # ND da celula orcamentaria + subelemento da Dotacao (ref_clas_orc) se faltar
                nd_doc = doc_clas_orc or (ref_clas_orc[0] if ref_clas_orc else "")
                if nd_bd and nd_doc and not cf._mesmos_digitos(nd_bd, nd_doc):
                    situacao = (f"Empenho {ne} — já cadastrado no BD, mas a ND {nd_bd} "
                                f"não confere com o documento ({nd_doc})")
                elif nd_bd:
                    situacao = f"Empenho {ne} e ND {nd_bd} — já cadastrado no BD"
                else:
                    situacao = f"Empenho {ne} — já cadastrado no BD, sem ND cadastrada"
            else:
                situacao = f"Empenho {ne} — não cadastrado no BD"
            partes.append(situacao)

        titulo = f"RO da NE — {ro_numero}" if ro_numero else "RO da NE"
        blocos.append(cf.montar_tabela(
            nome_arquivo, titulo, idx + 1, linhas,
            observacao=" | ".join(partes) if partes else None,
        ))

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


def processar_consistencia_documentos_ro(nome_arquivo, blocos, referencias, contrato, processo_p1):
    # roda DEPOIS dos processadores de documento (precisa dos blocos prontos). A Solicitacao de
    # Empenho (tela SIAFI da NC) guarda PTRES/Fonte/ND/PI/Valor na observacao (nao como linha),
    # entao entra aqui so como FONTE segura, via 'referencias' - a ultima do recorte e a referencia.
    # Rotulo "(NC)" (nao "(Solicitação)") pra bater com os blocos da Dotacao / RO da NE
    ref = referencias[-1] if referencias else None
    cnpj_bd = cf._formatar_cnpj(contrato["cnpj"]) if contrato and contrato.get("cnpj") else ""

    fontes = {
        "Contrato": (f"{contrato['numero_contrato']} (BD)" if contrato else "contrato não encontrado no banco",
                     bool(contrato), contrato["numero_contrato"] if contrato else ""),
        "Contratada": (f"{contrato['nome_contratada']} (BD)" if contrato else "contrato não encontrado no banco",
                       bool(contrato), contrato["nome_contratada"] if contrato else ""),
        "CNPJ": (f"{cnpj_bd} (BD)" if cnpj_bd else "contrato não encontrado no banco", bool(cnpj_bd), cnpj_bd),
        "Processo": (f"{processo_p1} (pág. 1)" if processo_p1 else "não encontrado", bool(processo_p1), processo_p1 or ""),
        "PTRES": (f"{ref['ptres']} (NC)" if ref and ref["ptres"] else "sem NC anterior",
                  bool(ref and ref["ptres"]), ref["ptres"] if ref else ""),
        "Fonte": (f"{ref['fonte']} (NC)" if ref and ref["fonte"] else "sem NC anterior",
                  bool(ref and ref["fonte"]), ref["fonte"] if ref else ""),
        "Natureza de Despesa": (
            f"{cf.ns.juntar_com_e(ref['nds'])} (NC)" if ref and ref["nds"] else "sem NC anterior",
            bool(ref and ref["nds"]), cf.ns.juntar_com_e(ref["nds"]) if ref and ref["nds"] else ""),
        "PI": (
            f"{cf.ns.juntar_com_e(ref['pis'])} (NC)" if ref and ref["pis"] else "sem NC anterior",
            bool(ref and ref["pis"]), cf.ns.juntar_com_e(ref["pis"]) if ref and ref["pis"] else ""),
        "Valor": (f"{ref['valor']} (NC)" if ref and ref["valor"] else "sem NC anterior",
                  bool(ref and ref["valor"]), ref["valor"] if ref else ""),
        "CLAS.ORC": (
            f"{cf.ns.juntar_com_e(ref['clas_orc'])} (Dotação)" if ref and ref["clas_orc"] else "sem subelemento na Dotação",
            bool(ref and ref["clas_orc"]), cf.ns.juntar_com_e(ref["clas_orc"]) if ref and ref["clas_orc"] else ""),
    }

    linhas = []
    for campo, rotulos in _CAMPOS_CONSISTENCIA_RO.items():
        ocorrencias = []
        for bloco in blocos:
            for linha in bloco["linhas"]:
                if linha["campo"] in rotulos and linha["documento_disponivel"]:
                    valor = linha["documento"]
                    if campo == "Valor":
                        valor = _valor_consistencia(valor)
                    ocorrencias.append((_nome_curto_documento_ro(bloco["documento"]), valor))

        if not ocorrencias:
            continue

        fonte_texto, fonte_disp, fonte_valor = fontes[campo]
        comparador = _COMPARADOR_CONSISTENCIA_RO[campo]
        if fonte_valor:
            bate = all(comparador(fonte_valor, valor) for _, valor in ocorrencias)
        elif len(ocorrencias) >= 2:
            referencia = ocorrencias[0][1]
            bate = all(comparador(referencia, valor) for _, valor in ocorrencias)
        else:
            continue  # 1 documento e sem fonte segura - nada a confrontar (ja aparece no bloco do proprio documento)

        # bate -> mostra so o valor (ja e o mesmo em todos); nao bate -> discrimina por documento
        doc_texto = ocorrencias[0][1] if bate else " | ".join(f"{nome}: {valor}" for nome, valor in ocorrencias)
        linhas.append(cf.linha_tabela(campo, fonte_texto, fonte_disp, doc_texto, True, bate))

    if not linhas:
        return None
    return cf.montar_tabela(nome_arquivo, "Consistência entre Documentos", None, linhas)


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

    solic_blocos, referencias = processar_solicitacao_empenho(nome_arquivo, paginas, contrato, corte)
    dot_blocos = processar_dotacao_orcamentaria(nome_arquivo, paginas, processo_p1, referencias, corte)
    rone_blocos, nes = processar_ro_da_ne(nome_arquivo, paginas, contrato, referencias, corte)

    # exibe na ordem em que os documentos aparecem no processo
    blocos = sorted(solic_blocos + dot_blocos + rone_blocos, key=lambda bloco: bloco["pagina"])

    # confronta a mesma informacao entre os documentos - sempre por ultimo (bloco sem pagina)
    consistencia = processar_consistencia_documentos_ro(nome_arquivo, blocos, referencias, contrato, processo_p1)
    if consistencia:
        blocos.append(consistencia)

    # "toda conferencia bater" = pelo menos um documento conferido e TODA linha "ok"
    # (nenhuma divergencia e nenhum campo que nao deu pra comparar)
    tudo_ok = bool(blocos) and all(
        linha["resultado"] == "ok" for bloco in blocos for linha in bloco["linhas"]
    )
    resumo = {"processo": processo_p1, "nes": nes, "tudo_ok": tudo_ok}
    return blocos, resumo


def rodar_conferencia_ro():
    # usado tanto pelo main()/CLI quanto pela janela - devolve (blocos, aviso, resumos):
    # blocos = todos os documentos conferidos; resumos = um por PDF valido, pra pintar a planilha
    contratos_db.inicializar_db()
    fontes, aviso = cf.coletar_fontes_pdf()
    blocos, resumos = [], []
    for nome_exibicao, paginas in fontes:
        blocos_pdf, resumo = gerar_conformidade_ro(nome_exibicao, paginas)
        blocos.extend(blocos_pdf)
        if resumo:
            resumos.append(resumo)
    return blocos, aviso, resumos


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
    blocos, aviso, resumos = rodar_conferencia_ro()
    if aviso:
        print(aviso)
    print("Resultado da Conformidade (RO):")
    if not blocos:
        print("Nenhum documento para conferencia encontrado nos PDFs disponiveis.")
    else:
        print(f"{len(blocos)} documento(s) conferido(s) - Abrindo janela com o resultado...")

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
