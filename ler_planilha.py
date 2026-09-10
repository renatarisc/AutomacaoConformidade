import pandas as pd


def _ranges_mesclados(aba):
    # faixas de células mescladas da aba, em índices 0-based com fim exclusivo (formato da API do
    # Sheets: {"startRowIndex", "endRowIndex", "startColumnIndex", "endColumnIndex"})
    meta = aba.spreadsheet.fetch_sheet_metadata()
    for s in meta.get("sheets", []):
        if s.get("properties", {}).get("sheetId") == aba.id:
            return s.get("merges", [])
    return []


def _preencher_mescladas(valores, merges):
    # o get_all_values() traz o conteúdo de um bloco mesclado só na célula superior-esquerda; as
    # demais vêm "". Replica esse valor para todo o bloco, para a leitura tabular não perder dado
    # (ex: coluna "Processo" mesclada verticalmente cobrindo várias linhas de diárias).
    for m in merges:
        r0, r1 = m.get("startRowIndex", 0), m.get("endRowIndex", 0)
        c0, c1 = m.get("startColumnIndex", 0), m.get("endColumnIndex", 0)
        if r0 >= len(valores) or c0 >= len(valores[r0]):
            continue
        valor = valores[r0][c0]
        if valor == "":
            continue
        for r in range(r0, min(r1, len(valores))):
            for c in range(c0, min(c1, len(valores[r]))):
                if valores[r][c] == "":
                    valores[r][c] = valor
    return valores


def carregar_registros(aba, preencher_mescladas=False):
    # Substitui aba.get_all_records(numericise_ignore=['all']): lê a 1ª linha como
    # cabeçalho e o restante como texto (sem numericizar, pra não quebrar valores no
    # formato brasileiro). Diferença: ignora as colunas cujo cabeçalho está em branco.
    #
    # O get_all_records() do gspread pad-eia todas as linhas para o mesmo tamanho, então
    # qualquer célula solta à direita (ou um cabeçalho com buracos) gera várias colunas
    # com nome '' e ele aborta com "the header row in the worksheet contains duplicates: ['']".
    #
    # preencher_mescladas=True: faz uma chamada extra (fetch_sheet_metadata) e replica o valor
    # de cada célula mesclada para todas as linhas/colunas do bloco - necessário quando uma
    # coluna lida por nome (ex: "Processo") aparece mesclada verticalmente na planilha.
    valores = aba.get_all_values()
    if not valores:
        return pd.DataFrame()

    if preencher_mescladas:
        valores = _preencher_mescladas(valores, _ranges_mesclados(aba))

    cabecalho = valores[0]
    indices = [i for i, nome in enumerate(cabecalho) if str(nome).strip() != ""]
    colunas = [cabecalho[i] for i in indices]
    linhas = [
        [linha[i] if i < len(linha) else "" for i in indices]
        for linha in valores[1:]
    ]
    return pd.DataFrame(linhas, columns=colunas)
