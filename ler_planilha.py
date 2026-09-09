import pandas as pd


def carregar_registros(aba):
    # Substitui aba.get_all_records(numericise_ignore=['all']): lê a 1ª linha como
    # cabeçalho e o restante como texto (sem numericizar, pra não quebrar valores no
    # formato brasileiro). Diferença: ignora as colunas cujo cabeçalho está em branco.
    #
    # O get_all_records() do gspread pad-eia todas as linhas para o mesmo tamanho, então
    # qualquer célula solta à direita (ou um cabeçalho com buracos) gera várias colunas
    # com nome '' e ele aborta com "the header row in the worksheet contains duplicates: ['']".
    valores = aba.get_all_values()
    if not valores:
        return pd.DataFrame()

    cabecalho = valores[0]
    indices = [i for i, nome in enumerate(cabecalho) if str(nome).strip() != ""]
    colunas = [cabecalho[i] for i in indices]
    linhas = [
        [linha[i] if i < len(linha) else "" for i in indices]
        for linha in valores[1:]
    ]
    return pd.DataFrame(linhas, columns=colunas)
