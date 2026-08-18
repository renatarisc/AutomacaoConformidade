# faz uma única chamada à API e monta um dicionário com as cores de todas as células da planilha
def executar(var_aba):

    metadata = var_aba.spreadsheet.fetch_sheet_metadata({"includeGridData": True})

    sheet = next(
        s for s in metadata["sheets"]
        if s["properties"]["sheetId"] == var_aba.id)

    row_data = sheet["data"][0]["rowData"]
    cores = {}

    for i, linha in enumerate(row_data):
        for j, celula in enumerate(linha.get("values", [])):
            cor = (celula.get("userEnteredFormat", {}).get("backgroundColor", {}))
            cores[(i + 1, j + 1)] = (cor.get("red", 0), cor.get("green", 0), cor.get("blue", 0))
    return cores
