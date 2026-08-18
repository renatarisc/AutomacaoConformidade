def executar(aba, linha, coluna, cor): # pinta uma célula do Google Sheets (na Planilha de Controle)
    # linha: nº da linha na planilha (começa em 1) | coluna: nº da coluna (A=1) | cor: tupla RGB, verde = (0, 1, 0)

    # converte número da coluna para letra
    def numero_coluna_para_letra(n):
        letra = ""
        while n:
            n, resto = divmod(n - 1, 26)
            letra = chr(65 + resto) + letra
        return letra

    letra_coluna = numero_coluna_para_letra(coluna)
    aba.format(f"{letra_coluna}{linha}", {"backgroundColor": {"red": cor[0], "green": cor[1], "blue": cor[2]}})