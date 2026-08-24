import json
import os
import re
import subprocess
import time
from pathlib import Path

# encontra PDFs abertos em QUALQUER visualizador no Windows (Foxit, Adobe, Edge, SumatraPDF,
# etc.) - existe como alternativa ao fluxo normal (abas do Chrome na tela djtools/
# process_progress2) para quando o PDF já foi baixado e está aberto direto num visualizador,
# sem passar pelo Chrome. Não depende de biblioteca nova (pywin32 etc.) - só invoca PowerShell,
# que já vem em qualquer Windows, e pega a linha de comando do processo via WMI pra extrair o
# caminho do arquivo (o título da janela geralmente só tem o nome do arquivo, não o caminho)
_PS_LISTAR_JANELAS_PDF = r"""
Get-Process | Where-Object { $_.MainWindowTitle -match '\.pdf' } | ForEach-Object {
    $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId = $($_.Id)").CommandLine
    [PSCustomObject]@{ Titulo = $_.MainWindowTitle; ComandoLinha = $cmd }
} | ConvertTo-Json -Compress
"""

# pega o primeiro argumento da linha de comando que termine em .pdf, entre aspas ou não
RE_CAMINHO_PDF = re.compile(r'"([^"]+\.pdf)"|(\S+\.pdf)', re.IGNORECASE)

def listar_pdfs_abertos():
    # devolve uma lista de caminhos (sem repetir) dos PDFs atualmente abertos em visualizadores
    # no Windows - lista vazia se não achar nenhum ou se não conseguir rodar o PowerShell
    try:
        resultado = subprocess.run(
            ["powershell", "-NoProfile", "-Command", _PS_LISTAR_JANELAS_PDF],
            capture_output=True, text=True, timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return []

    saida = resultado.stdout.strip()
    if not saida:
        return []

    try:
        dados = json.loads(saida)
    except json.JSONDecodeError:
        return []

    if isinstance(dados, dict): # o PowerShell devolve um objeto único (não uma lista) quando só há 1 resultado
        dados = [dados]

    caminhos = []
    for item in dados:
        comando = item.get("ComandoLinha") or ""
        match = RE_CAMINHO_PDF.search(comando)
        if not match:
            continue
        caminho = match.group(1) or match.group(2)
        if os.path.exists(caminho) and caminho not in caminhos:
            caminhos.append(caminho)
    return caminhos

DIAS_RECENTES_PADRAO = 7
LIMITE_ARQUIVOS_PADRAO = 50

def listar_pdfs_recentes(pasta=None, dias=DIAS_RECENTES_PADRAO, limite=LIMITE_ARQUIVOS_PADRAO):
    # varre uma pasta (por padrão, Downloads) por PDFs modificados nos últimos `dias` dias -
    # complementa listar_pdfs_abertos(), que só enxerga a aba em primeiro plano quando o
    # visualizador usa abas dentro de uma única janela (ex: Foxit) - não tem como saber quais
    # abas estão "abertas" nesse caso de fora do programa. Em vez disso, considera candidato
    # qualquer PDF baixado recentemente; a validação de conteúdo (extrair_dados, no script que
    # chama isso) descarta na hora qualquer um que não seja realmente um PDF de Andamento do
    # processo, então um PDF errado na pasta não vira um problema, só é ignorado
    pasta = Path(pasta) if pasta else (Path.home() / "Downloads")
    if not pasta.is_dir():
        return []

    limite_tempo = time.time() - dias * 86400
    candidatos = [
        arquivo for arquivo in pasta.glob("*.pdf")
        if arquivo.is_file() and arquivo.stat().st_mtime >= limite_tempo
    ]
    candidatos.sort(key=lambda arquivo: arquivo.stat().st_mtime, reverse=True)
    return [str(arquivo) for arquivo in candidatos[:limite]]

if __name__ == "__main__":
    for caminho in listar_pdfs_abertos():
        print("aberto:", caminho)
    for caminho in listar_pdfs_recentes():
        print("recente:", caminho)
