import json
import os
import re
import subprocess
import time
from pathlib import Path

try:
    import winreg # Windows-only (stdlib) - usado só pra resolver as pastas Desktop/Documents reais
except ImportError:
    winreg = None

# encontra PDFs abertos em QUALQUER visualizador no Windows (Foxit, Adobe, Edge, SumatraPDF,
# etc.) - existe como alternativa ao fluxo normal (abas do Chrome na tela djtools/
# process_progress2) para quando o PDF já foi baixado e está aberto direto num visualizador,
# sem passar pelo Chrome. Não depende de biblioteca nova (pywin32 etc.) - só invoca PowerShell,
# que já vem em qualquer Windows, e pega a linha de comando do processo via WMI pra extrair o
# caminho do arquivo (o título da janela geralmente só tem o nome do arquivo, não o caminho)
# a 1ª linha força o PowerShell a escrever a saída em UTF-8 - sem isso, um caminho com acento
# (ex: "...\OneDrive\Área de Trabalho\arquivo.pdf") volta noutro encoding, o Python decodifica
# errado, "Área" vira "µrea" e o os.path.exists() lá embaixo falha, descartando o PDF em silêncio
_PS_LISTAR_JANELAS_PDF = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
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
            capture_output=True, timeout=15,
        )
    except (OSError, subprocess.SubprocessError):
        return []

    # decodifica na mão como UTF-8 (ver _PS_LISTAR_JANELAS_PDF) - text=True usaria o encoding da
    # locale (cp1252) e corromperia caminho com acento
    saida = resultado.stdout.decode("utf-8", errors="replace").strip()
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

# valores em HKCU\...\User Shell Folders - resolvem o caminho REAL de Downloads/Desktop/Documents
# mesmo quando redirecionados pro OneDrive ou com nome localizado ("Área de Trabalho")
_SHELL_FOLDERS = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
_CHAVES_PASTAS = ("{374DE290-123F-4565-9164-39C4925E467B}", "Desktop", "Personal") # Downloads, Desktop, Documents

def _pastas_padrao():
    # Downloads + Desktop + Documents, na ordem - varridas juntas porque o usuário costuma mover o
    # PDF baixado pra Área de Trabalho antes de rodar a conferência
    pastas = []
    if winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _SHELL_FOLDERS) as chave:
                for nome in _CHAVES_PASTAS:
                    try:
                        bruto, _ = winreg.QueryValueEx(chave, nome)
                        pastas.append(Path(os.path.expandvars(bruto)))
                    except OSError:
                        pass
        except OSError:
            pass
    if not pastas: # fallback se o registro não colaborar
        pastas = [Path.home() / "Downloads", Path.home() / "Desktop", Path.home() / "Documents"]
    vistos, unicas = set(), []
    for p in pastas:
        if p not in vistos:
            vistos.add(p)
            unicas.append(p)
    return unicas

def listar_pdfs_recentes(pasta=None, dias=DIAS_RECENTES_PADRAO, limite=LIMITE_ARQUIVOS_PADRAO):
    # varre pastas (por padrão Downloads + Desktop + Documents) por PDFs modificados nos últimos
    # `dias` dias - complementa listar_pdfs_abertos(), que só enxerga a aba em primeiro plano
    # quando o visualizador usa abas dentro de uma única janela (ex: Foxit) - não tem como saber
    # quais abas estão "abertas" nesse caso de fora do programa. Em vez disso, considera candidato
    # qualquer PDF recente; a validação de conteúdo (extrair_dados, no script que chama isso)
    # descarta na hora qualquer um que não seja realmente um PDF de Andamento do processo, então
    # um PDF errado numa dessas pastas não vira problema, só é ignorado
    pastas = [Path(pasta)] if pasta else _pastas_padrao()
    limite_tempo = time.time() - dias * 86400

    candidatos = []
    for p in pastas:
        if not p.is_dir():
            continue
        for arquivo in p.glob("*.pdf"):
            try:
                if arquivo.is_file() and arquivo.stat().st_mtime >= limite_tempo:
                    candidatos.append(arquivo)
            except OSError:
                pass
    candidatos.sort(key=lambda arquivo: arquivo.stat().st_mtime, reverse=True)

    vistos, resultado = set(), []
    for arquivo in candidatos:
        chave = str(arquivo).lower()
        if chave not in vistos:
            vistos.add(chave)
            resultado.append(str(arquivo))
        if len(resultado) >= limite:
            break
    return resultado

if __name__ == "__main__":
    for caminho in listar_pdfs_abertos():
        print("aberto:", caminho)
    for caminho in listar_pdfs_recentes():
        print("recente:", caminho)
