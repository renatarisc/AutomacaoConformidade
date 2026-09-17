import contextlib
import csv
import os
import shutil
import sqlite3
from datetime import datetime

# mesmo padrão de contratos_db.py: sqlite nativo (sem dependência externa), banco e backups
# próprios (não compartilha arquivo com contratos.db) - ver [[project-menu-and-packaging]]
CAMINHO_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bolsas.db")
PASTA_BACKUPS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bolsas_backups")
MAX_BACKUPS = 300

_ESQUEMA_SQL = """
CREATE TABLE IF NOT EXISTS bolsas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_bolsa TEXT NOT NULL,
    processo_anual TEXT,
    edital TEXT,
    -- texto livre, não número: às vezes há mais de um valor unitário pra anotar (ex:
    -- "250,00 e 300,00" quando o edital reajustou no meio do ano)
    valor_unitario TEXT,
    empenho TEXT,
    criado_em TEXT NOT NULL DEFAULT (datetime('now')),
    atualizado_em TEXT NOT NULL DEFAULT (datetime('now'))
);
"""

@contextlib.contextmanager
def _conexao():
    conexao = sqlite3.connect(CAMINHO_DB)
    conexao.row_factory = sqlite3.Row
    try:
        with conexao:
            yield conexao
    finally:
        conexao.close()

def inicializar_db():
    fazer_backup()  # snapshot do estado atual ANTES de qualquer criação/migração de tabela
    with _conexao() as conexao:
        conexao.executescript(_ESQUEMA_SQL)

# ---------- backup (mesmo esquema de contratos_db.py: cópia binária + CSV a cada gravação) ----------

def fazer_backup():
    if not os.path.exists(CAMINHO_DB):
        return None

    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pasta_destino = os.path.join(PASTA_BACKUPS, carimbo)
    os.makedirs(pasta_destino, exist_ok=True)

    shutil.copy2(CAMINHO_DB, os.path.join(pasta_destino, "bolsas.db"))

    conexao = sqlite3.connect(CAMINHO_DB)
    conexao.row_factory = sqlite3.Row
    try:
        linhas = conexao.execute("SELECT * FROM bolsas").fetchall()
        colunas = (
            list(linhas[0].keys()) if linhas
            else [descricao[0] for descricao in conexao.execute("SELECT * FROM bolsas LIMIT 0").description]
        )
        with open(os.path.join(pasta_destino, "bolsas.csv"), "w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo)
            escritor.writerow(colunas)
            for linha in linhas:
                escritor.writerow([linha[coluna] for coluna in colunas])
    finally:
        conexao.close()

    _rotacionar_backups()
    return pasta_destino

def _rotacionar_backups():
    if not os.path.isdir(PASTA_BACKUPS):
        return
    pastas = sorted(
        nome for nome in os.listdir(PASTA_BACKUPS)
        if os.path.isdir(os.path.join(PASTA_BACKUPS, nome))
    )
    for nome in pastas[:-MAX_BACKUPS] if len(pastas) > MAX_BACKUPS else []:
        shutil.rmtree(os.path.join(PASTA_BACKUPS, nome), ignore_errors=True)

# ---------- CRUD ----------

def listar_bolsas():
    with _conexao() as conexao:
        linhas = conexao.execute(
            "SELECT * FROM bolsas ORDER BY tipo_bolsa COLLATE NOCASE, edital COLLATE NOCASE"
        ).fetchall()
        return [dict(linha) for linha in linhas]

def obter_bolsa(bolsa_id):
    with _conexao() as conexao:
        linha = conexao.execute("SELECT * FROM bolsas WHERE id = ?", (bolsa_id,)).fetchone()
        return dict(linha) if linha else None

_COLUNAS_BOLSA = ("tipo_bolsa", "processo_anual", "edital", "valor_unitario", "empenho")

def _valores_colunas(dados):
    return [dados.get(coluna) for coluna in _COLUNAS_BOLSA]

def criar_bolsa(dados):
    with _conexao() as conexao:
        marcadores = ", ".join("?" for _ in _COLUNAS_BOLSA)
        cursor = conexao.execute(
            f"INSERT INTO bolsas ({', '.join(_COLUNAS_BOLSA)}) VALUES ({marcadores})",
            _valores_colunas(dados),
        )
        bolsa_id = cursor.lastrowid
    fazer_backup()
    return bolsa_id

def atualizar_bolsa(bolsa_id, dados):
    with _conexao() as conexao:
        atribuicoes = ", ".join(f"{coluna} = ?" for coluna in _COLUNAS_BOLSA)
        conexao.execute(
            f"UPDATE bolsas SET {atribuicoes}, atualizado_em = datetime('now') WHERE id = ?",
            _valores_colunas(dados) + [bolsa_id],
        )
    fazer_backup()

def excluir_bolsa(bolsa_id):
    with _conexao() as conexao:
        conexao.execute("DELETE FROM bolsas WHERE id = ?", (bolsa_id,))
    fazer_backup()
