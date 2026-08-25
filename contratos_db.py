import contextlib
import csv
import os
import re
import shutil
import sqlite3
from datetime import datetime

# banco nativo do Python (sqlite3, sem dependência externa) - volume de contratos é pequeno,
# não justifica um servidor de banco de dados separado
CAMINHO_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contratos.db")

# esquema de backup real: a cada abertura da tela e a cada gravação (criar/editar/excluir),
# grava um snapshot completo (cópia binária do .db + CSV de cada tabela) numa pasta datada
# dentro de PASTA_BACKUPS. Existe pra que um erro de programação, migração ou operação
# manual nunca mais custe dado real - ver [[project-suap-ob-pipeline]]/histórico do projeto
PASTA_BACKUPS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "contratos_backups")
MAX_BACKUPS = 300  # válvula de segurança contra crescimento sem fim - cada snapshot é minúsculo

_ESQUEMA_SQL = """
CREATE TABLE IF NOT EXISTS contratos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo_contrato TEXT NOT NULL,
    situacao TEXT NOT NULL DEFAULT 'vigente',
    nome_contratada TEXT NOT NULL,
    nome_planilha_controle TEXT,
    cnpj TEXT,
    objeto_resumido TEXT,
    objeto_detalhado TEXT,
    numero_pregao TEXT,
    numero_contrato TEXT,
    vigencia_inicio TEXT,
    vigencia_fim TEXT,
    processo_contratacao TEXT,
    processo_empenho_anual TEXT,
    banco TEXT,
    agencia TEXT,
    conta TEXT,
    iss_incide INTEGER NOT NULL DEFAULT 0,
    iss_aliquota REAL,
    previdenciaria_incide INTEGER NOT NULL DEFAULT 0,
    previdenciaria_aliquota REAL,
    federais_incide INTEGER NOT NULL DEFAULT 0,
    federais_codigo_darf TEXT,
    federais_aliquota_total REAL,
    observacao TEXT,
    criado_em TEXT NOT NULL DEFAULT (datetime('now')),
    atualizado_em TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS contrato_valores_mensais (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contrato_id INTEGER NOT NULL REFERENCES contratos(id) ON DELETE CASCADE,
    descricao_servico TEXT,
    valor REAL NOT NULL
);

-- só existe pra contrato de almoxarifado: um mesmo contrato tem vários processos de
-- empenho (ao contrário do de serviço, que tem um processo_empenho_anual único lá em
-- "contratos"), e cada processo agrupa vários empenhos
CREATE TABLE IF NOT EXISTS contrato_processos_empenho (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contrato_id INTEGER NOT NULL REFERENCES contratos(id) ON DELETE CASCADE,
    numero_processo TEXT NOT NULL
);

-- processo_empenho_id fica NULL para empenho de contrato de serviço (vinculado direto
-- ao contrato, com controle de saldo em empenho_movimentacoes) e preenchido para
-- empenho de almoxarifado (vinculado ao processo de empenho que o agrupa, sem saldo)
CREATE TABLE IF NOT EXISTS contrato_empenhos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contrato_id INTEGER NOT NULL REFERENCES contratos(id) ON DELETE CASCADE,
    processo_empenho_id INTEGER REFERENCES contrato_processos_empenho(id) ON DELETE CASCADE,
    numero_empenho TEXT NOT NULL,
    natureza_despesa TEXT
);

-- cada entrada/saída no empenho grava o saldo já resultante daquela movimentação
-- (não só o valor movimentado) para permitir consultar o saldo vigente em cada data,
-- sem precisar re-somar o histórico inteiro toda vez - só usada por empenho de serviço
-- (almoxarifado não tem controle de saldo)
CREATE TABLE IF NOT EXISTS empenho_movimentacoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    empenho_id INTEGER NOT NULL REFERENCES contrato_empenhos(id) ON DELETE CASCADE,
    data TEXT NOT NULL,
    tipo TEXT NOT NULL,
    valor_movimentado REAL NOT NULL,
    saldo_apos REAL NOT NULL
);
"""

# ordem importa na restauração (tabelas-pai antes das que têm FK pra elas)
_TABELAS_EXPORTAVEIS = (
    "contratos",
    "contrato_valores_mensais",
    "contrato_processos_empenho",
    "contrato_empenhos",
    "empenho_movimentacoes",
)

# colunas que precisam voltar a ser int/float na restauração a partir do CSV (que só tem
# texto) - todo o resto (inclusive as colunas de texto vazias) volta como string ou None
_COLUNAS_NUMERICAS = {
    "contratos": {
        "id": int, "iss_incide": int, "iss_aliquota": float,
        "previdenciaria_incide": int, "previdenciaria_aliquota": float,
        "federais_incide": int, "federais_aliquota_total": float,
    },
    "contrato_valores_mensais": {"id": int, "contrato_id": int, "valor": float},
    "contrato_processos_empenho": {"id": int, "contrato_id": int},
    "contrato_empenhos": {"id": int, "contrato_id": int, "processo_empenho_id": int},
    "empenho_movimentacoes": {"id": int, "empenho_id": int, "valor_movimentado": float, "saldo_apos": float},
}

@contextlib.contextmanager
def _conexao():
    # "with sqlite3.Connection" só cuida do commit/rollback da transação, não fecha a
    # conexão - sem esse wrapper cada chamada vazaria uma conexão aberta (o gui.py fica
    # rodando a sessão inteira, então isso se acumularia)
    conexao = sqlite3.connect(CAMINHO_DB)
    conexao.row_factory = sqlite3.Row
    conexao.execute("PRAGMA foreign_keys = ON")
    try:
        with conexao:
            yield conexao
    finally:
        conexao.close()

def inicializar_db():
    fazer_backup()  # snapshot do estado atual ANTES de qualquer criação/migração de tabela
    with _conexao() as conexao:
        conexao.executescript(_ESQUEMA_SQL)
        _migrar_esquema(conexao)

def _migrar_esquema(conexao):
    # bancos criados antes desta versão têm "processo_empenho" (não "_anual") e
    # contrato_empenhos sem natureza_despesa/processo_empenho_id - ajusta em cima do banco
    # já existente (preservando os contratos já cadastrados) em vez de exigir apagar tudo
    colunas_contratos = {linha["name"] for linha in conexao.execute("PRAGMA table_info(contratos)")}
    if "processo_empenho" in colunas_contratos and "processo_empenho_anual" not in colunas_contratos:
        conexao.execute("ALTER TABLE contratos RENAME COLUMN processo_empenho TO processo_empenho_anual")
    if "situacao" not in colunas_contratos:
        conexao.execute("ALTER TABLE contratos ADD COLUMN situacao TEXT NOT NULL DEFAULT 'vigente'")
    if "objeto" in colunas_contratos and "objeto_resumido" not in colunas_contratos:
        conexao.execute("ALTER TABLE contratos RENAME COLUMN objeto TO objeto_resumido")
    colunas_contratos = {linha["name"] for linha in conexao.execute("PRAGMA table_info(contratos)")}
    if "objeto_detalhado" not in colunas_contratos:
        conexao.execute("ALTER TABLE contratos ADD COLUMN objeto_detalhado TEXT")

    colunas_empenhos = {linha["name"] for linha in conexao.execute("PRAGMA table_info(contrato_empenhos)")}
    if "natureza_despesa" not in colunas_empenhos:
        conexao.execute("ALTER TABLE contrato_empenhos ADD COLUMN natureza_despesa TEXT")
    if "processo_empenho_id" not in colunas_empenhos:
        conexao.execute(
            "ALTER TABLE contrato_empenhos ADD COLUMN processo_empenho_id INTEGER "
            "REFERENCES contrato_processos_empenho(id) ON DELETE CASCADE"
        )

# ---------- backup ----------

def fazer_backup():
    # grava um snapshot completo (contratos.db copiado bit a bit + um CSV por tabela) numa
    # pasta datada - chamado na abertura da tela e depois de toda gravação (ver criar_/
    # atualizar_/excluir_contrato). Se o .db ainda não existe (primeiríssima execução) não
    # há o que copiar; retorna None nesse caso.
    if not os.path.exists(CAMINHO_DB):
        return None

    carimbo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    pasta_destino = os.path.join(PASTA_BACKUPS, carimbo)
    os.makedirs(pasta_destino, exist_ok=True)

    shutil.copy2(CAMINHO_DB, os.path.join(pasta_destino, "contratos.db"))

    conexao = sqlite3.connect(CAMINHO_DB)
    conexao.row_factory = sqlite3.Row
    try:
        _exportar_csvs(conexao, pasta_destino)
    finally:
        conexao.close()

    _rotacionar_backups()
    return pasta_destino

def _exportar_csvs(conexao, pasta_destino):
    for tabela in _TABELAS_EXPORTAVEIS:
        linhas = conexao.execute(f"SELECT * FROM {tabela}").fetchall()
        colunas = (
            list(linhas[0].keys()) if linhas
            else [descricao[0] for descricao in conexao.execute(f"SELECT * FROM {tabela} LIMIT 0").description]
        )
        # utf-8-sig (com BOM) pra abrir certo no Excel, com acento, sem configuração manual
        with open(os.path.join(pasta_destino, f"{tabela}.csv"), "w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo)
            escritor.writerow(colunas)
            for linha in linhas:
                escritor.writerow([linha[coluna] for coluna in colunas])

def _rotacionar_backups():
    if not os.path.isdir(PASTA_BACKUPS):
        return
    pastas = sorted(
        nome for nome in os.listdir(PASTA_BACKUPS)
        if os.path.isdir(os.path.join(PASTA_BACKUPS, nome))
    )
    for nome in pastas[:-MAX_BACKUPS] if len(pastas) > MAX_BACKUPS else []:
        shutil.rmtree(os.path.join(PASTA_BACKUPS, nome), ignore_errors=True)

def listar_backups():
    # mais recente primeiro - útil pra uma futura tela de "restaurar backup"
    if not os.path.isdir(PASTA_BACKUPS):
        return []
    return sorted(
        (nome for nome in os.listdir(PASTA_BACKUPS) if os.path.isdir(os.path.join(PASTA_BACKUPS, nome))),
        reverse=True,
    )

# ---------- restauração ----------
#
# duas vias independentes de propósito - nenhuma delas escreve em cima de contratos.db ou
# de um arquivo já existente: o chamador escolhe o destino e decide, depois de conferir o
# conteúdo restaurado, se quer substituir o contratos.db real manualmente.
#
# 1) restaurar_de_copia_binaria: mais simples e fiel - é literalmente o arquivo .db daquele
#    momento. Via principal de recuperação.
# 2) restaurar_de_csv: reconstrói o banco linha a linha a partir dos CSVs. Existe como via
#    independente do arquivo .db (por exemplo, se ele também tiver sido perdido/corrompido
#    mas os CSVs sobreviveram, ou pra restaurar num ambiente sem o arquivo binário à mão).

def restaurar_de_copia_binaria(pasta_backup, caminho_destino):
    origem = os.path.join(pasta_backup, "contratos.db")
    if not os.path.exists(origem):
        raise FileNotFoundError(f"{origem} não existe nesse backup")
    if os.path.exists(caminho_destino):
        raise FileExistsError(f"{caminho_destino} já existe - escolha outro destino pra não sobrescrever nada")
    shutil.copy2(origem, caminho_destino)
    return caminho_destino

def _converter_linha(tabela, linha_csv):
    convertida = dict(linha_csv)
    for coluna, tipo in _COLUNAS_NUMERICAS.get(tabela, {}).items():
        valor = convertida.get(coluna)
        convertida[coluna] = tipo(valor) if valor not in (None, "") else None
    for coluna, valor in convertida.items():
        if valor == "":
            convertida[coluna] = None
    return convertida

def restaurar_de_csv(pasta_backup, caminho_destino):
    if os.path.exists(caminho_destino):
        raise FileExistsError(f"{caminho_destino} já existe - escolha outro destino pra não sobrescrever nada")

    conexao = sqlite3.connect(caminho_destino)
    try:
        conexao.executescript(_ESQUEMA_SQL)
        for tabela in _TABELAS_EXPORTAVEIS:
            caminho_csv = os.path.join(pasta_backup, f"{tabela}.csv")
            if not os.path.exists(caminho_csv):
                continue
            with open(caminho_csv, newline="", encoding="utf-8-sig") as arquivo:
                linhas = list(csv.DictReader(arquivo))
            if not linhas:
                continue
            colunas = list(linhas[0].keys())
            marcadores = ", ".join("?" for _ in colunas)
            for linha_csv in linhas:
                linha = _converter_linha(tabela, linha_csv)
                conexao.execute(
                    f"INSERT INTO {tabela} ({', '.join(colunas)}) VALUES ({marcadores})",
                    [linha[coluna] for coluna in colunas],
                )
        conexao.commit()
    finally:
        conexao.close()
    return caminho_destino

def listar_contratos():
    with _conexao() as conexao:
        linhas = conexao.execute(
            "SELECT id, tipo_contrato, situacao, nome_contratada, nome_planilha_controle, numero_contrato, "
            "vigencia_inicio, vigencia_fim FROM contratos ORDER BY nome_contratada COLLATE NOCASE"
        ).fetchall()
        return [dict(linha) for linha in linhas]

def obter_abreviacao_empresa(cnpj=None, nome_contratada=None):
    # usado pelos scripts preencher_planilha_ro.py/preencher_planilha_ns.py: eles extraem o nome
    # completo da empresa (e o CNPJ, quando disponível) direto do PDF/SIAFI, e usam isso aqui pra
    # descobrir a abreviação já cadastrada em "Planilha de controle" (ex: "A M GAMBA ALIMENTOS" ->
    # "GAMBA") - CNPJ é a chave preferida (não depende de bater maiúscula/pontuação/etc.), nome é
    # só reforço se o CNPJ não vier ou não bater com nada. None se não achar o contrato - quem
    # chama decide o que fazer nesse caso (ex: manter o nome completo em vez de abreviar)
    digitos_cnpj = re.sub(r"\D", "", cnpj) if cnpj else ""
    with _conexao() as conexao:
        if digitos_cnpj:
            linha = conexao.execute(
                "SELECT nome_planilha_controle FROM contratos WHERE cnpj = ?", (digitos_cnpj,)
            ).fetchone()
            if linha and linha["nome_planilha_controle"]:
                return linha["nome_planilha_controle"]
        if nome_contratada:
            linha = conexao.execute(
                "SELECT nome_planilha_controle FROM contratos WHERE nome_contratada = ? COLLATE NOCASE",
                (nome_contratada.strip(),),
            ).fetchone()
            if linha and linha["nome_planilha_controle"]:
                return linha["nome_planilha_controle"]
    return None

def obter_contrato_por_cnpj(cnpj):
    # usado pelo script de conformidade: acha o contrato inteiro (com valores mensais, empenhos
    # etc.) a partir do CNPJ extraído do PDF - aceita com ou sem máscara
    digitos = re.sub(r"\D", "", cnpj) if cnpj else ""
    if not digitos:
        return None
    with _conexao() as conexao:
        linha = conexao.execute("SELECT id FROM contratos WHERE cnpj = ?", (digitos,)).fetchone()
    return obter_contrato(linha["id"]) if linha else None

def obter_contrato(contrato_id):
    with _conexao() as conexao:
        linha = conexao.execute("SELECT * FROM contratos WHERE id = ?", (contrato_id,)).fetchone()
        if linha is None:
            return None
        contrato = dict(linha)
        contrato["valores_mensais"] = [
            dict(v) for v in conexao.execute(
                "SELECT id, descricao_servico, valor FROM contrato_valores_mensais WHERE contrato_id = ? ORDER BY id",
                (contrato_id,),
            ).fetchall()
        ]

        # empenhos "diretos" (contrato de serviço) - com controle de saldo
        empenhos = [
            dict(e) for e in conexao.execute(
                "SELECT id, numero_empenho, natureza_despesa FROM contrato_empenhos "
                "WHERE contrato_id = ? AND processo_empenho_id IS NULL ORDER BY id",
                (contrato_id,),
            ).fetchall()
        ]
        for empenho in empenhos:
            empenho["movimentacoes"] = [
                dict(m) for m in conexao.execute(
                    "SELECT id, data, tipo, valor_movimentado, saldo_apos FROM empenho_movimentacoes "
                    "WHERE empenho_id = ? ORDER BY data, id",
                    (empenho["id"],),
                ).fetchall()
            ]
        contrato["empenhos"] = empenhos

        # processos de empenho e seus empenhos (contrato de almoxarifado) - sem saldo
        processos = [
            dict(p) for p in conexao.execute(
                "SELECT id, numero_processo FROM contrato_processos_empenho WHERE contrato_id = ? ORDER BY id",
                (contrato_id,),
            ).fetchall()
        ]
        for processo in processos:
            processo["empenhos"] = [
                dict(e) for e in conexao.execute(
                    "SELECT id, numero_empenho, natureza_despesa FROM contrato_empenhos "
                    "WHERE processo_empenho_id = ? ORDER BY id",
                    (processo["id"],),
                ).fetchall()
            ]
        contrato["processos_empenho"] = processos

        return contrato

_COLUNAS_CONTRATO = (
    "tipo_contrato", "situacao", "nome_contratada", "nome_planilha_controle", "cnpj",
    "objeto_resumido", "objeto_detalhado",
    "numero_pregao", "numero_contrato", "vigencia_inicio", "vigencia_fim",
    "processo_contratacao", "processo_empenho_anual", "banco", "agencia", "conta",
    "iss_incide", "iss_aliquota", "previdenciaria_incide", "previdenciaria_aliquota",
    "federais_incide", "federais_codigo_darf", "federais_aliquota_total", "observacao",
)

def _valores_colunas(dados):
    valores = []
    for coluna in _COLUNAS_CONTRATO:
        valor = dados.get(coluna)
        if coluna.endswith("_incide"):
            valor = int(bool(valor))
        valores.append(valor)
    return valores

def _salvar_movimentacoes(conexao, empenho_id, movimentacoes):
    # saldo é sempre recalculado no servidor a partir da ordem cronológica (data, depois a
    # ordem em que a tela mandou os itens para desempate) - nunca confia num saldo_apos que
    # a tela eventualmente já tenha mostrado, pra não deixar o saldo salvo divergir da soma
    # real de entradas/saídas
    ordenadas = sorted(enumerate(movimentacoes), key=lambda par: (par[1]["data"], par[0]))
    saldo = 0.0
    for _, mov in ordenadas:
        valor_movimentado = float(mov["valor_movimentado"])
        saldo += valor_movimentado if mov["tipo"] == "entrada" else -valor_movimentado
        conexao.execute(
            "INSERT INTO empenho_movimentacoes (empenho_id, data, tipo, valor_movimentado, saldo_apos) "
            "VALUES (?, ?, ?, ?, ?)",
            (empenho_id, mov["data"], mov["tipo"], valor_movimentado, saldo),
        )

def _salvar_filhos(conexao, contrato_id, dados):
    # substitui os filhos por completo em vez de tentar casar diffs - volume pequeno e a
    # tela sempre manda a lista inteira (valores mensais, empenhos/processos e movimentações)
    # a cada save
    conexao.execute("DELETE FROM contrato_valores_mensais WHERE contrato_id = ?", (contrato_id,))
    for item in dados.get("valores_mensais", []):
        conexao.execute(
            "INSERT INTO contrato_valores_mensais (contrato_id, descricao_servico, valor) VALUES (?, ?, ?)",
            (contrato_id, item.get("descricao_servico"), float(item["valor"])),
        )

    # apaga os empenhos "diretos" (serviço) e os processos de empenho - apagar um processo
    # já remove em cascata os empenhos de almoxarifado dele (ON DELETE CASCADE); apagar
    # qualquer empenho já remove em cascata suas movimentações
    conexao.execute(
        "DELETE FROM contrato_empenhos WHERE contrato_id = ? AND processo_empenho_id IS NULL",
        (contrato_id,),
    )
    conexao.execute("DELETE FROM contrato_processos_empenho WHERE contrato_id = ?", (contrato_id,))

    # empenhos diretos (contrato de serviço) - com controle de saldo
    for empenho in dados.get("empenhos", []):
        cursor = conexao.execute(
            "INSERT INTO contrato_empenhos (contrato_id, numero_empenho, natureza_despesa) VALUES (?, ?, ?)",
            (contrato_id, empenho["numero_empenho"], empenho.get("natureza_despesa")),
        )
        _salvar_movimentacoes(conexao, cursor.lastrowid, empenho.get("movimentacoes", []))

    # processos de empenho e seus empenhos (contrato de almoxarifado) - sem saldo
    for processo in dados.get("processos_empenho", []):
        cursor = conexao.execute(
            "INSERT INTO contrato_processos_empenho (contrato_id, numero_processo) VALUES (?, ?)",
            (contrato_id, processo["numero_processo"]),
        )
        processo_id = cursor.lastrowid
        for empenho in processo.get("empenhos", []):
            conexao.execute(
                "INSERT INTO contrato_empenhos (contrato_id, processo_empenho_id, numero_empenho, natureza_despesa) "
                "VALUES (?, ?, ?, ?)",
                (contrato_id, processo_id, empenho["numero_empenho"], empenho.get("natureza_despesa")),
            )

def criar_contrato(dados):
    with _conexao() as conexao:
        marcadores = ", ".join("?" for _ in _COLUNAS_CONTRATO)
        cursor = conexao.execute(
            f"INSERT INTO contratos ({', '.join(_COLUNAS_CONTRATO)}) VALUES ({marcadores})",
            _valores_colunas(dados),
        )
        contrato_id = cursor.lastrowid
        _salvar_filhos(conexao, contrato_id, dados)
    fazer_backup()  # snapshot pós-gravação, fora da transação (já commitada e fechada)
    return contrato_id

def atualizar_contrato(contrato_id, dados):
    with _conexao() as conexao:
        atribuicoes = ", ".join(f"{coluna} = ?" for coluna in _COLUNAS_CONTRATO)
        conexao.execute(
            f"UPDATE contratos SET {atribuicoes}, atualizado_em = datetime('now') WHERE id = ?",
            _valores_colunas(dados) + [contrato_id],
        )
        _salvar_filhos(conexao, contrato_id, dados)
    fazer_backup()

def excluir_contrato(contrato_id):
    with _conexao() as conexao:
        conexao.execute("DELETE FROM contratos WHERE id = ?", (contrato_id,))
    fazer_backup()  # snapshot pós-exclusão - a recuperação do contrato excluído vem do
                     # snapshot ANTERIOR a esta chamada, mantido pela rotação de backups
