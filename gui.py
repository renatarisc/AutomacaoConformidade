import sys
import os
import json
import threading
import traceback

import webview # pip install pywebview - abre uma janela nativa renderizando HTML/CSS/JS (WebView2 no Windows)

import escolher_planilha
import relacionar_valor_op
import relacionar_op_ob
import baixar_ob
import anexar_ob
import baixar_anexar_ne
import preencher_planilha_ro
import preencher_planilha_ns
import cadastrar_contrato
import conformidade

# comando que abre o Chrome em modo de depuração remota (porta 9222), exigido pelos
# scripts que se conectam via options.debugger_address - ver comentário no topo de cada main()
COMANDO_CHROME_DEBUG = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeSelenium"'

# ícones de cada card em SVG inline (estilo flat/outline, no espírito do Flaticon) em vez de
# emoji - não usamos ícones hotlinkados do Flaticon de fato porque isso exigiria internet em
# tempo de execução e atribuição visível na interface; inline evita as duas coisas e funciona
# offline também. Todos com o mesmo viewBox/traço pra ficarem consistentes lado a lado.
_SVG_ATRIBUTOS = "viewBox='0 0 24 24' width='20' height='20' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'"
ICONE_DOCUMENTO = f"<svg {_SVG_ATRIBUTOS}><path d='M7 3h7l4 4v14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z'/><path d='M14 3v4h4'/><path d='M9 13h6M9 17h6M9 9h2'/></svg>"
ICONE_DOCUMENTO_BAIXAR = f"<svg {_SVG_ATRIBUTOS}><path d='M7 3h7l4 4v14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z'/><path d='M14 3v4h4'/><path d='M12 11.5v6'/><path d='m9.5 15 2.5 2.5L14.5 15'/></svg>"
ICONE_LUPA = f"<svg {_SVG_ATRIBUTOS}><circle cx='11' cy='11' r='6'/><path d='m20 20-3.8-3.8'/></svg>"
ICONE_ELO = f"<svg {_SVG_ATRIBUTOS}><path d='m10.5 13.5 3-3'/><rect x='2.5' y='12.5' width='8' height='5' rx='2.5' transform='rotate(-45 6.5 15)'/><rect x='13.5' y='6.5' width='8' height='5' rx='2.5' transform='rotate(-45 17.5 9)'/></svg>"
ICONE_BAIXAR = f"<svg {_SVG_ATRIBUTOS}><path d='M12 4v10'/><path d='m8 10.5 4 4 4-4'/><path d='M4 18h16'/></svg>"
ICONE_CLIPE = f"<svg {_SVG_ATRIBUTOS}><path d='M8 12.5V7a4 4 0 1 1 8 0v9a2.5 2.5 0 0 1-5 0V8.5'/></svg>"
ICONE_CONFERENCIA = f"<svg {_SVG_ATRIBUTOS}><circle cx='12' cy='12' r='9'/><path d='m8.5 12.5 2.5 2.5 4.5-5'/></svg>"

# só os scripts que rodam sozinhos (fazem algo ao serem executados diretamente) entram no menu;
# os módulos auxiliares (só definem funções, chamados por esses scripts) ficam de fora.
# cada um chama main() direto (em vez de abrir um subprocesso) pra funcionar também dentro
# do executável gerado pelo PyInstaller, onde não existe mais um python.exe separado pra chamar.
# "coluna": em qual coluna da interface o card aparece (1 ou 2)
OPCOES = [
    {
        "arquivo": "preencher_planilha_ro.py",
        "modulo": preencher_planilha_ro,
        "coluna": 2,
        "icone": ICONE_DOCUMENTO,
        "titulo": "Preencher Planilha de Controle (RO)",
        "descricao": "Lê as abas do Chrome com os Processos abertos em PDF e preenche a Planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug, com as abas dos processos em PDF.",
    },
    {
        "arquivo": "preencher_planilha_ns.py",
        "modulo": preencher_planilha_ns,
        "coluna": 1,
        "icone": ICONE_DOCUMENTO,
        "titulo": "Preencher Planilha de Controle (NS)",
        "descricao": "Lê as abas do Chrome com os Processos abertos em PDF e preenche a Planilha.",
        "requisito": "Usa os PDFs no Chrome em modo debug se disponível; senão, PDFs baixados.",
    },
    {
        "arquivo": "conformidade.py",
        "modulo": conformidade,
        "coluna": 1,
        "icone": ICONE_CONFERENCIA,
        "titulo": "Fazer Conformidade (NS)",
        "descricao": "Cara-Crachá: Documentos preenchidos x Fontes seguras (BD, Termo Gestor e NF).",
        "requisito": "Usa os PDFs no Chrome em modo debug se disponível; senão, PDFs baixados.",
    },
    {
        "arquivo": "baixar_anexar_ne.py",
        "modulo": baixar_anexar_ne,
        "coluna": 2,
        "icone": ICONE_DOCUMENTO_BAIXAR,
        "titulo": "Baixar e anexar NE e tramitar o processo",
        "descricao": "Baixa do Siafi a NE pintada de Cinza (assinada), anexa no processo e tramita.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "relacionar_valor_op.py",
        "modulo": relacionar_valor_op,
        "coluna": 1,
        "icone": ICONE_LUPA,
        "titulo": "Relacionar Valor → OP",
        "descricao": "Busca a OP no Siafi pelo valor (ISS/PG), grava na planilha e pinta de Amarelo.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "relacionar_op_ob.py",
        "modulo": relacionar_op_ob,
        "coluna": 1,
        "icone": ICONE_ELO,
        "titulo": "Relacionar OP → OB",
        "descricao": "Busca a OB no Siafi a partir da OP pintada de Cinza (assinada), substitui na planilha e pinta de Amarelo.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "baixar_ob.py",
        "modulo": baixar_ob,
        "coluna": 1,
        "icone": ICONE_BAIXAR,
        "titulo": "Baixar OB",
        "descricao": "Baixa do Siafi Operacional (Cara Preta) o PDF da OB pintada de Amarelo.",
        "requisito": "Requer o Sistema já aberto e com o foco antes de começar a rodar a automação (pyautogui).",
    },
    {
        "arquivo": "anexar_ob.py",
        "modulo": anexar_ob,
        "coluna": 1,
        "icone": ICONE_CLIPE,
        "titulo": "Anexar OB e tramitar o processo",
        "descricao": "Anexa no processo o PDF da OB pintada de Amarelo e tramita.",
        "requisito": "Abre e loga no Suap sozinho - não precisa preparar nada antes.",
    },
]

OPCOES_POR_ARQUIVO = {opcao["arquivo"]: opcao for opcao in OPCOES}

class EscritorJS:
    # arquivo "falso" que faz print() dentro dos scripts virar uma chamada JS, aparecendo ao vivo no console da página
    def __init__(self, window):
        self.window = window

    def write(self, texto):
        if texto:
            self.window.evaluate_js(f"appendOutput({json.dumps(texto)})")

    def flush(self):
        pass

class Api:
    # ponte Python <-> JS (window.pywebview.api.<metodo> no lado JS) - cada método aqui vira uma função
    # assíncrona chamável a partir da página. O atributo "window" é preenchido depois de criar a janela
    # (não dá pra referenciar a janela dentro dela mesma antes dela existir)
    def __init__(self):
        self.window = None
        self.executando = False

    def obter_opcoes(self):
        # não dá pra mandar "modulo" (objeto Python) pro JS - só os dados exibidos na tela
        return [
            {chave: valor for chave, valor in opcao.items() if chave != "modulo"}
            for opcao in OPCOES
        ]

    def obter_comando_chrome(self):
        return COMANDO_CHROME_DEBUG

    def abrir_cadastro_contrato(self):
        # abre a tela de contratos numa janela própria (lista + formulário + banco local),
        # separada da engine de execução de scripts usada pelos cards
        cadastrar_contrato.abrir_janela()

    def listar_planilhas(self):
        # a Planilha de Controle muda todo mês, sem data certa - em vez de editar o nome fixo em
        # cada script, lista as planilhas da pasta do Drive (escolher_planilha.PASTA_DRIVE_ID) e
        # deixa o usuário escolher aqui qual vale para todos os scripts executados a seguir
        try:
            arquivos = escolher_planilha.listar()
            return {"nomes": [arquivo["name"] for arquivo in arquivos], "erro": None}
        except Exception as e:
            return {"nomes": [], "erro": str(e)}

    def executar(self, arquivo, nome_planilha):
        if self.executando:
            return # já tem algo rodando

        if not nome_planilha:
            self.window.evaluate_js(f"definirEstado('aviso', {json.dumps('Escolha a Planilha de Controle antes de executar.')})")
            return

        opcao = OPCOES_POR_ARQUIVO.get(arquivo)
        if opcao is None:
            return

        self.executando = True
        self.window.evaluate_js(f"definirEstado('executando', {json.dumps(f'Executando {arquivo} (planilha: {nome_planilha})...')})")
        self.window.evaluate_js("definirExecutando(true)")

        thread = threading.Thread(target=self._rodar, args=(opcao, nome_planilha), daemon=True)
        thread.start()

    def _rodar(self, opcao, nome_planilha):
        # troca a saída padrão só durante a chamada, pra capturar os print()s do script
        # e mostrar ao vivo no console da página
        saida_original, erro_original = sys.stdout, sys.stderr
        escritor = EscritorJS(self.window)
        sys.stdout = escritor
        sys.stderr = escritor
        codigo_retorno = 0

        try:
            opcao["modulo"].main(nome_planilha)
        except SystemExit as e:
            # algum script chama sys.exit() no próprio fluxo de erro (ex: baixar_anexar_ne.py)
            codigo_retorno = e.code if isinstance(e.code, int) else 1
        except Exception:
            traceback.print_exc() # cai no escritor (stderr redirecionado), aparece no console da página
            codigo_retorno = 1
        finally:
            sys.stdout, sys.stderr = saida_original, erro_original

        self.executando = False
        if codigo_retorno == 0:
            self.window.evaluate_js(f"definirEstado('sucesso', {json.dumps('Concluído.')})")
        else:
            self.window.evaluate_js(f"definirEstado('erro', {json.dumps(f'Terminou com erro (código {codigo_retorno}).')})")
        self.window.evaluate_js("definirExecutando(false)")

HTML_INTERFACE = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Automações CCRGCI</title>
<style>
  :root {
    --mist: #f2f5f3;
    --mist-deep: #e7ece9;
    --cloud: #ffffff;
    --hairline: #dde4e0;
    --ink: #16201b;
    --ink-soft: #56625b;
    --ink-faint: #8a958e;
    --pine: #178c4e;
    --pine-deep: #0f6b3b;
    --pine-tint: #e2f5ea;
    --pine-tint-strong: #c3ecd6;
    --console-bg: #1a1d1b;
    --console-fg: #e7ece9;
    --status-info: #2f7fd6;
    --status-success: #2f9e63;
    --status-warning: #b3800b;
    --status-error: #d1453d;
    --shadow-1: 0 1px 2px rgba(20,32,27,0.07), 0 1px 1px rgba(20,32,27,0.05);
    --shadow-2: 0 6px 16px rgba(20,32,27,0.12), 0 2px 4px rgba(20,32,27,0.08);
  }

  * { box-sizing: border-box; }

  html, body {
    margin: 0;
    height: 100%;
    background: var(--mist);
    color: var(--ink);
    font-family: "Segoe UI Variable Display", "Segoe UI Variable Text", "Segoe UI", system-ui, sans-serif;
    font-size: 14px;
    -webkit-font-smoothing: antialiased;
  }

  body {
    background:
      radial-gradient(1100px 620px at 8% -8%, var(--pine-tint) 0%, transparent 55%),
      var(--mist);
  }

  .app {
    display: grid;
    grid-template-columns: 1fr 1fr;
    height: 100vh;
  }

  .coluna {
    display: flex;
    flex-direction: column;
    gap: 14px;
    padding: 22px 26px;
    min-width: 0;
    overflow-y: auto;
  }

  .coluna--2 { border-left: 1px solid var(--hairline); }

  .cabecalho { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
  .cabecalho__marca { display: flex; align-items: center; gap: 10px; min-width: 0; }

  .marca-icone {
    width: 34px; height: 34px; border-radius: 8px;
    background: var(--pine);
    color: #fff;
    display: flex; align-items: center; justify-content: center;
    font-size: 17px; box-shadow: var(--shadow-1); flex: 0 0 auto;
  }

  .marca-icone svg { width: 19px; height: 19px; }

  h1 {
    margin: 0;
    font-size: 19px;
    font-weight: 600;
    letter-spacing: -0.01em;
  }

  .subtitulo { margin: 2px 0 0; color: var(--ink-soft); font-size: 12.5px; }

  .painel {
    border: 1px solid var(--hairline);
    border-radius: 10px;
    background: var(--cloud);
    box-shadow: var(--shadow-1);
    padding: 12px 14px;
    flex: 0 0 auto;
  }

  .painel--acento { border-color: var(--pine-tint-strong); background: var(--pine-tint); }

  .rotulo {
    font-weight: 600; font-size: 12px; margin: 0 0 8px;
    display: flex; align-items: center; gap: 6px;
    color: var(--pine);
  }

  .rotulo .ponto { width: 6px; height: 6px; border-radius: 50%; background: var(--pine); flex: 0 0 auto; }

  .linha-controle { display: flex; gap: 8px; }

  select, input.campo-comando {
    flex: 1;
    min-width: 0;
    font-family: inherit;
    font-size: 13px;
    padding: 7px 10px;
    border: 1px solid var(--hairline);
    border-radius: 6px;
    background: var(--cloud);
    color: var(--ink);
    transition: border-color 120ms ease, box-shadow 120ms ease;
  }

  select:focus-visible, input.campo-comando:focus-visible, .btn:focus-visible {
    outline: none;
    border-color: var(--pine);
    box-shadow: 0 0 0 3px var(--pine-tint-strong);
  }

  input.campo-comando { font-family: "Cascadia Code", "Consolas", monospace; font-size: 11.5px; color: var(--ink-soft); }

  .btn {
    font-family: inherit;
    font-size: 12.5px;
    font-weight: 600;
    border: 1px solid transparent;
    border-radius: 6px;
    cursor: pointer;
    white-space: nowrap;
    transition: background 120ms ease, border-color 120ms ease, transform 80ms ease, box-shadow 120ms ease;
  }

  .btn:active { transform: translateY(1px); }
  .btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }

  .btn--icone { padding: 7px 11px; background: var(--cloud); border-color: var(--hairline); color: var(--ink-soft); }
  .btn--icone:hover:not(:disabled) { border-color: var(--pine); color: var(--pine); }

  .btn--outline { padding: 7px 16px; background: var(--cloud); border-color: var(--hairline); color: var(--ink); }
  .btn--outline:hover:not(:disabled) { border-color: var(--pine); color: var(--pine-deep); }

  .btn--acento { padding: 7px 16px; background: var(--pine-tint); border-color: var(--pine-tint-strong); color: var(--pine-deep); }
  .btn--acento:hover:not(:disabled) { background: var(--pine-tint-strong); }

  .ajuda { margin: 6px 0 0; font-size: 11px; color: var(--ink-faint); }

  .card {
    display: flex;
    align-items: center;
    gap: 13px;
    border: 1px solid var(--hairline);
    border-radius: 10px;
    background: var(--cloud);
    box-shadow: var(--shadow-1);
    padding: 11px 14px;
    flex: 0 0 auto;
    transition: box-shadow 150ms ease, border-color 150ms ease, transform 150ms ease;
  }

  .card:hover { box-shadow: var(--shadow-2); border-color: var(--pine-tint-strong); transform: translateY(-1px); }

  .card__icone {
    flex: 0 0 auto; width: 36px; height: 36px; border-radius: 8px;
    background: var(--pine-tint);
    display: flex; align-items: center; justify-content: center;
    color: var(--pine-deep);
  }

  .card__icone svg { width: 20px; height: 20px; }

  .card__texto { flex: 1 1 auto; min-width: 0; }
  .card__titulo { margin: 0; font-size: 13px; font-weight: 600; }
  .card__descricao { margin: 2px 0 0; font-size: 11.5px; color: var(--ink-soft); line-height: 1.45; }
  .card__requisito { margin: 4px 0 0; font-size: 10.5px; color: var(--ink-faint); display: flex; align-items: center; gap: 4px; }

  .card__acao {
    flex: 0 0 auto;
    padding: 8px 16px;
    background: var(--pine);
    color: #fff;
    font-size: 12.5px;
    box-shadow: var(--shadow-1);
  }

  .card__acao:hover:not(:disabled) { background: var(--pine-deep); }

  .execucao { display: flex; flex-direction: column; flex: 1 1 auto; min-height: 0; margin-top: 2px; }

  .execucao-cabecalho { display: flex; align-items: baseline; justify-content: space-between; gap: 10px; }
  .execucao h2 { margin: 6px 0 0; font-size: 15px; font-weight: 600; color: var(--pine-deep); }
  .execucao .subtitulo-exec { margin: 2px 0 8px; font-size: 12px; color: var(--ink-soft); }

  .status-chip {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 3px 10px 3px 8px; border-radius: 999px;
    font-size: 11px; font-weight: 600;
    background: var(--mist-deep); color: var(--ink-soft);
    border: 1px solid var(--hairline);
    white-space: nowrap;
  }
  .status-chip .ponto { width: 6px; height: 6px; border-radius: 50%; background: var(--ink-faint); }
  .status-chip[data-estado="executando"] { background: rgba(47,127,214,0.12); color: var(--status-info); border-color: rgba(47,127,214,0.3); }
  .status-chip[data-estado="executando"] .ponto { background: var(--status-info); animation: pulso 1.1s ease-in-out infinite; }
  .status-chip[data-estado="sucesso"] { background: rgba(47,158,99,0.12); color: var(--status-success); border-color: rgba(47,158,99,0.3); }
  .status-chip[data-estado="sucesso"] .ponto { background: var(--status-success); }
  .status-chip[data-estado="aviso"] { background: rgba(179,128,11,0.12); color: var(--status-warning); border-color: rgba(179,128,11,0.3); }
  .status-chip[data-estado="aviso"] .ponto { background: var(--status-warning); }
  .status-chip[data-estado="erro"] { background: rgba(209,69,61,0.12); color: var(--status-error); border-color: rgba(209,69,61,0.3); }
  .status-chip[data-estado="erro"] .ponto { background: var(--status-error); }

  @keyframes pulso { 0%, 100% { opacity: 1; } 50% { opacity: 0.35; } }

  .console {
    flex: 1 1 auto;
    min-height: 120px;
    background: var(--console-bg);
    color: var(--console-fg);
    border-radius: 10px;
    padding: 12px 14px;
    font-family: "Cascadia Code", "Consolas", monospace;
    font-size: 12px;
    line-height: 1.65;
    overflow-y: auto;
    white-space: pre-wrap;
    box-shadow: var(--shadow-1);
  }

  .console:empty::before { content: "Nenhuma automação em execução."; color: var(--ink-faint); opacity: 0.55; }
  .console.executando:empty::before { content: "Executando... aguarde a primeira mensagem aparecer aqui."; }
</style>
</head>
<body>

<div class="app">
  <div class="coluna coluna--1">
    <div class="cabecalho">
      <div class="cabecalho__marca">
        <div class="marca-icone"><svg viewBox="0 0 24 24" width="19" height="19" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg></div>
        <div>
          <h1>Automações CCRGCI</h1>
          <p class="subtitulo">Escolha a planilha e uma automação para executar.</p>
        </div>
      </div>
      <button class="btn btn--acento" id="cadastrar-contrato" title="Cadastrar/Editar contrato">Cadastro de Contrato</button>
    </div>

    <div class="painel painel--acento">
      <p class="rotulo"><span class="ponto"></span>Planilha de Controle</p>
      <div class="linha-controle">
        <select id="planilha-select"></select>
        <button class="btn btn--icone" id="planilha-atualizar" title="Atualizar lista">⟳</button>
      </div>
      <p class="ajuda" id="planilha-status"></p>
    </div>

    <div id="coluna1-cards"></div>
  </div>

  <div class="coluna coluna--2">
    <div class="painel painel--acento">
      <p class="rotulo"><span class="ponto"></span>Comando para abrir o Chrome em modo debug</p>
      <div class="linha-controle">
        <input class="campo-comando" id="chrome-comando" readonly>
        <button class="btn btn--outline" id="chrome-copiar">Copiar</button>
      </div>
      <p class="ajuda">Feche todos os Chromes abertos. Clique em Abrir ou copie e cole no CMD antes de rodar a automação.</p>
    </div>

    <div id="coluna2-cards"></div>

    <div class="execucao">
      <div class="execucao-cabecalho">
        <h2>Execução da Automação</h2>
        <span class="status-chip" id="status-chip" data-estado="ocioso"><span class="ponto"></span><span id="status-texto">Ocioso</span></span>
      </div>
      <p class="subtitulo-exec">Acompanhe aqui o andamento e as mensagens da automação em execução.</p>
      <div class="console" id="console"></div>
    </div>
  </div>
</div>

<script>
  let planilhaEscolhida = "";

  function montarCard(opcao) {
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <div class="card__icone">${opcao.icone}</div>
      <div class="card__texto">
        <p class="card__titulo"></p>
        <p class="card__descricao"></p>
        <p class="card__requisito"></p>
      </div>
      <button class="btn card__acao">Executar ▶</button>
    `;
    card.querySelector(".card__titulo").textContent = opcao.titulo;
    card.querySelector(".card__descricao").textContent = opcao.descricao;
    card.querySelector(".card__requisito").textContent = "⚠ " + opcao.requisito;

    const botao = card.querySelector(".card__acao");
    botao.dataset.arquivo = opcao.arquivo;
    botao.addEventListener("click", () => executarScript(opcao.arquivo));

    return card;
  }

  function definirEstado(estado, texto) {
    document.getElementById("status-chip").dataset.estado = estado;
    document.getElementById("status-texto").textContent = texto;
  }

  function appendOutput(texto) {
    const console_ = document.getElementById("console");
    console_.textContent += texto;
    console_.scrollTop = console_.scrollHeight;
  }

  function definirExecutando(executando) {
    document.querySelectorAll(".card__acao").forEach((botao) => { botao.disabled = executando; });
    document.getElementById("console").classList.toggle("executando", executando);
  }

  async function executarScript(arquivo) {
    planilhaEscolhida = document.getElementById("planilha-select").value;
    document.getElementById("console").textContent = "";
    await window.pywebview.api.executar(arquivo, planilhaEscolhida);
  }

  async function copiarComandoChrome() {
    const texto = document.getElementById("chrome-comando").value;
    try {
      await navigator.clipboard.writeText(texto);
    } catch (e) {
      // navegadores/engines mais antigos não têm Clipboard API - alternativa via seleção + execCommand
      const campo = document.getElementById("chrome-comando");
      campo.removeAttribute("readonly");
      campo.select();
      document.execCommand("copy");
      campo.setAttribute("readonly", "true");
    }
  }

  async function carregarPlanilhas() {
    const status = document.getElementById("planilha-status");
    const select = document.getElementById("planilha-select");
    status.textContent = "⏳ Carregando planilhas da pasta do Drive...";

    const resultado = await window.pywebview.api.listar_planilhas();
    if (resultado.erro) {
      status.textContent = "❌ Erro ao listar planilhas: " + resultado.erro;
      return;
    }

    select.innerHTML = "";
    resultado.nomes.forEach((nome) => {
      const item = document.createElement("option");
      item.value = nome;
      item.textContent = nome;
      select.appendChild(item);
    });

    if (resultado.nomes.length) {
      status.textContent = "Confira se é a planilha certa antes de executar.";
    } else {
      status.textContent = "Nenhuma planilha encontrada na pasta indicada.";
    }
  }

  async function iniciar() {
    document.getElementById("chrome-comando").value = await window.pywebview.api.obter_comando_chrome();
    document.getElementById("chrome-copiar").addEventListener("click", copiarComandoChrome);
    document.getElementById("planilha-atualizar").addEventListener("click", carregarPlanilhas);
    document.getElementById("cadastrar-contrato").addEventListener("click", () => window.pywebview.api.abrir_cadastro_contrato());

    const opcoes = await window.pywebview.api.obter_opcoes();
    const coluna1 = document.getElementById("coluna1-cards");
    const coluna2 = document.getElementById("coluna2-cards");
    opcoes.forEach((opcao) => {
      const alvo = opcao.coluna === 1 ? coluna1 : coluna2;
      alvo.appendChild(montarCard(opcao));
    });

    await carregarPlanilhas();
  }

  window.addEventListener("pywebviewready", iniciar);
</script>
</body>
</html>
"""

if __name__ == "__main__":
    api = Api()
    window = webview.create_window(
        "Automações da Coordenação de Conformidade de Registro de Gestão do Campus Itaperuna - CCRGCI",
        html=HTML_INTERFACE,
        js_api=api,
        width=1400,
        height=860,
        min_size=(1150, 520),
        maximized=True,
        text_select=True, # sem isso não dá pra selecionar/copiar texto na página (ex: o traceback no console) - False é o padrão do pywebview
    )
    # NÃO atribui api.window = window aqui direto - fazer isso antes da página carregar é a causa
    # confirmada do erro "[pywebview] Error while processing window.native..." (travamento por acesso
    # cross-thread ao COM do WebView2 antes de pronto - https://github.com/r0x0r/pywebview/issues/1815).
    # Só atribui quando o evento "loaded" da janela dispara, depois que o WebView2 já está pronto.
    def _ao_carregar():
        api.window = window

    window.events.loaded += _ao_carregar

    icone = os.path.join(os.path.dirname(__file__), "icon.ico") # webview.start() (não create_window) que recebe o ícone
    webview.start(icon=icone if os.path.exists(icone) else None)
