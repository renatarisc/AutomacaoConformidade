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

# comando que abre o Chrome em modo de depuração remota (porta 9222), exigido pelos
# scripts que se conectam via options.debugger_address - ver comentário no topo de cada main()
COMANDO_CHROME_DEBUG = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeSelenium"'

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
        "icone": "📄",
        "titulo": "Preencher Planilha de Controle (RO)",
        "descricao": "Lê as abas do Chrome com os Processos abertos em PDF e preenche a Planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug, com as abas dos processos em PDF.",
    },
    {
        "arquivo": "preencher_planilha_ns.py",
        "modulo": preencher_planilha_ns,
        "coluna": 1,
        "icone": "📄",
        "titulo": "Preencher Planilha de Controle (NS)",
        "descricao": "Lê as abas do Chrome com os Processos abertos em PDF e preenche a Planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug, com as abas dos processos em PDF.",
    },
    {
        "arquivo": "baixar_anexar_ne.py",
        "modulo": baixar_anexar_ne,
        "coluna": 2,
        "icone": "📥",
        "titulo": "Baixar e anexar NE e tramitar o processo",
        "descricao": "Baixa do Siafi a NE pintada de Cinza, anexa no processo e tramita.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "relacionar_valor_op.py",
        "modulo": relacionar_valor_op,
        "coluna": 1,
        "icone": "🔍",
        "titulo": "Relacionar Valor → OP",
        "descricao": "Busca a OP no Siafi pelo valor (ISS/PG), grava na planilha e pinta de Amarelo.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "relacionar_op_ob.py",
        "modulo": relacionar_op_ob,
        "coluna": 1,
        "icone": "🔗",
        "titulo": "Relacionar OP → OB",
        "descricao": "Busca a OB no Siafi a partir da OP pintada de Cinza e substitui na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "baixar_ob.py",
        "modulo": baixar_ob,
        "coluna": 1,
        "icone": "⬇️",
        "titulo": "Baixar OB",
        "descricao": "Baixa do Cara Preta o PDF da OB pintada de Amarelo (via pyautogui).",
        "requisito": "Requer o Sistema já aberto e com foco antes de rodar a automação.",
    },
    {
        "arquivo": "anexar_ob.py",
        "modulo": anexar_ob,
        "coluna": 1,
        "icone": "📎",
        "titulo": "Anexar OB e tramitar o processo",
        "descricao": "Anexa no processo o PDF da OB pintada de amarelo e tramita.",
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
            self.window.evaluate_js(f"definirStatus({json.dumps('⚠ Escolha a Planilha de Controle antes de executar.')})")
            return

        opcao = OPCOES_POR_ARQUIVO.get(arquivo)
        if opcao is None:
            return

        self.executando = True
        self.window.evaluate_js(f"definirStatus({json.dumps(f'⏳ Executando {arquivo} (planilha: {nome_planilha})...')})")
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
            self.window.evaluate_js(f"definirStatus({json.dumps('✅ Concluído.')})")
        else:
            self.window.evaluate_js(f"definirStatus({json.dumps(f'❌ Terminou com erro (código {codigo_retorno}).')})")
        self.window.evaluate_js("definirExecutando(false)")

HTML_INTERFACE = r"""
<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<title>Automação da Conformidade</title>
<style>
  :root {
    --bg: #eef1f0;
    --surface: #ffffff;
    --border: #d7ddda;
    --accent: #45a37e;
    --accent-soft: #7fb89e;
    --text: #23292c;
    --text-muted: #667077;
    --console-bg: #1e1e1e;
    --console-fg: #e6e6e6;
    --status-info: #2f8fd6;
  }

  * { box-sizing: border-box; }

  html, body {
    margin: 0;
    height: 100%;
    background: var(--bg);
    color: var(--text);
    font-family: "Segoe UI", "Segoe UI Variable", system-ui, sans-serif;
    font-size: 14px;
  }

  .app {
    display: grid;
    grid-template-columns: 1fr 1fr;
    height: 100vh;
  }

  .coluna {
    display: flex;
    flex-direction: column;
    gap: 12px;
    padding: 18px 22px;
    min-width: 0;
    overflow-y: auto;
  }

  .coluna--2 { border-left: 1px solid var(--border); }

  h1 {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 0;
    font-size: 22px;
    font-weight: 700;
    color: var(--accent);
    letter-spacing: -0.01em;
  }

  .subtitulo { margin: 2px 0 4px; color: var(--accent-soft); font-size: 13px; }

  .caixa-destaque {
    border: 2px solid var(--accent);
    border-radius: 6px;
    background: var(--surface);
    padding: 10px 12px;
    flex: 0 0 auto;
  }

  .caixa-destaque .rotulo { font-weight: 700; font-size: 12.5px; margin: 0 0 6px; }

  .linha-controle { display: flex; gap: 8px; }

  select, input.campo-comando {
    flex: 1;
    min-width: 0;
    font-family: inherit;
    font-size: 13px;
    padding: 6px 8px;
    border: 1px solid var(--border);
    border-radius: 4px;
    background: var(--surface);
    color: var(--text);
  }

  input.campo-comando { font-family: "Consolas", "Cascadia Mono", monospace; font-size: 12px; color: var(--text-muted); }

  .btn {
    font-family: inherit;
    font-size: 12.5px;
    font-weight: 600;
    border: none;
    border-radius: 4px;
    cursor: pointer;
    white-space: nowrap;
  }

  .btn:disabled { opacity: 0.55; cursor: not-allowed; }

  .btn--icone { padding: 6px 10px; background: transparent; border: 1px solid var(--border); color: var(--text-muted); }
  .btn--outline { padding: 6px 14px; background: transparent; border: 1px solid var(--accent); color: var(--accent); }
  .btn--outline:hover { background: rgba(69, 163, 126, 0.1); }

  .ajuda { margin: 6px 0 0; font-size: 11.5px; color: var(--text-muted); }

  .card {
    display: flex;
    align-items: center;
    gap: 12px;
    border: 1px solid var(--border);
    border-radius: 6px;
    background: var(--surface);
    padding: 10px 14px;
    flex: 0 0 auto;
  }

  .card__icone { flex: 0 0 auto; width: 34px; text-align: center; font-size: 20px; line-height: 1; }
  .card__texto { flex: 1 1 auto; min-width: 0; max-width: 46ch; }
  .card__titulo { margin: 0; font-size: 13.5px; font-weight: 700; }
  .card__descricao { margin: 2px 0 0; font-size: 12px; color: var(--text-muted); line-height: 1.4; }
  .card__requisito { margin: 3px 0 0; font-size: 10.5px; color: var(--accent-soft); }

  .card__acao {
    flex: 0 0 auto;
    margin-left: 12px;
    padding: 7px 14px;
    background: var(--accent);
    color: #fff;
    font-size: 12.5px;
  }

  .card__acao:hover:not(:disabled) { background: #397f62; }

  .execucao { display: flex; flex-direction: column; flex: 1 1 auto; min-height: 0; }
  .execucao h2 { margin: 4px 0 0; font-size: 17px; font-weight: 700; color: var(--accent); }
  .execucao .subtitulo-exec { margin: 2px 0 6px; font-size: 12.5px; color: var(--text-muted); }
  .execucao .status { margin: 0 0 6px; font-size: 12px; font-weight: 700; color: var(--status-info); min-height: 1.2em; }

  .console {
    flex: 1 1 auto;
    min-height: 120px;
    background: var(--console-bg);
    color: var(--console-fg);
    border-radius: 4px;
    padding: 10px 12px;
    font-family: "Consolas", "Cascadia Mono", monospace;
    font-size: 12px;
    line-height: 1.6;
    overflow-y: auto;
    white-space: pre-wrap;
  }
</style>
</head>
<body>

<div class="app">
  <div class="coluna coluna--1">
    <div>
      <h1>✨ Automação da Conformidade</h1>
      <p class="subtitulo">Escolha a planilha e um script para executar</p>
    </div>

    <div class="caixa-destaque">
      <p class="rotulo">Planilha de Controle</p>
      <div class="linha-controle">
        <select id="planilha-select"></select>
        <button class="btn btn--icone" id="planilha-atualizar" title="Atualizar lista">⟳</button>
      </div>
      <p class="ajuda" id="planilha-status"></p>
    </div>

    <div id="coluna1-cards"></div>
  </div>

  <div class="coluna coluna--2">
    <div class="caixa-destaque">
      <p class="rotulo">Comando para abrir o Chrome em modo debug (cole no CMD)</p>
      <div class="linha-controle">
        <input class="campo-comando" id="chrome-comando" readonly>
        <button class="btn btn--outline" id="chrome-copiar">Copiar</button>
      </div>
    </div>

    <div id="coluna2-cards"></div>

    <div class="execucao">
      <h2>Execução do script</h2>
      <p class="subtitulo-exec">Acompanhe aqui o andamento e as mensagens do script em execução</p>
      <p class="status" id="status-execucao"></p>
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

  function definirStatus(texto) {
    document.getElementById("status-execucao").textContent = texto;
  }

  function appendOutput(texto) {
    const console_ = document.getElementById("console");
    console_.textContent += texto;
    console_.scrollTop = console_.scrollHeight;
  }

  function definirExecutando(executando) {
    document.querySelectorAll(".card__acao").forEach((botao) => { botao.disabled = executando; });
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
      status.textContent = "Confira se é a certa antes de executar.";
    } else {
      status.textContent = "Nenhuma planilha encontrada na pasta indicada.";
    }
  }

  async function iniciar() {
    document.getElementById("chrome-comando").value = await window.pywebview.api.obter_comando_chrome();
    document.getElementById("chrome-copiar").addEventListener("click", copiarComandoChrome);
    document.getElementById("planilha-atualizar").addEventListener("click", carregarPlanilhas);

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
        "Automação da Conformidade",
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
