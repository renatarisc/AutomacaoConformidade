import sys
import os
import threading
import queue
import traceback
import tkinter as tk
from tkinter.scrolledtext import ScrolledText

import ttkbootstrap as tb
from ttkbootstrap.constants import *

import escolher_planilha
import relacionar_valor_op
import relacionar_op_ob
import baixar_ob
import anexar_ob
import baixar_anexar_ne

# comando que abre o Chrome em modo de depuração remota (porta 9222), exigido pelos
# scripts que se conectam via options.debugger_address - ver comentário no topo de cada main()
COMANDO_CHROME_DEBUG = r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeSelenium"'

VERDE_EXECUTAR = "#45a37e" # mesmo verde do bootstyle "success" (botão Executar) no tema minty
VERDE_CLARO_AVISO = "#98ccb8" # VERDE_EXECUTAR clareado (45% em direção ao branco), usado nos avisos ⚠

# só os scripts que rodam sozinhos (fazem algo ao serem executados diretamente) entram no menu;
# os módulos auxiliares (só definem funções, chamados por esses scripts) ficam de fora.
# cada um chama main() direto (em vez de abrir um subprocesso) pra funcionar também dentro
# do executável gerado pelo PyInstaller, onde não existe mais um python.exe separado pra chamar
OPCOES = [
    {
        "arquivo": "relacionar_valor_op.py",
        "modulo": relacionar_valor_op,
        "icone": "🔍",
        "titulo": "Relacionar Valor → OP",
        "descricao": "Busca a OP no Siafi pelo valor (ISS/PG), grava na planilha e pinta de Amarelo.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "relacionar_op_ob.py",
        "modulo": relacionar_op_ob,
        "icone": "🔗",
        "titulo": "Relacionar OP → OB",
        "descricao": "Busca a OB no Siafi a partir da OP pintada de Cinza e substitui na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
    {
        "arquivo": "baixar_ob.py",
        "modulo": baixar_ob,
        "icone": "⬇️",
        "titulo": "Baixar OB",
        "descricao": "Baixa do Cara Preta o PDF da OB pintada de Amarelo (via pyautogui).",
        "requisito": "Requer o Sistema já aberto e com foco antes de rodar a automação.",
    },
    {
        "arquivo": "anexar_ob.py",
        "modulo": anexar_ob,
        "icone": "📎",
        "titulo": "Anexar OB e tramitar o processo",
        "descricao": "Anexa no processo o PDF da OB pintada de amarelo e tramita.",
        "requisito": "Abre e loga no Suap sozinho - não precisa preparar nada antes.",
    },
    {
        "arquivo": "baixar_anexar_ne.py",
        "modulo": baixar_anexar_ne,
        "icone": "📥",
        "titulo": "Baixar e anexar NE e tramitar o processo.",
        "descricao": "Baixa do Siafi a NE pintada de Cinza, anexa no processo e tramita.",
        "requisito": "Requer o Chrome já aberto em modo debug e logado no Siafi.",
    },
]

class EscritorFila:
    # arquivo "falso" que faz print() dentro dos scripts cair na fila em vez do console
    def __init__(self, fila):
        self.fila = fila

    def write(self, texto):
        if texto:
            self.fila.put(("linha", texto))

    def flush(self):
        pass

class App(tb.Window):
    def __init__(self):
        super().__init__(title="Automação da Conformidade", themename="minty", size=(1200, 660))
        self.minsize(1150, 520)
        self.place_window_center() # posição/tamanho de restauração, caso o usuário desmaximize depois
        self.state("zoomed") # abre já maximizada

        icone = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icone):
            self.iconbitmap(icone)

        self.fila_saida = queue.Queue()
        self.fila_planilhas = queue.Queue()
        self.executando = False
        self.botoes = []

        # menu à esquerda (não ocupa toda a largura) e a saída da automação à direita,
        # lado a lado, em vez de empilhados - aproveita melhor o espaço horizontal da janela.
        # a esquerda fica num canvas com scroll porque o conteúdo (planilha + 5 cards + chrome debug)
        # é mais alto do que a janela permite mostrar de uma vez só
        canvas_esquerdo = tb.Canvas(self, highlightthickness=0, width=680) # icone(~68) + texto (wraplength 460) + botão (~100) + paddings
        canvas_esquerdo.pack(side="left", fill="y")

        scrollbar_esquerda = tb.Scrollbar(self, orient="vertical", command=canvas_esquerdo.yview)
        scrollbar_esquerda.pack(side="left", fill="y")
        canvas_esquerdo.configure(yscrollcommand=scrollbar_esquerda.set)

        painel_esquerdo = tb.Frame(canvas_esquerdo)
        janela_id = canvas_esquerdo.create_window((0, 0), window=painel_esquerdo, anchor="nw")
        painel_esquerdo.bind("<Configure>", lambda e: canvas_esquerdo.configure(scrollregion=canvas_esquerdo.bbox("all")))
        canvas_esquerdo.bind("<Configure>", lambda e: canvas_esquerdo.itemconfigure(janela_id, width=e.width))

        # roda do mouse funciona sobre a coluna esquerda mesmo sem o cursor estar em cima da scrollbar
        canvas_esquerdo.bind("<Enter>", lambda e: canvas_esquerdo.bind_all("<MouseWheel>", lambda ev: canvas_esquerdo.yview_scroll(-1 * (ev.delta // 120), "units")))
        canvas_esquerdo.bind("<Leave>", lambda e: canvas_esquerdo.unbind_all("<MouseWheel>"))

        painel_direito = tb.Frame(self)
        painel_direito.pack(side="left", fill="both", expand=True)

        tb.Label(painel_esquerdo, text="✨ Automação da Conformidade", font=("Segoe UI", 16, "bold"), bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=(12, 0))
        tb.Label(painel_esquerdo, text="Escolha a planilha e um script para executar", font=("Segoe UI", 10), bootstyle=SECONDARY).pack(anchor="w", padx=20, pady=(0, 4))

        self._montar_seletor_planilha(painel_esquerdo)
        self._montar_menu(painel_esquerdo)
        self._montar_chrome_debug(painel_esquerdo)
        self._montar_saida(painel_direito)

    def _montar_seletor_planilha(self, pai):
        # a Planilha de Controle muda todo mês, sem data certa - em vez de editar o nome fixo em
        # cada script, lista as planilhas da pasta do Drive (escolher_planilha.PASTA_DRIVE_ID) e
        # deixa o usuário escolher aqui qual vale para todos os scripts executados a seguir.
        # contorno verde (mesmo tom do botão Executar) pra diferenciar visualmente dos cards dos
        # scripts - usa tk.Frame puro (não o tb.Frame temático) porque só assim dá pra colorir a
        # borda numa cor exata via highlightbackground
        quadro_borda = tk.Frame(pai, highlightbackground=VERDE_EXECUTAR, highlightcolor=VERDE_EXECUTAR, highlightthickness=2, bd=0)
        quadro_borda.pack(fill="x", padx=20, pady=(12, 3))

        quadro = tb.Frame(quadro_borda, padding=8)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(0, weight=1)

        tb.Label(quadro, text="Planilha de Controle", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")

        self.planilha_var = tb.StringVar(value="")
        self.combo_planilha = tb.Combobox(quadro, textvariable=self.planilha_var, state="readonly", font=("Segoe UI", 9))
        self.combo_planilha.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        self.botao_atualizar_planilhas = tb.Button(quadro, text="🔄", bootstyle="secondary-outline", width=3, command=self._atualizar_planilhas)
        self.botao_atualizar_planilhas.grid(row=1, column=1, sticky="e", padx=(6, 0), pady=(4, 0))

        self.status_planilha_var = tb.StringVar(value="")
        tb.Label(quadro, textvariable=self.status_planilha_var, font=("Segoe UI", 8), bootstyle=SECONDARY, wraplength=460, justify="left").grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))

        self._atualizar_planilhas() # já carrega a lista ao abrir o app

    def _atualizar_planilhas(self):
        self.combo_planilha.configure(state="disabled")
        self.botao_atualizar_planilhas.configure(state="disabled")
        self.status_planilha_var.set("⏳ Carregando planilhas da pasta do Drive...")

        thread = threading.Thread(target=self._carregar_planilhas, daemon=True)
        thread.start()
        self.after(100, self._checar_fila_planilhas)

    def _carregar_planilhas(self):
        # roda numa thread separada pra não travar a janela enquanto espera a API do Drive
        try:
            arquivos = escolher_planilha.listar()
            self.fila_planilhas.put(([arquivo["name"] for arquivo in arquivos], None))
        except Exception as e:
            self.fila_planilhas.put(([], str(e)))

    def _checar_fila_planilhas(self):
        try:
            nomes, erro = self.fila_planilhas.get_nowait()
        except queue.Empty:
            self.after(100, self._checar_fila_planilhas)
            return

        self.combo_planilha.configure(state="readonly")
        self.botao_atualizar_planilhas.configure(state="normal")

        if erro:
            self.status_planilha_var.set(f"❌ Erro ao listar planilhas: {erro}")
            return

        self.combo_planilha.configure(values=nomes)
        if nomes:
            self.planilha_var.set(nomes[0]) # lista já vem da mais recente pra mais antiga
            self.status_planilha_var.set(f"{len(nomes)} planilha(s) encontrada(s) - confira se é a certa antes de executar.")
        else:
            self.status_planilha_var.set("Nenhuma planilha encontrada na pasta indicada.")

    def _montar_menu(self, pai):
        for opcao in OPCOES:
            # sem bootstyle no frame: fundo fica igual ao da janela, então os Labels (também sem bootstyle
            # de cor) não ficam com texto claro sobre fundo claro - só a borda marca o "card"
            quadro = tb.Frame(pai, padding=8, borderwidth=1, relief="solid")
            quadro.pack(fill="x", padx=20, pady=3)
            # minsize fixo: cada card tem sua própria grade, e emojis diferentes (ex: 📎, 18px)
            # renderizam mais estreitos que os outros (🔍🔗⬇️📥, 37px cada, medido com Font.measure)
            # na fonte - sem isso, a coluna do ícone varia de card pra card e o texto não começa
            # alinhado entre eles. 56 > 37+12(padx) garante que o piso vale pra todos igualmente
            quadro.columnconfigure(0, minsize=56)
            quadro.columnconfigure(1, weight=1)

            tb.Label(quadro, text=opcao["icone"], font=("Segoe UI Emoji", 20)).grid(row=0, column=0, rowspan=3, padx=(0, 12), sticky="n")

            tb.Label(quadro, text=opcao["titulo"], font=("Segoe UI", 10, "bold"), wraplength=460, justify="left").grid(row=0, column=1, sticky="w")
            tb.Label(quadro, text=opcao["descricao"], font=("Segoe UI", 9), bootstyle=SECONDARY, wraplength=460, justify="left").grid(row=1, column=1, sticky="w", pady=(1, 0))
            tb.Label(quadro, text=f"⚠ {opcao['requisito']}", font=("Segoe UI", 8), foreground=VERDE_CLARO_AVISO, wraplength=460, justify="left").grid(row=2, column=1, sticky="w", pady=(1, 0))

            botao = tb.Button(quadro, text="Executar ▶", bootstyle="success", width=13, command=lambda o=opcao: self.executar(o))
            botao.grid(row=0, column=2, rowspan=3, sticky="e", padx=(6, 0))
            self.botoes.append(botao)

    def _montar_chrome_debug(self, pai):
        # os scripts "requer Chrome em modo debug" (ver OPCOES) precisam desse comando rodado
        # no CMD antes - fica aqui pra copiar com um clique, sem precisar abrir o PyCharm.
        # mesmo contorno verde do card da Planilha de Controle, pra sinalizar que também vale
        # para todos os scripts (não é um card de script específico)
        quadro_borda = tk.Frame(pai, highlightbackground=VERDE_EXECUTAR, highlightcolor=VERDE_EXECUTAR, highlightthickness=2, bd=0)
        quadro_borda.pack(fill="x", padx=20, pady=3)

        quadro = tb.Frame(quadro_borda, padding=8)
        quadro.pack(fill="both", expand=True)
        quadro.columnconfigure(0, weight=1)

        tb.Label(quadro, text="Comando para abrir o Chrome em modo debug (cole no CMD)", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        campo = tb.Entry(quadro, font=("Consolas", 9))
        campo.insert(0, COMANDO_CHROME_DEBUG)
        campo.configure(state="readonly")
        campo.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        tb.Button(quadro, text="Copiar", bootstyle="secondary-outline", width=13, command=self._copiar_comando_chrome).grid(row=1, column=1, sticky="e", padx=(14, 0), pady=(4, 0))

    def _copiar_comando_chrome(self):
        self.clipboard_clear()
        self.clipboard_append(COMANDO_CHROME_DEBUG)

    def _montar_saida(self, pai):
        tb.Label(pai, text="Execução do script", font=("Segoe UI", 16, "bold"), bootstyle=PRIMARY).pack(anchor="w", padx=(12, 20), pady=(12, 0))
        tb.Label(pai, text="Acompanhe aqui o andamento e as mensagens do script em execução", font=("Segoe UI", 10), bootstyle=SECONDARY).pack(anchor="w", padx=(12, 20), pady=(0, 4))

        self.status_var = tb.StringVar(value="")
        tb.Label(pai, textvariable=self.status_var, font=("Segoe UI", 9, "bold"), bootstyle=INFO).pack(anchor="w", padx=(12, 20), pady=(6, 0))

        self.saida_texto = ScrolledText(
            pai, font=("Consolas", 9),
            background="#1e1e1e", foreground="#e6e6e6", insertbackground="#e6e6e6",
            borderwidth=0, highlightthickness=0,
        )
        self.saida_texto.pack(fill="both", expand=True, padx=(12, 20), pady=(4, 12))
        self.saida_texto.configure(state="disabled")

    def executar(self, opcao):
        if self.executando:
            return # já tem algo rodando

        nome_planilha = self.planilha_var.get()
        if not nome_planilha:
            self.status_var.set("⚠ Escolha a Planilha de Controle antes de executar.")
            return

        self.executando = True
        for botao in self.botoes:
            botao.configure(state="disabled")

        self.status_var.set(f"⏳ Executando {opcao['arquivo']} (planilha: {nome_planilha})...")
        self._limpar_saida()

        thread = threading.Thread(target=self._rodar_main, args=(opcao, nome_planilha), daemon=True)
        thread.start()
        self.after(100, self._checar_fila)

    def _limpar_saida(self):
        self.saida_texto.configure(state="normal")
        self.saida_texto.delete("1.0", "end")
        self.saida_texto.configure(state="disabled")

    def _rodar_main(self, opcao, nome_planilha):
        # troca a saída padrão só durante a chamada, pra capturar os print()s do script
        # e mostrar ao vivo na caixa de texto, igual acontecia com o subprocess antes
        saida_original, erro_original = sys.stdout, sys.stderr
        escritor = EscritorFila(self.fila_saida)
        sys.stdout = escritor
        sys.stderr = escritor
        codigo_retorno = 0

        try:
            opcao["modulo"].main(nome_planilha)
        except SystemExit as e:
            # algum script chama sys.exit() no próprio fluxo de erro (ex: baixar_anexar_ne.py);
            # sem isso aqui, isso encerraria a janela inteira em vez de só marcar erro
            codigo_retorno = e.code if isinstance(e.code, int) else 1
        except Exception:
            traceback.print_exc() # cai no escritor (stderr redirecionado), aparece na caixa de saída
            codigo_retorno = 1
        finally:
            sys.stdout, sys.stderr = saida_original, erro_original

        self.fila_saida.put(("fim", codigo_retorno))

    def _checar_fila(self):
        try:
            while True:
                tipo, valor = self.fila_saida.get_nowait()
                if tipo == "linha":
                    self.saida_texto.configure(state="normal")
                    self.saida_texto.insert("end", valor)
                    self.saida_texto.see("end")
                    self.saida_texto.configure(state="disabled")
                elif tipo == "fim":
                    self._finalizar(valor)
                    return # terminou, não reagenda
        except queue.Empty:
            pass

        self.after(100, self._checar_fila)

    def _finalizar(self, codigo_retorno):
        self.executando = False
        for botao in self.botoes:
            botao.configure(state="normal")

        if codigo_retorno == 0:
            self.status_var.set("✅ Concluído.")
        else:
            self.status_var.set(f"❌ Terminou com erro (código {codigo_retorno}).")


if __name__ == "__main__":
    App().mainloop()
