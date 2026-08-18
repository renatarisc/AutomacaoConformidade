import subprocess
import sys
import os
import threading
import queue
from tkinter.scrolledtext import ScrolledText

import ttkbootstrap as tb
from ttkbootstrap.constants import *

# só os scripts que rodam sozinhos (fazem algo ao serem executados diretamente) entram no menu;
# os módulos auxiliares (só definem funções, chamados por esses scripts) ficam de fora
OPCOES = [
    {
        "arquivo": "relacionar_valor_op.py",
        "icone": "🔍",
        "titulo": "Relacionar Valor → OP",
        "descricao": "Busca a OP no Siafi pelo valor (ISS/PG) e grava na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
    {
        "arquivo": "relacionar_op_ob.py",
        "icone": "🔗",
        "titulo": "Relacionar OP → OB",
        "descricao": "Busca a OB no Siafi a partir da OP e substitui na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
    {
        "arquivo": "baixar_ob.py",
        "icone": "⬇️",
        "titulo": "Baixar OBs",
        "descricao": "Baixa o PDF de cada OB pendente direto do Cara Preta (via pyautogui).",
        "requisito": "Requer o Sistema já aberto e com foco na tela antes de rodar.",
    },
    {
        "arquivo": "anexar_ob.py",
        "icone": "📎",
        "titulo": "Anexar OBs e Tramitar processo",
        "descricao": "Anexa o PDF da OB no processo e tramita.",
        "requisito": "Abre e loga no Suap sozinho - não precisa preparar nada antes.",
    },
    {
        "arquivo": "baixar_anexar_ne.py",
        "icone": "📥",
        "titulo": "Baixar e Anexar NEs. Tramitar quando necessário.",
        "descricao": "Baixa a NE no Siafi, anexa e tramita no Suap.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
]

class App(tb.Window):
    def __init__(self):
        super().__init__(title="Automação da Conformidade", themename="minty", size=(700, 620))
        self.minsize(620, 520)

        icone = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icone):
            self.iconbitmap(icone)

        self.fila_saida = queue.Queue()
        self.processo_atual = None
        self.botoes = []

        self._montar_menu()
        self._montar_saida()

    def _montar_menu(self):
        tb.Label(self, text="✨ Automação da Conformidade", font=("Segoe UI", 16, "bold"), bootstyle=PRIMARY).pack(anchor="w", padx=20, pady=(12, 0))
        tb.Label(self, text="Escolha um script para executar", font=("Segoe UI", 10), bootstyle=SECONDARY).pack(anchor="w", padx=20, pady=(0, 8))

        for opcao in OPCOES:
            # sem bootstyle no frame: fundo fica igual ao da janela, então os Labels (também sem bootstyle
            # de cor) não ficam com texto claro sobre fundo claro - só a borda marca o "card"
            quadro = tb.Frame(self, padding=8, borderwidth=1, relief="solid")
            quadro.pack(fill="x", padx=20, pady=3)
            quadro.columnconfigure(1, weight=1)

            tb.Label(quadro, text=opcao["icone"], font=("Segoe UI Emoji", 20)).grid(row=0, column=0, rowspan=3, padx=(0, 12), sticky="n")

            tb.Label(quadro, text=opcao["titulo"], font=("Segoe UI", 10, "bold"), wraplength=360, justify="left").grid(row=0, column=1, sticky="w")
            tb.Label(quadro, text=opcao["descricao"], font=("Segoe UI", 9), bootstyle=SECONDARY, wraplength=360, justify="left").grid(row=1, column=1, sticky="w", pady=(1, 0))
            tb.Label(quadro, text=f"⚠ {opcao['requisito']}", font=("Segoe UI", 8), bootstyle="warning", wraplength=360, justify="left").grid(row=2, column=1, sticky="w", pady=(1, 0))

            botao = tb.Button(quadro, text="Executar ▶", bootstyle="success", width=13, command=lambda o=opcao: self.executar(o))
            botao.grid(row=0, column=2, rowspan=3, sticky="e", padx=(14, 0))
            self.botoes.append(botao)

    def _montar_saida(self):
        self.status_var = tb.StringVar(value="")
        tb.Label(self, textvariable=self.status_var, font=("Segoe UI", 9, "bold"), bootstyle=INFO).pack(anchor="w", padx=20, pady=(6, 0))

        self.saida_texto = ScrolledText(
            self, height=8, font=("Consolas", 9),
            background="#1e1e1e", foreground="#e6e6e6", insertbackground="#e6e6e6",
            borderwidth=0, highlightthickness=0,
        )
        self.saida_texto.pack(fill="both", expand=True, padx=20, pady=(4, 12))
        self.saida_texto.configure(state="disabled")

    def executar(self, opcao):
        if self.processo_atual is not None:
            return # já tem algo rodando

        for botao in self.botoes:
            botao.configure(state="disabled")

        self.status_var.set(f"⏳ Executando {opcao['arquivo']}...")
        self._limpar_saida()

        thread = threading.Thread(target=self._rodar_subprocesso, args=(opcao["arquivo"],), daemon=True)
        thread.start()
        self.after(100, self._checar_fila)

    def _limpar_saida(self):
        self.saida_texto.configure(state="normal")
        self.saida_texto.delete("1.0", "end")
        self.saida_texto.configure(state="disabled")

    def _rodar_subprocesso(self, arquivo):
        processo = subprocess.Popen(
            [sys.executable, arquivo],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        self.processo_atual = processo

        for linha in processo.stdout:
            self.fila_saida.put(("linha", linha))

        processo.wait()
        self.fila_saida.put(("fim", processo.returncode))

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
        self.processo_atual = None
        for botao in self.botoes:
            botao.configure(state="normal")

        if codigo_retorno == 0:
            self.status_var.set("✅ Concluído.")
        else:
            self.status_var.set(f"❌ Terminou com erro (código {codigo_retorno}).")


if __name__ == "__main__":
    App().mainloop()
