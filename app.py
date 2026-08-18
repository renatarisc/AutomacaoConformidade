import subprocess
import sys

import streamlit as st

# só os scripts que rodam sozinhos (fazem algo ao serem executados diretamente) entram no menu;
# os módulos auxiliares (só definem funções, chamados por esses scripts) ficam de fora
OPCOES = [
    {
        "arquivo": "relacionar_valor_op.py",
        "titulo": "Relacionar Valor → OP",
        "descricao": "Busca a OP no Siafi pelo valor (ISS/PG) e grava na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
    {
        "arquivo": "relacionar_op_ob.py",
        "titulo": "Relacionar OP → OB",
        "descricao": "Busca a OB no Siafi a partir da OP e substitui na planilha.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
    {
        "arquivo": "baixar_ob.py",
        "titulo": "Baixar OBs",
        "descricao": "Baixa o PDF de cada OB pendente direto do Cara Preta (via pyautogui).",
        "requisito": "Requer o Sistema já aberto e com foco na tela antes de rodar.",
    },
    {
        "arquivo": "anexar_ob.py",
        "titulo": "Anexar OBs no Suap",
        "descricao": "Anexa o PDF da OB no processo, assina e tramita no Suap.",
        "requisito": "Abre e loga no Suap sozinho - não precisa preparar nada antes.",
    },
    {
        "arquivo": "baixar_anexar_ne.py",
        "titulo": "Baixar e anexar NEs",
        "descricao": "Baixa a NE no Siafi e anexa/assina/tramita no Suap.",
        "requisito": "Requer o Chrome já aberto em modo debug (porta 9222), logado no Siafi.",
    },
]

st.set_page_config(page_title="Automação Conformidade", page_icon="📋", layout="centered")
st.title("📋 Automação Conformidade")
st.caption("Escolha um script para executar")

if "executando" not in st.session_state:
    st.session_state.executando = None

for opcao in OPCOES:
    with st.container(border=True):
        col_texto, col_botao = st.columns([4, 1])
        with col_texto:
            st.markdown(f"**{opcao['titulo']}**")
            st.caption(opcao["descricao"])
            st.caption(f"⚠️ {opcao['requisito']}")
        with col_botao:
            if st.button("Executar", key=opcao["arquivo"], disabled=st.session_state.executando is not None):
                st.session_state.executando = opcao["arquivo"]
                st.rerun()

if st.session_state.executando:
    arquivo = st.session_state.executando
    st.divider()
    st.subheader(f"Executando {arquivo}...")
    saida = st.empty()
    texto_saida = ""

    processo = subprocess.Popen(
        [sys.executable, arquivo],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )

    for linha in processo.stdout:
        texto_saida += linha
        saida.code(texto_saida)

    processo.wait()

    if processo.returncode == 0:
        st.success(f"{arquivo} concluído.")
    else:
        st.error(f"{arquivo} terminou com erro (código {processo.returncode}).")

    st.session_state.executando = None
    if st.button("Voltar ao menu"):
        st.rerun()
