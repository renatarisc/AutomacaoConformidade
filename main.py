import subprocess
import sys

# só os scripts que rodam sozinhos (fazem algo ao serem executados diretamente) entram no menu;
# anexar_ne.py, baixar_ne.py, carregar_cores_planilha.py, encaminhar_processo.py e pintar_celula_planilha.py
# são módulos auxiliares (só definem funções) chamados por esses scripts, não fazem nada executados sozinhos
OPCOES = [
    {"numero": "1", "arquivo": "relacionar_valor_op.py", "descricao": "Relacionar Valor -> OP (busca a OP no Siafi pelo valor)"},
    {"numero": "2", "arquivo": "relacionar_op_ob.py", "descricao": "Relacionar OP -> OB (busca a OB no Siafi a partir da OP)"},
    {"numero": "3", "arquivo": "baixar_ob.py", "descricao": "Baixar OBs (pyautogui, baixa o PDF da OB no Sistema)"},
    {"numero": "4", "arquivo": "anexar_ob.py", "descricao": "Anexar OBs no Suap (anexa, assina e tramita o processo)"},
    {"numero": "5", "arquivo": "baixar_anexar_ne.py", "descricao": "Baixar e anexar NEs no Suap"},
    {"numero": "6", "arquivo": "preencher_planilha_ro.py", "descricao": "Preencher RO/NC/NE na aba TesteRO a partir das abas de Andamento do processo abertas no Chrome"},
]

def mostrar_menu():
    print("\n===== Automação Suap =====")
    for opcao in OPCOES:
        print(f"{opcao['numero']} - {opcao['descricao']}")
    print("0 - Sair")

def executar_script(caminho_arquivo):
    print(f"\n> Executando {caminho_arquivo}...\n")
    resultado = subprocess.run([sys.executable, caminho_arquivo]) # roda como processo separado, com o mesmo Python do venv
    if resultado.returncode != 0:
        print(f"\n{caminho_arquivo} terminou com erro (código {resultado.returncode}).")
    else:
        print(f"\n{caminho_arquivo} concluído.")

while True:
    mostrar_menu()
    escolha = input("\nEscolha uma opção: ").strip()

    if escolha == "0":
        break

    opcao_escolhida = next((o for o in OPCOES if o["numero"] == escolha), None)
    if opcao_escolhida is None:
        print("Opção inválida.")
        continue

    executar_script(opcao_escolhida["arquivo"])
