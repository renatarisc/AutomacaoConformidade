from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from google.oauth2.service_account import Credentials
import gspread # para manipular as planilhas do Drive
import time # para fazer pausa

import carregar_cores_planilha
import credenciais_suap
import escolher_planilha
import ler_planilha
import pintar_celula_planilha
import encaminhar_processo

AMARELO_CLARO_1 = (1, 217 / 255, 102 / 255) # mesmo amarelo usado nos demais scripts do fluxo
VERMELHO = (1, 0, 0)
BRANCO = (1, 1, 1) # sinaliza que a OB já foi anexada e o processo tramitado - não precisa de mais nenhuma ação

def coluna(dados, nome):
    # o texto do cabeçalho da aba NS já mudou de caixa mais de uma vez neste projeto
    # (ex: "Processo" virou "PROCESSO"), então resolve o nome da coluna sem depender de
    # maiúsculas/minúsculas nem de espaços nas pontas - mesmo critério do preencher_planilha_ns.py
    alvo = nome.strip().casefold()
    for atual in dados.columns:
        if str(atual).strip().casefold() == alvo:
            return atual
    raise KeyError(f"coluna {nome!r} não encontrada no cabeçalho da aba NS: {list(dados.columns)}")

def cor_bate(cor_celula, cor_alvo, tolerancia=0.01):
    # a API do Sheets guarda a cor com menos precisão do que o float do Python,
    # então uma comparação exata (==) quase nunca bate mesmo com a cor visualmente igual
    if not cor_celula:
        return False
    return all(abs(c - alvo) < tolerancia for c, alvo in zip(cor_celula, cor_alvo))

def executar(navegador, var_OB):

    botao_upload_externo = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, "//a[contains(., 'Upload de documento externo')]")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_upload_externo)  # rola a página até o botão, pois ele precisa estar visível na tela para não dar erro
    time.sleep(1)  # aguarda um instante para a rolagem terminar
    botao_upload_externo.click()

    OB = str(var_OB)
    campo_arquivo = fr"C:\Users\renat\Downloads\{OB}.pdf"
    navegador.find_element(By.ID, "id_arquivo").send_keys(campo_arquivo)

    select_conferencia = Select(navegador.find_element(By.ID, "id_tipo_conferencia"))
    select_conferencia.select_by_value("4")  # = Documento Original

    # o carregamento do campo "Tipo" é feito com AJAX, por isso uma solução diferente
    campo_tipo_select = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[id^='select2-tipo_']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_tipo_select)
    time.sleep(0.3)
    try:
        campo_tipo_select.click()
    except Exception:
        # Chrome 152+ é mais rígido com "element click intercepted" (o <b> da setinha do select2
        # recebe o clique) - força via JS, que dispara o mesmo handler de abertura do select2
        navegador.execute_script("arguments[0].click();", campo_tipo_select)
    time.sleep(0.5)  # pequena pausa para garantir que a animação de abertura do menu terminou
    texto = "Ordem Bancária (OB)"
    actions = ActionChains(navegador)  # envia o texto DIRETAMENTE para a parte focada usando ActionChains
    actions.send_keys(texto)
    actions.perform()
    time.sleep(1.5)  # aguarda 1.5s para o AJAX carregar as opções buscadas do servidor do SUAP
    actions.send_keys(Keys.ENTER)  # envia o comando ENTER para confirmar a seleção do primeiro resultado filtrado
    actions.perform()

    navegador.find_element(By.ID, "id_assunto").send_keys(OB)

    select_nivel_acesso = Select(navegador.find_element(By.ID, "id_nivel_acesso"))
    select_nivel_acesso.select_by_visible_text("Restrito")

    WebDriverWait(navegador, 20, ignored_exceptions=(StaleElementReferenceException,)).until(
        lambda d: "7" in [
            op.get_attribute("value")
            for op in Select(d.find_element(By.ID, "id_hipotese_legal")).options
        ]
    )
    campo_hipotese_legal = navegador.find_element(By.ID, "id_hipotese_legal")  # localiza novamente o select, pois ele pode ter sido recriado pelo AJAX
    navegador.execute_script("arguments[0].scrollIntoView({block:'center'});", campo_hipotese_legal)  # rola a página até o campo
    Select(campo_hipotese_legal).select_by_value("7")

    botao_salvar = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Salvar']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_salvar)
    navegador.execute_script("arguments[0].click();", botao_salvar)  # força o clique via JavaScript, caso o clique nativo ainda esteja bloqueado por modais ocultos

    time.sleep(2)  # pq vai carregar outa página HTML

    # ------- Assina o documento anexado -------
    select_perfil = Select(navegador.find_element(By.ID, "id_papel"))
    select_perfil.select_by_value("5260")  # = ASSISTENTE EM ADMINISTRACAO
    navegador.find_element(By.ID, "id_senha").send_keys(credenciais_suap.senha())
    navegador.find_element(By.XPATH, "//input[@value='Assinar Documento']").click()  # botão Assinar Documento

    time.sleep(3)

def main(nome_planilha=None):
    # nome_planilha: passado pelo gui.py com a planilha escolhida na interface; rodando o
    # script sozinho (sem gui.py), usa escolher_planilha.NOME_PLANILHA_PADRAO
    # ------- Acessa a Planilha de Controle da Conformidade (mensal) -------
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credenciais = Credentials.from_service_account_file("credenciais.json", scopes=SCOPES) # nome do arq dentro da pasta do Projeto
    gc = gspread.authorize(credenciais)
    planilha = gc.open(nome_planilha or escolher_planilha.NOME_PLANILHA_PADRAO)
    aba = planilha.worksheet("NS")

    dados = ler_planilha.carregar_registros(aba, preencher_mescladas=True) # 1ª linha vira cabeçalho, resto como texto; ignora colunas sem nome; replica célula mesclada (ex: "Processo" de diárias) para todo o bloco
    cores = carregar_cores_planilha.executar(aba) # chama a def

    # a cor amarelo claro 1 sinaliza que a OB (ISS ou PG) já foi baixada pelo baixar_ob.py e ainda precisa ser
    # anexada ao processo - coluna calculada pelo nome do cabeçalho, não fixa, pra não quebrar se a coluna mudar de lugar
    col_processo = coluna(dados, "Processo")
    col_despacho = coluna(dados, "DESPACHO PROC") # coluna renomeada de "DESPACHO" na planilha
    COLUNAS_OB = [
        {"coluna": dados.columns.get_loc(coluna(dados, "OB ISS")) + 1, "nome": coluna(dados, "OB ISS")},
        {"coluna": dados.columns.get_loc(coluna(dados, "OB PG")) + 1, "nome": coluna(dados, "OB PG")},
    ]

    COLUNA_TRAMITADO = dados.columns.get_loc(coluna(dados, "TRAMITADO")) + 1 # calculado pelo nome do cabeçalho, não fixo, pra não quebrar se a coluna mudar de lugar
    COLUNA_PROCESSO = dados.columns.get_loc(col_processo) + 1 # p/ pintar a célula do processo de branco depois de tramitar

    # agrupa por NÚMERO DE PROCESSO, não por linha: além de um processo poder ter OB ISS e OB PG
    # pendentes na mesma linha, várias linhas podem ser do mesmo processo (ex: várias diárias numa
    # célula "Processo" mesclada). O processo é buscado uma vez, todas as OBs pendentes dele são
    # anexadas, e só então ele é tramitado uma única vez - tramitar tira o processo da fila do
    # usuário, então tramitar por linha deixaria as OBs das outras linhas sem como anexar.
    processos = {} # numero_processo -> {"processo", "despacho", "obs": [...], "linhas": set()}
    for linha in dados.index:
        linha_planilha = linha + 2 # linha do DataFrame começa em 0, a planilha em 2 (cabeçalho na linha 1)

        obs_pendentes = [
            {"linha_planilha": linha_planilha, "coluna": info["coluna"], "OB": str(dados.loc[linha, info["nome"]]).strip()}
            for info in COLUNAS_OB
            if cor_bate(cores.get((linha_planilha, info["coluna"])), AMARELO_CLARO_1)
            and "OB" in str(dados.loc[linha, info["nome"]]) # a mesma cor amarela também marca uma OP ainda pendente de assinatura/conversão em OB - só anexa quem já é OB de fato
        ]
        if not obs_pendentes:
            continue

        numero_processo = str(dados.loc[linha, col_processo]).strip()
        grupo = processos.setdefault(numero_processo, {"processo": numero_processo, "despacho": "", "obs": [], "linhas": set()})
        grupo["obs"].extend(obs_pendentes)
        grupo["linhas"].add(linha_planilha)
        despacho_linha = str(dados.loc[linha, col_despacho]).strip()
        if despacho_linha and not grupo["despacho"]: # 1º despacho não vazio do grupo (deve ser o mesmo em todas as linhas do processo)
            grupo["despacho"] = despacho_linha

    lista_processo = list(processos.values()) # dict preserva a ordem de inserção -> mantém a ordem da planilha

    # ------- Abre o Chrome maximizado -------
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # detach=True p/ impedir que o Selenium feche o navegador ao terminar a execução ou quando ocorrer um erro
    navegador_suap = webdriver.Chrome(options=options)  # navegador controlado pelo Selenium | o Chrome tem mais compatibilidade com os sites
    navegador_suap.maximize_window()

    # ------- Entra na tela de login do Suap -------
    navegador_suap.get("https://suap.iff.edu.br/accounts/login/?next=/")  # pode ser o caminho de um arquivo local
    # navegador_suap.get("http://suap.dev.iff.edu.br/accounts/login/?next=/")  # ambiente de homologação (dev)

    # ------- Faz o login no Suap -------
    navegador_suap.find_element(By.ID, "id_username").send_keys(credenciais_suap.usuario())
    navegador_suap.find_element(By.ID, "id_password").send_keys(credenciais_suap.senha() + Keys.ENTER)
    time.sleep(10)

    for processo in lista_processo:

        # ------- Localiza o processo e carrega na página -------
        campo_busca_rapida = WebDriverWait(navegador_suap, 30).until(EC.element_to_be_clickable((By.NAME, "q")))
        campo_busca_rapida.clear()
        campo_busca_rapida.send_keys(processo["processo"] + Keys.ENTER)

        todas_anexadas = True
        for info in processo["obs"]:
            try:
                executar(navegador_suap, info["OB"])
                pintar_celula_planilha.executar(aba, info["linha_planilha"], info["coluna"], BRANCO)
            except Exception as e:
                print(f"Erro ao anexar a OB {info['OB']}: {e}")
                pintar_celula_planilha.executar(aba, info["linha_planilha"], info["coluna"], VERMELHO)
                todas_anexadas = False

        # só tramita o processo se todas as OBs pendentes dele foram anexadas com sucesso;
        # senão o processo sai da fila do usuário antes de dar pra tentar de novo a que falhou
        if todas_anexadas and processo["despacho"]:
            try:
                encaminhar_processo.executar(navegador_suap, processo["despacho"])
                for linha_planilha in sorted(processo["linhas"]): # marca todas as linhas do processo (uma diária mesclada ocupa várias)
                    aba.update_cell(linha_planilha, COLUNA_TRAMITADO, "OK") # só marca se a tramitação realmente aconteceu
                    pintar_celula_planilha.executar(aba, linha_planilha, COLUNA_PROCESSO, BRANCO) # célula do processo (ciano) -> branco: processo concluído
            except Exception as e:
                print(f"Erro ao tramitar o processo {processo['processo']}: {e}")

        navegador_suap.get("https://suap.iff.edu.br/")  # volta para a tela de início, onde tem o campo Busca rápida
        # navegador_suap.get("http://suap.dev.iff.edu.br/")  # ambiente de homologação (dev)
        time.sleep(10)

if __name__ == "__main__":
    main()
