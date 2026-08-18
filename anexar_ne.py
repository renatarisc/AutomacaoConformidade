from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import time # para fazer pausa
import glob

def executar(navegador, var_NE):

    botao_upload_externo = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, "//a[contains(., 'Upload de documento externo')]")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_upload_externo)  # rola a página até o botão, pois ele precisa estar visível na tela para não dar erro
    time.sleep(1)  # aguarda um instante para a rolagem terminar
    botao_upload_externo.click()

    NE = str(var_NE)
    busca = fr"C:\Users\renat\Downloads\*{NE}*.pdf"
    arquivo = glob.glob(busca) # busca por TODOS os arquivos com o nome especificado e cria uma lista

    if arquivo:
        arq = arquivo[0] # pega o primeiro arquivo encontrado na lista
        navegador.find_element(By.ID, "id_arquivo").send_keys(arq)

    select_conferencia = Select(navegador.find_element(By.ID, "id_tipo_conferencia"))
    select_conferencia.select_by_value("4") # = Documento Original

    # o carregamento do campo "Tipo" é feito com AJAX, por isso uma solução diferente
    campo_tipo_select = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[id^='select2-tipo_']")))
    campo_tipo_select.click()
    time.sleep(0.5) # pequena pausa para garantir que a animação de abertura do menu terminou
    texto = "Nota de Empenho (NE)"
    actions = ActionChains(navegador) # envia o texto DIRETAMENTE para a parte focada usando ActionChains
    actions.send_keys(texto)
    actions.perform()
    time.sleep(3) # aguarda 1.5s para o AJAX carregar as opções buscadas do servidor do SUAP
    actions.send_keys(Keys.ENTER) # envia o comando ENTER para confirmar a seleção do primeiro resultado filtrado
    actions.perform()

    navegador.find_element(By.ID, "id_assunto").send_keys(NE)

    select_nivel_acesso = Select(navegador.find_element(By.ID, "id_nivel_acesso"))
    select_nivel_acesso.select_by_visible_text("Público")

    botao_salvar = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Salvar']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_salvar)
    navegador.execute_script("arguments[0].click();", botao_salvar) # força o clique via JavaScript, caso o clique nativo ainda esteja bloqueado por modais ocultos

    time.sleep(2) # pq vai carregar outa página HTML

    # ------- Assina o documento anexado -------
    # select_perfil = Select(navegador.find_element(By.ID, "id_papel"))
    select_perfil = Select(WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.ID, "id_papel"))))
    select_perfil.select_by_value("5260") # = ASSISTENTE EM ADMINISTRACAO
    navegador.find_element(By.ID, "id_senha").send_keys("Aj250104!")
    navegador.find_element(By.XPATH, "//input[@value='Assinar Documento']").click() # botão Assinar Documento

    time.sleep(5) # porque a assinatura do documento sempre demora um pouco
