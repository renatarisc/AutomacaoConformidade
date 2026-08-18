from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import time  # para fazer pausa

def executar(navegador, var_despacho):
    botao_encaminhar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Encaminhar')))
    botao_encaminhar.click()

    com_despacho = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, 'Com despacho')))
    com_despacho.click()

    campo_despacho = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, 'id_despacho_corpo')))
    campo_despacho.send_keys(var_despacho)

    campo_destino = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[id*='destinatario_setor']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_destino)
    time.sleep(1)
    campo_destino.click()
    time.sleep(2)
    # Agora que o campo de busca existe, digita e seleciona
    campo_busca_setor = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.select2-search__field')))
    campo_busca_setor.send_keys("COFCCI")
    time.sleep(3)
    campo_busca_setor.send_keys(Keys.ENTER)

    # Rola até o Perfil e seleciona
    select_perfil = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, 'id_papel')))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_perfil)
    time.sleep(1)
    perfil = Select(select_perfil)
    perfil.select_by_value("5260")  # = ASSISTENTE EM ADMINISTRACAO

    campo_senha = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.ID, 'id_senha')))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_senha)
    time.sleep(1)
    campo_senha.send_keys("Aj250104!")

    botao_salvar = WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit'][value='Salvar']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_salvar)
    time.sleep(1)
    botao_salvar.click()