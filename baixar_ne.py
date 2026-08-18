from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time  # para fazer pausa

def executar (navegador, num_NE):

    campo_comando = WebDriverWait(navegador, 30).until(EC.element_to_be_clickable((By.ID, "frmMenu:acessoRapido")))
    campo_comando.send_keys("conne" + Keys.ENTER)
    time.sleep(3)

    navegador.switch_to.default_content()
    iframes = navegador.find_elements(By.TAG_NAME, "iframe")
    navegador.switch_to.frame(iframes[1])  # porque já sei que o elemento está no segundo frame

    campo_NE = WebDriverWait(navegador, 15).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#numeroNE input")))
    campo_NE.clear()
    campo_NE.send_keys(num_NE)
    campo_NE.send_keys(Keys.ENTER)

    botao_imprimir = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, "//span[text()='Imprimir']")))
    navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_imprimir)  # rola a página até o botão, pois ele precisa estar visível na tela para não dar erro
    time.sleep(2)  # aguarda 2s para a rolagem terminar
    botao_imprimir.click()

    botao_confirmar = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class,'ui-button-text') and contains(text(),'Confirmar')]")))
    botao_confirmar.click()
    time.sleep(5)

    navegador.switch_to.window(navegador.current_window_handle)
    navegador.execute_script("window.scrollTo(0, 0);")  # rola a página para o topo