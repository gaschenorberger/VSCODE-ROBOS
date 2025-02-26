from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.common.exceptions import NoSuchElementException 
import pyautogui
import os

def iniciar_navegador(com_debugging_remoto=True):
    #chrome_driver_path = ChromeDriverManager().install()
    chrome_driver_path = r"C:\Users\mateus.silva\.wdm\drivers\chromedriver\win64\131.0.6778.204\chromedriver-win32\chromedriver.exe"
    chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
    
    print(f"ChromeDriver path: {chrome_driver_executable}")
    if not os.path.isfile(chrome_driver_executable):
        raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

    service = Service(executable_path=chrome_driver_executable)
    
    chrome_options = Options()
    if com_debugging_remoto:
        remote_debugging_port = 9222
        chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    return navegador

navegador = iniciar_navegador(com_debugging_remoto=True)


def pesquisa():
        # 2019,2025
        for ano in range (2022,2026):
                for mes in range (1,14):
#                                                                               //*[@id="PeriodoApuracaoPesquisa"]
                        INP_periodoApuracao = navegador.find_element(By.XPATH, '//*[@id="PeriodoApuracaoPesquisa"]')
                        if mes == 13:
                                INP_periodoApuracao.send_keys(f"{ano}")
                                
                        else:
                                INP_periodoApuracao.send_keys(f"{mes}{ano}")

                        BT_pesquisar = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/form/div/section/button')
                        BT_pesquisar.click()

                        try:
                                BT_baixar = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/form/div/div[3]/div/a')
                                BT_baixar.click()

                        except NoSuchElementException:
                                if ano != 2025:
                                        INP_periodoApuracao = navegador.find_element(By.XPATH, '//*[@id="PeriodoApuracaoPesquisa"]')
                                        INP_periodoApuracao.clear()
                                        continue
                                else:
                                        break

                        INP_periodoApuracao = navegador.find_element(By.XPATH, '//*[@id="PeriodoApuracaoPesquisa"]')
                        INP_periodoApuracao.clear()



pyautogui.keyDown('alt')
pyautogui.press(['tab'])
pyautogui.keyUp('alt')
time.sleep(0.5)
pyautogui.click(1140,220)
pesquisa()
