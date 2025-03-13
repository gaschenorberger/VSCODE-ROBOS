from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import openpyxl
import openpyxl.workbook
import time

planilha_caminhos = openpyxl.load_workbook(r"SistemaZ\ProcurarCnpj\nomeEmpresas.xlsx")
pagiCaminhos = planilha_caminhos['Planilha1']

def iniciar_navegador(com_debugging_remoto=True):
    chrome_driver_path = ChromeDriverManager().install()
    chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
    
    #print(f"ChromeDriver path: {chrome_driver_executable}")
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

def procurarCnpj():

    indices = [62, 64, 68]

    for indice in indices:
        #print(indice)

        time.sleep(1)

        btnEmpresas = navegador.find_element(By.XPATH, '//*[@id="empresa"]/button')
        btnEmpresas.click()
        
        time.sleep(2)

        empresa = navegador.find_element(By.XPATH, f'//*[@id="empresa_panel"]/ul/li[{indice}]')
        empresa.click()

        time.sleep(1)

        btnPesquisa = navegador.find_element(By.XPATH, '//*[@id="j_idt80"]')
        btnPesquisa.click()

        time.sleep(1)

        nomeEmpresa = navegador.find_element(By.XPATH, '//*[@id="empresa_input"]').get_attribute("value")

        cnpjs = navegador.find_element(By.XPATH, '//*[@id="resultado_pesquisa_data"]/tr[1]/td[2]')
        cnpj = cnpjs.text

        print(cnpj, '-', nomeEmpresa)
    

    """empresas = navegador.find_elements(By.XPATH, '//*[@id="empresa_panel"]/ul/li')
    
    for empresa in empresas:
        print(empresa.text)"""

procurarCnpj()
