from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import pyautogui

# ----------------------------------------------------------------------------------------------

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

def mover_arquivos_xml():

    pasta_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS"
    pasta_destino = r"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\XML"

    # Garante que o destino existe
    os.makedirs(pasta_destino, exist_ok=True)

    # Percorre os arquivos da pasta de origem
    for arquivo in os.listdir(pasta_origem):
        if arquivo.lower().endswith('.xml'):
            caminho_origem = os.path.join(pasta_origem, arquivo)
            caminho_destino = os.path.join(pasta_destino, arquivo)
            try:
                shutil.move(caminho_origem, caminho_destino)
            except Exception as e:
                print(f'Erro ao mover {arquivo}: {e}')

    print("Todos os arquivos movidos")

# Caminho para o arquivo de notas
caminho_arquivo = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\NOTAS.txt'
arquivo_erros = r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\BUSCA_XML\erros_download.txt'

with open(caminho_arquivo, 'r') as file:
    chaves = [linha.strip() for linha in file.readlines()]

inicio = 0 
fim = None 

chaves = chaves[inicio:fim]
chaves_com_erro = []

navegador = iniciar_navegador()

def baixarXml():
    wait = WebDriverWait(navegador, 10)
    for chave in chaves:

        inputChave = navegador.find_element(By.XPATH, '//*[@id="chave"]')
        inputChave.clear()
        inputChave.send_keys(chave), time.sleep(0.5)

        btnBuscar = navegador.find_element(By.XPATH, '//*[@id="form_one"]/button')
        btnBuscar.click()

        btnBaixar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modalNFe"]/div/div/div[2]/div/div[3]/p[2]/a')))
        btnBaixar.click(), time.sleep(0.5)

        btnFechar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="modalNFe"]/div/div/div[1]/button')))
        btnFechar.click(), time.sleep(0.5)
        

baixarXml()
