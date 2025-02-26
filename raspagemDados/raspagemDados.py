from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import os
import openpyxl
import openpyxl.workbook
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


#----------------------------------------------FUNÇÕES----------------------------------------------------
planilha_caminhos = openpyxl.Workbook() #ARQUIVO XLSX 
pagiCaminhos = planilha_caminhos.active
#pagiCaminhos.title = 'Dados'



def iniciar_navegador(com_debugging_remoto=True):
    chrome_driver_path = ChromeDriverManager().install()
    #chrome_driver_path = r'C:\Users\gabriel.alvise\.wdm\drivers\chromedriver\win64\130.0.6723.91\chromedriver-win32/chromedriver.exe'
    chrome_driver_executable = os.path.join(os.path.dirname(chrome_driver_path), 'chromedriver.exe')
    #print(chrome_driver_path)
    
    if not os.path.isfile(chrome_driver_executable):
        raise FileNotFoundError(f"O ChromeDriver não foi encontrado em {chrome_driver_executable}")

    service = Service(executable_path=chrome_driver_executable)
    
    chrome_options = Options()
    if com_debugging_remoto:
        remote_debugging_port = 9222
        chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    return navegador

def obter_produtos():
    produtos_dados = []

    try:
        WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'productLink')]"))
        )

        produtos = navegador.find_elements(By.XPATH, "//a[contains(@class, 'productLink')]")

        for produto in produtos:
            try:
                nome = produto.find_element(By.TAG_NAME, "h3").get_attribute("innerText").strip()
                
                try:
                    preco = produto.find_element(By.XPATH, ".//span[contains(@class, 'd79c9cf')]").get_attribute("innerText").strip()
                except NoSuchElementException:
                    try:
                        preco = produto.find_element(By.XPATH, ".//span[contains(@class, 'price')]").get_attribute("innerText").strip()
                    except NoSuchElementException:
                        preco = "Preço não encontrado"

                if nome:
                    produtos_dados.append({"Produto": nome, "Preço": preco})
                    print(f"Produto: {nome} | Preço: {preco}")
                else:
                    print("Produto sem nome")

            except NoSuchElementException:
                print("Erro ao obter nome do produto")

    except Exception as e:
        print(f"Erro: {e}")

    return produtos_dados

navegador = iniciar_navegador()
produtos = obter_produtos()

df = pd.DataFrame(produtos)
df.to_excel(r"raspagemDados\produtos.xlsx", index=False)

print("Produtos salvos")