from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException
import os
import time
import pyautogui
import subprocess
import requests
import re
from tqdm import tqdm

# NÃO MEXER NO COMPUTADOR DURANTE A EXECUÇÃO DO ROBÔ

#------------------------------------------------------DADOS ABERTOS----------------------------------------------------------

def iniciar_navegador(com_debugging_remoto=True):
    #CASO ESTIVER CONECTADO NA INTERNET DA REDE, DEVE-SE COLOCAR O CAMINHO ESPECÍFICO DO SELENIUM, POR ALGUM MOTIVO A REDE ESTA BLOQUEANDO O "chrome_driver_path = ChromeDriverManager().install()"
    
    chrome_driver_path = ChromeDriverManager().install()
    # chrome_driver_path = r'C:\Users\gabriel.alvise\.wdm\drivers\chromedriver\win64\130.0.6723.91\chromedriver-win32/chromedriver.exe'
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

def abrir_chrome():
    #CHROME ESPECÍFICO PARA UTILIZAÇÃO DO SELENIUM
    comando = r'start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile'
    subprocess.Popen(comando, shell=True)

    time.sleep(2)

    url = 'https://dados.gov.br/dados/conjuntos-dados/cadastro-nacional-da-pessoa-juridica---cnpj'
    pyautogui.write(url)
    time.sleep(0.5)

    pyautogui.press('enter')

def aguardando_download(pasta_download_origem, nome_arquivo):
    #pasta_download_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS"
    #nome_arquivo = 'arquivo_completo.zip'

    while any([filename.endswith('.crdownload') or filename.endswith('.part') 
               for filename in os.listdir(pasta_download_origem)]):
        print("Aguardando o download ser concluído...")
        time.sleep(2)  

    while not os.path.exists(os.path.join(pasta_download_origem, nome_arquivo)):
        print(f"Verificando se {nome_arquivo} foi baixado...")
        time.sleep(2)

    print(f"Download concluído: {nome_arquivo} está na pasta.")
    print()

def aguardando_download2(pasta_download_origem, nome_arquivo):
    while any([filename.endswith('.crdownload') or filename.endswith('.part') 
               for filename in os.listdir(pasta_download_origem)]):
        print("Aguardando o download ser concluído...")
        time.sleep(2)

    while not any([filename == nome_arquivo or filename.endswith('.odt') 
                   for filename in os.listdir(pasta_download_origem)]):
        print(f"Verificando se {nome_arquivo} foi baixado...")
        time.sleep(2)

    print(f"Download concluído: {nome_arquivo} está na pasta.")
    print()

def coleta_de_dados():
    url = 'https://arquivos.receitafederal.gov.br/dados/cnpj/dados_abertos_cnpj/?C=N;O=D'
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        
        linha = soup.find_all('tr') 

        for row in linha:
            colunas = row.find_all('td')
            if len(colunas) > 1: 
                ano_mes = colunas[1].find('a').text.strip() if colunas[1].find('a') else "N/A"
                data_modificacao = colunas[2].text.strip() if len(colunas) > 2 else "N/A"
                nome_arquivo1 = 'Temp'

                if "temp/" in ano_mes:
                    print(f"Ultima Atualização Dados Abertos: {data_modificacao}")
                    print()

                #print(f"{ano_mes} {data_modificacao}")
    else:
        print("Não foi possível abrir a pág")

def extracao_dados(navegador):
    
    indice_inicial = 4
    try:
        while True:
            try:
                start_time = time.time() 
                pasta_download_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS" #ALTERAR CAMINHO
                xpath_arquivo = f"/html/body/table/tbody/tr[{indice_inicial}]/td[2]/a"
                
                arquivo = navegador.find_element(By.XPATH, xpath_arquivo)

                nome_arquivo = arquivo.text
                print(nome_arquivo)

                arquivo.click()
                
                aguardando_download(pasta_download_origem=pasta_download_origem, nome_arquivo=nome_arquivo)
                indice_inicial += 1

            except NoSuchElementException as e:
                print(f"Arquivos Finalizados")
                break

    except Exception as e:
        print(f"Download finalizado ou erro encontrado: {e}")


def checagem_dados():

    #ABRINDO O CHROME 
    abrir_chrome()
    time.sleep(2)

    navegador = iniciar_navegador(com_debugging_remoto=True) #INICIANDO SELENIUM 

    time.sleep(5)

    btn_recursos = WebDriverWait(navegador, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div/section/div/div[3]/div[2]/div[3]/div[2]/header/button"))
    )

    btn_recursos.click()
    time.sleep(0.5)

    btn_dados_abertos = navegador.find_element(By.XPATH, '/html/body/div/section/div/div[3]/div[2]/div[3]/div[2]/div/div[4]/div[2]/div[2]/div/button')
    btn_dados_abertos.click()
    time.sleep(1)

    coleta_de_dados() #VERIFICAÇÃO DE DATA DAS ÚLTIMAS ATUALIZAÇÕES DE ARQUIVOS
    
    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])

    btn_ultimo_mes = navegador.find_element(By.XPATH, '/html/body/table/tbody/tr[5]/td[2]/a') #ENTRANDO NO ÚLTIMO MÊS LANÇADO
    btn_ultimo_mes.click()
    extracao_dados(navegador) #DOWNLOAD DE TODOS OS ARQUIVOS 

    time.sleep(1)
    pyautogui.hotkey('ctrl', 'w')
    navegador.switch_to.window(abas[0])

    btn_regime_tributario = navegador.find_element(By.XPATH, '/html/body/div/section/div/div[3]/div[2]/div[3]/div[2]/div/div[5]/div[2]/div[2]/div/button[1]')
    btn_regime_tributario.click()

    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])

    time.sleep(1)
    coleta_de_dados_regime_trib(navegador)
    time.sleep(1)

    extracao_dados_regime_trib(navegador)
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'w')
    navegador.switch_to.window(abas[0])
    time.sleep(1)

    pagina_portal_devedor()
    time.sleep(3)
    
    data = coleta_de_dados_portal_devedor()
    time.sleep(1)

    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])
    time.sleep(2)

    download_lista_devedores(navegador)


#-----------------------------------------------------REGIME TRIBUTÁRIO---------------------------------------------------------

def coleta_de_dados_regime_trib(navegador):
    indice_inicial = 4
    indice_inicial_arquivo = 4

    while True:
        try:
            inf = navegador.find_element(By.XPATH, f'/html/body/table/tbody/tr[{indice_inicial}]/td[3]')
            data = inf.text

            arquivo = navegador.find_element(By.XPATH, f'/html/body/table/tbody/tr[{indice_inicial_arquivo}]/td[2]/a')
            nome_arquivo = arquivo.text
            
            print(f'Arquivo {nome_arquivo}: {data}')

            indice_inicial +=1
            indice_inicial_arquivo +=1
        except NoSuchElementException:
            print()
            break

def extracao_dados_regime_trib(navegador):

    indice_inicial = 4
    try:
        total_arquivos = len(navegador.find_elements(By.XPATH, "/html/body/table/tbody/tr")) - 4
        with tqdm(total=total_arquivos, desc="Baixando arquivos", unit="arquivo") as barra:

            while True:
                try:
                    start_time = time.time() 
                    pasta_download_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS" #ALTERAR CAMINHO
                    xpath_arquivo = f"/html/body/table/tbody/tr[{indice_inicial}]/td[2]/a"
                    
                    arquivo = navegador.find_element(By.XPATH, xpath_arquivo)

                    nome_arquivo = arquivo.text
                    print(nome_arquivo)

                    arquivo.click()
                    
                    aguardando_download2(pasta_download_origem=pasta_download_origem, nome_arquivo=nome_arquivo)
                    indice_inicial += 1

                    barra.update(1)  

                    tempo_gasto = time.time() - start_time
                    tempo_restante = (len(total_arquivos) - barra.n) * tempo_gasto
                    print(f"Tempo estimado para conclusão: {tempo_restante:.2f} segundos")
                    
                except NoSuchElementException:
                    print(f"Arquivos Finalizados")
                    print()
                    break

    except Exception as e:
        print(f"Download finalizado ou erro encontrado: {e}")

#------------------------------------------------------PORTAL DEVEDOR----------------------------------------------------------

def pagina_portal_devedor():
    url_portal_devedor = 'https://listadevedores.sefa.pr.gov.br/portal-devedor/arquivo'
    pyautogui.hotkey('ctrl', 't')
    pyautogui.write(url_portal_devedor)
    pyautogui.press('enter')

def coleta_de_dados_portal_devedor():
    url = 'https://listadevedores.sefa.pr.gov.br/portal-devedor/arquivo'  
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    informacoes = soup.find('div', class_='alert alert-light mt-3')

    ultima_atualizacao = informacoes.get_text(strip=True) 
    data = re.search(r'\d{2}/\d{4}', ultima_atualizacao)
    data_atualizacao = f'01-{data.group().replace('/', '-')}'

    print(f'Última Atualização Portal Devedores: {data.group()}')
    return data_atualizacao

def download_lista_devedores(navegador):
    arquivo_href = navegador.find_element(By.XPATH, '//*[@id="navegacao-pesquisar"]/div[1]/a')
    botao_download = navegador.find_element(By.XPATH, '//*[@id="navegacao-pesquisar"]/div[1]/a/button')
    botao_download.click()

    pasta_download_origem = r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS" #ALTERAR CAMINHO 
    href = arquivo_href.get_attribute("href")
    nome_arquivo = href.split('/')[-1]
    nome_arquivo2 = nome_arquivo.split('.zip')[0] #SEM EXTENSÃO
    print(nome_arquivo)
    print()
    
    aguardando_download(pasta_download_origem=pasta_download_origem, nome_arquivo=nome_arquivo)


checagem_dados()

#navegador = iniciar_navegador()
