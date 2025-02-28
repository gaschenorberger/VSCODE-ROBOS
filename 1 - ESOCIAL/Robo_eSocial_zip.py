from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import traceback
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import JavascriptException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException 
import os
import calendar
from datetime import datetime


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
"""

        ("01/01/2018", "31/01/2018"),
        ("01/02/2018", "28/02/2018"),
        ("01/03/2018", "31/03/2018"),
        ("01/04/2018", "30/04/2018"),
        ("01/05/2018", "31/05/2018"),
        ("01/06/2018", "30/06/2018"),
        ("01/07/2018", "31/07/2018"),
        ("01/08/2018", "31/08/2018"),
        ("01/09/2018", "30/09/2018"),
        ("01/10/2018", "31/10/2018"),
        ("01/11/2018", "30/11/2018"),
        ("01/12/2018", "31/12/2018"),
        
        ("01/01/2019", "31/01/2019"),
        ("01/02/2019", "28/02/2019"),
        ("01/03/2019", "31/03/2019"),
        ("01/04/2019", "30/04/2019"),
        ("01/05/2019", "31/05/2019"),
        ("01/06/2019", "30/06/2019"),
        ("01/07/2019", "31/07/2019"),
        ("01/08/2019", "31/08/2019"),
        ("01/09/2019", "30/09/2019"),
        ("01/10/2019", "31/10/2019"),
        ("01/11/2019", "30/11/2019"),
        ("01/12/2019", "31/12/2019"),
        
        ("01/01/2020", "31/01/2020"),
        ("01/02/2020", "29/02/2020"),
        ("01/03/2020", "31/03/2020"),
        ("01/04/2020", "30/04/2020"),
        ("01/05/2020", "31/05/2020"),
        ("01/06/2020", "30/06/2020"),
        ("01/07/2020", "31/07/2020"),
        ("01/08/2020", "31/08/2020"),
        ("01/09/2020", "30/09/2020"),
        ("01/10/2020", "31/10/2020"),
        ("01/11/2020", "30/11/2020"),
        ("01/12/2020", "31/12/2020"),
        
        ("01/01/2021", "31/01/2021"),
        ("01/02/2021", "28/02/2021"),
        ("01/03/2021", "31/03/2021"),
        ("01/04/2021", "30/04/2021"),
        ("01/05/2021", "31/05/2021"),
        ("01/06/2021", "30/06/2021"),
        ("01/07/2021", "31/07/2021"),
        ("01/08/2021", "31/08/2021"),
        ("01/09/2021", "30/09/2021"),
        ("01/10/2021", "31/10/2021"),
        ("01/11/2021", "30/11/2021"),
        ("01/12/2021", "31/12/2021"),
        
        ("01/01/2022", "31/01/2022"),
        ("01/02/2022", "28/02/2022"),
        ("01/03/2022", "31/03/2022"),
        ("01/04/2022", "30/04/2022"),
        ("01/05/2022", "31/05/2022"),
        ("01/06/2022", "30/06/2022"),
        ("01/07/2022", "31/07/2022"),
        ("01/08/2022", "31/08/2022"),
        ("01/09/2022", "30/09/2022"),
        ("01/10/2022", "31/10/2022"),
        ("01/11/2022", "30/11/2022"),
        ("01/12/2022", "31/12/2022"),
        
        ("01/01/2023", "31/01/2023"),
        ("01/02/2023", "28/02/2023"),
        ("01/03/2023", "31/03/2023"),
        ("01/04/2023", "30/04/2023"),
        ("01/05/2023", "31/05/2023"),
        ("01/06/2023", "30/06/2023"),
        ("01/07/2023", "31/07/2023"),
        ("01/08/2023", "31/08/2023"),
        ("01/09/2023", "30/09/2023"),
        ("01/10/2023", "31/10/2023"),
        ("01/11/2023", "30/11/2023"),
        ("01/12/2023", "31/12/2023"),
        
        ("01/01/2024", "31/01/2024"),
        ("01/02/2024", "29/02/2024"),
        ("01/03/2024", "31/03/2024"),
        ("01/04/2024", "30/04/2024"),
        ("01/05/2024", "31/05/2024"),
        ("01/06/2024", "30/06/2024"),
        ("01/07/2024", "31/07/2024"),
        ("01/08/2024", "31/08/2024"),
        ("01/09/2024", "30/09/2024"),
        ("01/10/2024", "31/10/2024"),
        ("01/11/2024", "30/11/2024"),
        ("01/12/2024", "31/12/2024"),
"""


def gerar_lista_meses(ano_inicio, ano_fim, mes_inicio=1):
    lista_meses = []
    for ano in range(ano_inicio, ano_fim + 1):
        for mes in range(mes_inicio, 13):
            inicio_mes = f"01/{mes:02d}/{ano}"
            ultimo_dia_mes = calendar.monthrange(ano, mes)[1]
            fim_mes = f"{ultimo_dia_mes:02d}/{mes:02d}/{ano}"
            lista_meses.append((inicio_mes, fim_mes))
        mes_inicio = 1 
    return lista_meses

def listaMeses():
    lista_meses = [   
           
        ("01/01/2020", "31/01/2020"),
        ("01/02/2020", "29/02/2020"),
        ("01/03/2020", "31/03/2020"),
        ("01/04/2020", "30/04/2020"),
        ("01/05/2020", "31/05/2020"),
        ("01/06/2020", "30/06/2020"),
        ("01/07/2020", "31/07/2020"),
        ("01/08/2020", "31/08/2020")
    ]
    return lista_meses

def paginaSolicitacao():
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
        #                                                    /html/body/div[3]/div[3]/div/div/ul/li[4]
        dropDownDownload = navegador.find_element(By.XPATH, '/html/body/div[3]/div[3]/div/div/ul/li[4]')
        navegador.execute_script('arguments[0].setAttribute("class","dropdown open");', dropDownDownload)
        #                                                                     ⇓
        #                                               /html/body/div[3]/div[3]/div/div/ul/li[4]/ul/li[2]/a
        solicitacao = navegador.find_element(By.XPATH, '/html/body/div[3]/div[3]/div/div/ul/li[4]/ul/li[2]/a').click()
    except (TimeoutException, NoSuchElementException, Exception):
        print("erro ao ir para a pagina de Solicitação:")
        traceback.print_exc()

def filtroSolicitacao():
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        filtroTipoPedido = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, f'//*[@id="TipoPedido"]')))

        #filtroTipoPedido = navegador.find_element(By.XPATH, '//*[@id="TipoPedido"]')

        Ofiltro = Select(filtroTipoPedido)
        Ofiltro.select_by_index(1)
        filtroSelecionado = Ofiltro.first_selected_option
        filtroAplicado = filtroSelecionado.text

    except (TimeoutException, NoSuchElementException, Exception):
        print("erro no filtro!")
        traceback.print_exc()

def pesquisarArquivos(primeiro_dia, ultimo_dia):
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        campoDataInicio = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, f'//*[@id="DataInicial"]')))


        #campoDataInicio = navegador.find_element(By.XPATH, '//*[@id="DataInicial"]')
        campoDataInicio.clear()
        campoDataInicio.send_keys(primeiro_dia)


        campoDataFim = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, f'//*[@id="DataFinal"]')))

        #campoDataFim = navegador.find_element(By.XPATH, '//*[@id="DataFinal"]')
        campoDataFim.clear()
        campoDataFim.send_keys(ultimo_dia)


        salvar = navegador.find_element(By.XPATH, '//*[@id="btnSalvar"]')
        salvar.click()

        print("Pesquisa feita")

    except (TimeoutException, NoSuchElementException, Exception):
        print("erro ao pesquisar!")
        traceback.print_exc()


print()

meses = listaMeses()
for primeiro_dia, ultimo_dia in meses:
    print(f"Primeiro dia: {primeiro_dia}, Último dia: {ultimo_dia}")
    paginaSolicitacao()
    filtroSolicitacao()
    pesquisarArquivos(primeiro_dia, ultimo_dia)
