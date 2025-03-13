from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
import time
from selenium.common.exceptions import NoSuchElementException 
from selenium.common.exceptions import TimeoutException
import os
import traceback
from datetime import datetime
import shutil
import pandas as pd

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

#----------------------------------------------------------------------------------------------------------------
def aplicarFiltros(anoAp, anoTr, filtroAlternativo=False):
    DtIniApuracao = f"01/01/{anoAp}"
    DtFimApuracao = f"31/12/{anoAp}"

    DtIniTransmissao = f"01/01/{anoTr}"

    anoAtual = datetime.now().year
    if anoTr == anoAtual:
        hoje = datetime.now().strftime("%d/%m/%Y")
        DtFimTransmissao = hoje
    else:
        DtFimTransmissao = f"26/12/{anoTr}"

    if filtroAlternativo:
        DtIniTransmissao = f"27/12/{anoTr}"
        DtFimTransmissao = f"31/12/{anoTr}"

    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        navegador.switch_to.default_content()
        iframe = navegador.find_element(By.ID, 'frmApp')
        navegador.switch_to.frame(iframe)

        #Apuracao
        campoDataInicioAp = navegador.find_element(By.XPATH, '//*[@id="txtDataInicio"]')
        campoDataInicioAp.clear()
        campoDataInicioAp.send_keys(DtIniApuracao)

        campoDataFimAp = navegador.find_element(By.XPATH, '//*[@id="txtDataFinal"]')
        campoDataFimAp.clear()
        campoDataFimAp.send_keys(DtFimApuracao)

        #Transmisssao
        campoDataInicioTr = navegador.find_element(By.XPATH, '//*[@id="txtDataTransmissaoInicial"]')
        campoDataInicioTr.clear()
        campoDataInicioTr.send_keys(DtIniTransmissao)

        campoDataFimTr = navegador.find_element(By.XPATH, '//*[@id="txtDataTransmissaoFinal"]')
        campoDataFimTr.clear()
        campoDataFimTr.send_keys(DtFimTransmissao)


        #categoria declaração
        categoriasFiltradas = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div/button/span[1]')
        txtCategoriasFiltrdas = categoriasFiltradas.text

        if "selecionados" in txtCategoriasFiltrdas:
            filtroTipoDeclaracao = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div')
            navegador.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span11 open';", filtroTipoDeclaracao)
            selecionarTodos = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[3]/div/div/div/div/div/div/button[1]')
            time.sleep(0.5)
            navegador.execute_script("arguments[0].click();", selecionarTodos)

        #situação declaração
        situacoesFiltradas = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div/button/span[1]')
        txtSituacoesFiltradas = situacoesFiltradas.text

        if 'Em andamento' in txtSituacoesFiltradas:
            filtroSituacaoDeclaracao = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div')
            navegador.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span10 open';", filtroSituacaoDeclaracao)
            filtroEmAndamento = navegador.find_element(By.XPATH, '//*[@id="conteudo-pagina"]/div[1]/div[4]/div/div/div/div/ul/li[1]/a')
            time.sleep(0.5)
            navegador.execute_script("arguments[0].click();", filtroEmAndamento)

        botaoPesquisar = navegador.find_element(By.XPATH, '//*[@id="ctl00_cphConteudo_btnFiltar"]')
        time.sleep(0.5)
        navegador.execute_script("arguments[0].click();", botaoPesquisar)

    except (TimeoutException, NoSuchElementException):
        print("erro ao aplicaro o filtro:")
        traceback.print_exc()

def baixarXML():
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        navegador.switch_to.default_content()

        iframe = navegador.find_element(By.ID, 'frmApp')
        navegador.switch_to.frame(iframe)

        time.sleep(0.5)
        dropDownRelatorios = WebDriverWait(navegador, 15).until(EC.visibility_of_element_located(
                    (By.XPATH, '//*[@id="dropDown_Relatorios"]')))
        time.sleep(0.5)
        #dropDownRelatorios = navegador.find_element(By.XPATH, '//*[@id="dropDown_Relatorios"]')
        #driver.execute_script("arguments[0].className = 'btn-group bootstrap-select show-tick span11 open';", filtroTipoDeclaracao)

        navegador.execute_script("arguments[0].className = 'class=dropdown open' ", dropDownRelatorios)

        time.sleep(1)
        baixarXMLdeSaida = WebDriverWait(navegador, 15).until(EC.visibility_of_element_located(
                    (By.XPATH, '//*[@id="dropDown_Relatorios_Menu"]/li[3]/a')))
        time.sleep(0.5)
        #baixarXMLdeSaida = navegador.find_element(By.XPATH, '//*[@id="dropDown_Relatorios_Menu"]/li[3]/a')
        baixarXMLdeSaida.click()
        
    except (TimeoutException, Exception):
        print("erro ao baixar xml!")
        traceback.print_exc()

def percorrer_dctf():
    anoAp, anoTr = 2018, 2018
    d = 2
    filtroAlternativo = False

    aplicarFiltros(anoAp, anoTr, filtroAlternativo)


    while True:
        try:
            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

            navegador.switch_to.default_content()
            iframe = navegador.find_element(By.ID, 'frmApp') 
            navegador.switch_to.frame(iframe)
            
            # Percorrer lista de DCTF
            Dctf = True
            Visualizar = True
            try:
                campoPeriodoApu = navegador.find_element(By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[1]')
                periodoApuracao = campoPeriodoApu.text
                periodoApuracao = periodoApuracao.replace("/","_")


                campoTipo = navegador.find_element(By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[5]')
                tipo = campoTipo.text

                campoSituacao = navegador.find_element(By.XPATH, f'//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs"]/tbody/tr[{d}]/td[6]')
                situacao = campoSituacao.text
                
                try:
                    botaoVisualizar = navegador.find_element(By.XPATH, f'/html/body/form/div[5]/div/div/div[6]/div/div/div[1]/table/tbody/tr[{d}]/td[9]/a[1]')
       
                except NoSuchElementException:
                    Visualizar = False

            except NoSuchElementException:
                Dctf = False
                pass

            # lógica
            if anoTr == 2026:
                break

            if not Dctf:
                print()
                d = 2

                if not filtroAlternativo:
                    filtroAlternativo = True
                    aplicarFiltros(anoAp, anoTr, filtroAlternativo)
                    continue
                else:
                    filtroAlternativo = False
                    anoTr += 1

                    if anoTr == 2026:
                        anoAp += 1
                        anoTr = anoAp

                    aplicarFiltros(anoAp, anoTr, filtroAlternativo)
                    continue

            elif Visualizar:
                print(f'{periodoApuracao} - {tipo} - {situacao}')
                time.sleep(0.5)
                botaoVisualizar.click()
                d += 1
                baixarXML()
                navegador.back()
                continue

            else:
                print(f"sem opção vizualizar na dctf da linha {d-1}º pulando pra proxima")
                d += 1
                continue

        except (TimeoutException, Exception):
            print("erro ao percorrer a lista!")
            traceback.print_exc()
            break

def redirecionarXml(origem=r"C:\Users\gabriel.alvise\Desktop\DOWNLOAD'S ARQUIVOS",destino=r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - WEB\XMLs'):

    for arquivo in os.listdir(origem):
        if "XMLSaida" in arquivo and arquivo.endswith(".xml"):
            caminhoOrigem = os.path.join(origem, arquivo)
            caminhoDestino = os.path.join(destino, arquivo)
            
            try:
                shutil.move(caminhoOrigem, caminhoDestino)
                print(f"Arquivo {arquivo} foi movido para {destino}")
            except Exception as e:
                print(f"Erro ao mover o arquivo {arquivo}: {e}") 


print("iniciando...")

time.sleep(1)
percorrer_dctf()
redirecionarXml()

