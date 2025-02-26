import openpyxl.workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.common.exceptions import NoSuchElementException 
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import JavascriptException
import pyperclip
import openpyxl
from openpyxl.styles import PatternFill
import traceback
import os
from pathlib import Path
from datetime import datetime
from openpyxl.styles import PatternFill
import pandas as pd

# ----------------------------------------------------------------------------------------------

def iniciar_navegador(com_debugging_remoto=True):
    #chrome_driver_path = ChromeDriverManager().install()
    chrome_driver_path = r"C:\Users\mateus.silva\.wdm\drivers\chromedriver\win64\131.0.6778.204\chromedriver-win32\chromedriver.exe"
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

navegador = iniciar_navegador()

def calcular_tempoExecucao(func, *args, **kwargs):
    tempoInicio = datetime.now()
    resultado = func(*args, **kwargs)
    tempoFim = datetime.now()
    tempoExecucao = str(tempoFim - tempoInicio)

    larguraTotal = max(len(tempoExecucao) + 8, 30)  # Garante espaço suficiente
    titulo = " TEMPO DE EXECUÇÃO "
    
    linhaSuperior = f"╭{'─' * (larguraTotal - 2)}╮"
    linhaTitulo = f"│{titulo.center(larguraTotal - 2)}│"
    linhaMeio = f"│{tempoExecucao.center(larguraTotal - 2)}│"
    linhaInferior = f"╰{'─' * (larguraTotal - 2)}╯"
    
    print(linhaSuperior)
    print(linhaTitulo)
    print(linhaMeio)
    print(linhaInferior)
    
    return resultado


def adicionar_dados(dados):
    data_frame = pd.DataFrame(dados, columns=['Nº_INSCRICAO', 'CNPJ', 'NOME', 'DT_INSCRICAO', 'RECEITA', 'Nº_PROCESSO', 'SITUACAO', 'NATUREZA_DIVIDA', 'VL_PRINCIPAL', 'VL_MULTA', 'VL_JUROS', 'VL_ENCARGO', 'VL_TOTAL'])
    data_frame.to_excel('0 - ECAC\\PGFN\\Relatório - PGFN.xlsx', index=False)


# PGFN
def verificar_url_PGFN(navegador):
    abas = navegador.window_handles

    for aba in abas:
        navegador.switch_to.window(aba)
        url = navegador.current_url

        if 'regularize.pgfn.gov.br/consultaDividas' in url:
            print("URL PGFN encontrada")
            #navegador.refresh()
            WebDriverWait(navegador, 10).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
            return True
        elif 'regularize.pgfn.gov.br' in url or 'regularize.pgfn.gov.br/consultaDividas/sida/detalhar' in url:
            print("URL PGFN, quase correta encontrada, ajustando URL")
            navegador.get('https://www.regularize.pgfn.gov.br/consultaDividas')
            return True  

    print("Nenhuma aba com a URL correta foi encontrada")
    return False

def percorrer_inscricoes():
    verificar_url_PGFN(navegador)
    a, g, i = 1, 1, 1
    dados = []
    while True:
        filtro(navegador, a, g)
    
        try:
            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

            try:
                aba = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/ul/li[{a}]/a/span[2]')))
            except TimeoutException:
                print("Não há mais abas")
                break

            try:
                grupo = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/div/tab[{a}]/app-consulta-divida-aba/app-consulta-divida-grupo[{g}]/div[1]/div/div[1]')))
            except TimeoutException:
                print("Não há mais grupos")
                a += 1
                g = 1
                i = 1
                continue

            if g == 4:
                g = 1
                i = 1
                a += 1
                navegador.execute_script("arguments[0].click();", aba)
                continue

            navegador.execute_script("arguments[0].click();", aba)
            abaIncricao = aba.text
            if '(0)' in abaIncricao:
                g = 1
                i = 1
                a += 1
                navegador.execute_script("arguments[0].click();", aba)
                continue

            tipoIncricao = grupo.text
            if 'Extinta' in tipoIncricao:
                g = 1
                i = 1
                a += 1
                navegador.execute_script("arguments[0].click();", aba)
                continue


            filtro(navegador, a, g)

            inscricoes = navegador.find_elements(By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/div/tab[{a}]/app-consulta-divida-aba/app-consulta-divida-grupo[{g}]/div[2]/div[3]/table/tbody/tr[{i}]/td[2]/a')
            
            if not inscricoes:
                g += 1
                i = 1
                continue


            inscricao = inscricoes[0]
            numInscricao = inscricao.text

            print("===========================================================")
            print(f'{a}º | Aba       -> {abaIncricao}')
            print(f'{g}º | Grupo     -> {tipoIncricao}')
            print(f'{i}º | Inscricao -> {numInscricao}')

            
            extrair_informacao(inscricao, dados)
            
            i += 1
        except TimeoutException:
            traceback.print_exc()
            return
        except Exception:
            print("erro ao percorrer a lista!")
            traceback.print_exc()
            break

def filtro(navegador, a, g):
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        filtroQntPaginas = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/div/tab[{a}]/app-consulta-divida-aba/app-consulta-divida-grupo[{g}]/div[2]/div[3]/div[1]/div[2]/select')))

        filtrarTodasPaginas = navegador.find_element(By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/div/tab[{a}]/app-consulta-divida-aba/app-consulta-divida-grupo[{g}]/div[2]/div[3]/div[1]/div[2]/select/option[6]')

        Ofiltro = Select(filtroQntPaginas)
        filtroSelecionado = Ofiltro.first_selected_option
        filtroAplicado = filtroSelecionado.text

        if filtroAplicado != 'Todas':
            filtrarTodasPaginas.click()

        else:
            pass
    except (TimeoutException, JavascriptException):
        pass
    except Exception as e:
        print("erro no filtro!")
        traceback.print_exc()

def extrair_informacao(inscricao, dados):
    tentativa = 0
    
    while tentativa < 3:
        try:
            WebDriverWait(navegador, 60).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
            navegador.execute_script("arguments[0].click();", inscricao)

            xpaths = {
                '/html/body/app-root/div/div[2]/app-inscricao/main/div[1]/h1': 'Padrao',
                '/html/body/app-root/div/div[2]/app-debcad/main/div[1]/h1': 'Debcad',
                '/html/body/app-root/div/div[2]/app-fgts/main/div[1]/h1': 'Fgts'
            }

            tipoInscricao = None

            for xpath, tipo in xpaths.items():
                try:
                    tipoPagina = WebDriverWait(navegador, 5).until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    )
                    tipoInscricao = tipo
                    break
                except TimeoutException:
                    continue

            match tipoInscricao:
                case 'Padrao':
                    relatDetalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/ul/li[2]/a')))
                    relatDetalElement.click()

                    filtTodosElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/fieldset/form/div/div[1]/div[1]/input')))                    
                    filtTodosElement.click()

                    gerarRelatElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/fieldset/div/button[1]')))                    
                    gerarRelatElement.click()

                    time.sleep(0.5)
                    

                    nomeEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/div[1]/span')))
                    cnpjEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/div[2]/span')))
                    numSitInscriElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/div[1]/div[2]')))
                    dataInscricaoElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[1]')))
                    natDividaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[3]')))
                    receitaDividaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[4]')))
                    numProceAdminElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[8]')))
                    valPrincipalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[1]')))
                    valMultaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[2]')))
                    valJurosElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[3]')))
                    valEncargoElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[4]')))
                    valTotalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[5]')))
                    voltarElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/div[3]/button[1]')))



                    nomeEmpresa = nomeEmpresaElement.text.strip()
                    cnpjEmpresa = cnpjEmpresaElement.text.strip()
                    numSitInscri = numSitInscriElement.text.replace("\n",":").strip()
                    numSitInscri = numSitInscri.split(sep=":")
                    numInscricao = numSitInscri[1].strip()
                    SitInscricao = numSitInscri[3].strip()
                    dataInscricao = dataInscricaoElement.text.strip()
                    natDivida = natDividaElement.text.strip()
                    receitaDivida = receitaDividaElement.text.strip()
                    numProceAdmin = numProceAdminElement.text.strip()

                    valPrincipal = valPrincipalElement.text.strip()
                    valMulta = valMultaElement.text.strip()
                    valJuros = valJurosElement.text.strip()
                    valEncargo = valEncargoElement.text.strip()
                    valTotal = valTotalElement.text.strip()



                    #print(f'{numInscricao} | {cnpjEmpresa} | {nomeEmpresa} | {dataInscricao} | {receitaDivida} | {numProceAdmin} | {SitInscricao} | {natDivida}')
                    #print(f'{valPrincipal} | {valMulta} | {valJuros} | {valEncargo} | {valTotal}')                    


                    dados.append([numInscricao,cnpjEmpresa,nomeEmpresa,dataInscricao,receitaDivida,numProceAdmin,SitInscricao,natDivida,valPrincipal,valMulta,valJuros,valEncargo,valTotal])
                    adicionar_dados(dados)
                    voltarElement.click()
                    break
                case 'Debcad':
                    relatDetalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/ul/li[2]/a')))
                    relatDetalElement.click()

                    filtTodosElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset/form/div/div[1]/div[1]/input')))                    
                    filtTodosElement.click()

                    gerarRelatElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset/div/button[1]')))                    
                    gerarRelatElement.click()

                    time.sleep(0.5)
                    

                    nomeEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[1]')))
                    cnpjEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[2]')))
                    numSitInscriElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/div/div/div[2]')))
                    dataInscricaoElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[11]')))
                    natDividaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[8]')))
                    receitaDividaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[13]')))
                    valPrincipalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[1]')))
                    valMultaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[2]')))
                    valJurosElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[3]')))
                    valEncargoElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[4]')))
                    valTotalElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[5]')))
                    voltarElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/div[4]/button[1]')))


                    nomeeEmpresa = nomeEmpresaElement.text.split(sep=":")
                    nomeEmpresa = nomeeEmpresa[1].strip()
                    cnpjjEmpresa = cnpjEmpresaElement.text.split(sep=":")
                    cnpjEmpresa = cnpjjEmpresa[1].strip()
                    numSitInscri = numSitInscriElement.text.replace("\n",":")
                    numSitInscri = numSitInscri.split(sep=":")
                    numInscricao = numSitInscri[1].strip()
                    SitInscricao = numSitInscri[3].strip()
                    dataaInscricao = dataInscricaoElement.text.split(sep=":")
                    dataInscricao = dataaInscricao[1].strip()

                    nattDivida = natDividaElement.text.split(sep=":")
                    natDivida = nattDivida[1].strip()
                    receitaaDivida = receitaDividaElement.text.split(sep=":")
                    receitaDivida = receitaaDivida[1].strip()

                    valPrincipal = valPrincipalElement.text.strip()
                    valMulta = valMultaElement.text.strip()
                    valJuros = valJurosElement.text.strip()
                    valEncargo = valEncargoElement.text.strip()
                    valTotal = valTotalElement.text.strip()



                    #print(f'{numInscricao} | {cnpjEmpresa} | {nomeEmpresa} | {dataInscricao} | {receitaDivida} | {SitInscricao} | {natDivida}')
                    #print(f'{valPrincipal} | {valMulta} | {valJuros} | {valEncargo} | {valTotal}')                    



                    dados.append([numInscricao,cnpjEmpresa,nomeEmpresa,dataInscricao,receitaDivida,'',SitInscricao,natDivida,valPrincipal,valMulta,valJuros,valEncargo,valTotal])
                    adicionar_dados(dados)
                    voltarElement.click()
                    break
                case 'Fgts':
                    
                    nomeEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[1]')))
                    cnpjEmpresaElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[2]')))
                    numSitInscriElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/div[4]/div[2]')))
                    dataInscricaoElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[6]')))
                    valTotalConsElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[2]/div[2]')))
                    voltarElement = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/div[5]/button[1]')))                   
                    
                    nomeeEmpresa = nomeEmpresaElement.text.split(sep=":")
                    nomeEmpresa = nomeeEmpresa[1].strip()
                    cnpjjEmpresa = cnpjEmpresaElement.text.split(sep=":")
                    cnpjEmpresa = cnpjjEmpresa[1].strip()
                    numSitInscri = numSitInscriElement.text.replace("\n",":")
                    numSitInscri = numSitInscri.split(sep=":")
                    numInscricao = numSitInscri[1].strip()
                    SitInscricao = numSitInscri[3].strip()
                    dataaInscricao = dataInscricaoElement.text.split(sep=":")
                    dataInscricao = dataaInscricao[1].strip()
                    vallTotalCons = valTotalConsElement.text.replace("\n",":")
                    vallTotalCons = vallTotalCons.split(sep=":")
                    valTotalCons = vallTotalCons[1].strip()


                    #print(f'{numInscricao} | {cnpjEmpresa} | {nomeEmpresa} | {dataInscricao} | {valTotalCons}')  

                    dados.append([numInscricao,cnpjEmpresa,nomeEmpresa,dataInscricao,'','',SitInscricao,'','','','','',valTotalCons])
                    adicionar_dados(dados)
                    voltarElement.click()
                    break
                case _:
                    print("Nenhum valor correspondente foi encontrado")
                    tentativa += 1
                    continue
            
        except Exception:
            traceback.print_exc()
            tentativa += 1
            continue




# ----------------------------------------------------------------------------------------------
print("iniciando...")
time.sleep(2)
calcular_tempoExecucao(percorrer_inscricoes)
