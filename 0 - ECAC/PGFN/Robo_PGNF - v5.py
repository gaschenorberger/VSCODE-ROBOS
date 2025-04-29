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
import requests
from bs4 import BeautifulSoup

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

navegador = iniciar_navegador()

def calcular_tempoExecucao(func, *args, **kwargs):
    tempoInicio = datetime.now()
    resultado = func(*args, **kwargs)
    tempoFim = datetime.now()
    tempoExecucao = str(tempoFim - tempoInicio)

    larguraTotal = max(len(tempoExecucao) + 8, 30)
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
    colunas = ['Nº_INSCRICAO', 'CNPJ', 'NOME', 'DT_INSCRICAO', 'RECEITA', 'Nº_PROCESSO', 'SITUACAO', 'NATUREZA_DIVIDA',
               'VL_PRINCIPAL', 'VL_MULTA', 'VL_JUROS', 'VL_ENCARGO', 'VL_TOTAL_CONSOLIDADO','MULTA_DE_OFICIO','MULTA_ISOLADA','HONORARIO',
               'DT_APURACAO', 'DT_VENCIMENTO', 'VL_DEBITO', 'VL_REMANESCENTE', 'VL_MULTA_MORA', 'NATUREZA_DEBITO', 'FORMA_CONSTITUICAO',
               'DT_ARRECADACAO_PAGAMENTO', 'DT_RECEPCAO_PAGAMENTO', 'VL_RECOLHIDO_PAGAMENTO', 'BANCO/AGENCIA_PAGAMENTO', 'NUM_CREDITO_PAGAMENTO', 'TIPO_CREDITO_PAGAMENTO', 'NUM_PAGAMENTO', 'NUM_SENDA_PAGAMENTO',
               'ESTABELECIMENTO', 'COMPETENCIA', 'VL_ITEM', 'VL_DEVEDOR', 'DT_ARRECADACAO', 'DT_APROPRIACAO', 'VL_APROPRIADO', 'REFERENCIA', 'NUM_GUIA', 'TIPO_CRED']

    df = pd.DataFrame(dados, columns=colunas)
    df.to_excel('0 - ECAC\\PGFN\\Relatório - PGFN.xlsx', index=False)


# PGFN
def verificar_url_PGFN(navegador):
    abas = navegador.window_handles

    for aba in abas:
        navegador.switch_to.window(aba)
        url = navegador.current_url

        if 'regularize.pgfn.gov.br/consultaDividas' in url:
            print("URL PGFN encontrada")
            #navegador.refresh()
            WebDriverWait(navegador, 300).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
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
    listaExtraidos = [] 
    while True:
        filtro(navegador, a, g)
        try:
            WebDriverWait(navegador, 240).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

            try:
                aba = WebDriverWait(navegador, 300).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f'/html/body/app-root/div/div[2]/app-consulta-divida/main/div/div/div/tabset/ul/li[{a}]/a/span[2]')))
            except TimeoutException:
                print("Não há mais abas")
                break

            try:
                grupo = WebDriverWait(navegador, 300).until(
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

            time.sleep(0.4)
            navegador.execute_script("arguments[0].click();", inscricao)
            extrair_informacao(dados, listaExtraidos)
            listaExtraidos.append(numInscricao)

            i += 1
        except TimeoutException:
            traceback.print_exc()
            return
        except Exception:
            print("erro ao percorrer a lista!")
            traceback.print_exc()
            break

    lista_inscricoes = verificarInscricoes() or []

    nao_extraidas = [insc for insc in lista_inscricoes if insc not in listaExtraidos]

    if nao_extraidas:
        print("Inscrições que ainda não foram extraídas:")
        for i in nao_extraidas:
            print(i)
    else:
        print("Todas as inscrições foram extraídas")

        #carai, potente

def filtro(navegador, a, g):
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        filtroQntPaginas = WebDriverWait(navegador, 300).until(
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

def extrair_informacao(dados, listaExtraidos):
    WebDriverWait(navegador, 60).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
    


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
            print(tipoInscricao)
            break
        except TimeoutException:
            continue

    match tipoInscricao:
        case 'Padrao':
            relatDetalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/ul/li[2]/a')))
            relatDetalElement.click()

            filtTodosElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/fieldset/form/div/div[1]/div[1]/input')))                    
            filtTodosElement.click()

            gerarRelatElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/fieldset/div/button[1]')))                    
            gerarRelatElement.click()

            time.sleep(0.5)
            
            nomeEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/div[1]/span')))
            cnpjEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/div[2]/span')))
            numSitInscriElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/div[1]/div[2]')))
            dataInscricaoElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[1]')))
            natDividaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[3]')))
            receitaDividaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[4]')))
            numProceAdminElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[1]/div/span[8]')))
            
            valPrincipalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[1]')))
            valMultaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[2]')))
            valJurosElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[3]')))
            valEncargoElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[4]')))
            valTotalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[2]/table/tr[2]/td[5]')))
            
            debitoDetal = False
            try:
                dataApuraVenciElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[1]')))
                valorDebElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[2]')))
                valorRemanesElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[3]')))
                multaMoraElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[4]')))
                naturezaDebElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[5]')))
                formaConstituiElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[1]/td[6]')))
                
                debitoDetal = True
            except TimeoutException:
                pass

            pagDetal = False
            try:
                dataArrecadRecepPagElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[2]/td[1]')))         
                valRecolhidoPagElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[2]/td[2]')))                   
                bancoAgenciaPagElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[2]/td[3]')))                    
                numCredTipoCredPagElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[2]/td[4]')))                    
                numPagamSendaPagElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[2]/td[5]')))     

                pagDetal = True               
            except TimeoutException:
                pass


            voltarElement = WebDriverWait(navegador, 300).until(
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


            dados.append({
                    'Nº_INSCRICAO': numInscricao,
                    'CNPJ': cnpjEmpresa,
                    'NOME': nomeEmpresa,
                    'DT_INSCRICAO': dataInscricao,
                    'RECEITA': receitaDivida,
                    'Nº_PROCESSO': numProceAdmin,
                    'SITUACAO': SitInscricao,
                    'NATUREZA_DIVIDA': natDivida,

                    'VL_PRINCIPAL': valPrincipal,
                    'VL_MULTA': valMulta,
                    'VL_JUROS': valJuros,
                    'VL_ENCARGO': valEncargo,
                    'VL_TOTAL_CONSOLIDADO': valTotal
            })
            
            """
            print(f'{numInscricao} | {cnpjEmpresa} | {nomeEmpresa} | {dataInscricao} | {receitaDivida} | {numProceAdmin} | {SitInscricao} | {natDivida}')
            print(f'{valPrincipal} | {valMulta} | {valJuros} | {valEncargo} | {valTotal}')                    
            print(f'{dataApuracao} | {dataVencimen} | {valorDeb} | {valorRemanes} | {multaMora} | {naturezaDeb} | {formaConstitui}')
            print(f'{dataArrecadPag} | {dataRecepPag} | {valRecolhidoPag} | {bancoAgenciaPag} | {NumCredPag} | {tipoCredPag} | {numPagamentoPag} | {numSendaPag}')
            """

            if debitoDetal == True:
                y = 1
                while True:
                    try:
                        dataApuraVenciElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[1]')))
                        valorDebElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[2]')))
                        valorRemanesElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[3]')))
                        multaMoraElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[4]')))
                        naturezaDebElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[5]')))
                        formaConstituiElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[3]/table/tbody/tr[{y}]/td[6]')))

                        dataApuraVenci = dataApuraVenciElement.text.replace("\n",":").strip()
                        dataApuraVenci = dataApuraVenci.split(sep=":")
                        dataApuracao = dataApuraVenci[0].strip()
                        dataVencimen = dataApuraVenci[1].strip()
                        valorDeb = valorDebElement.text.strip()
                        valorRemanes = valorRemanesElement.text.strip()
                        multaMora = multaMoraElement.text.strip()
                        naturezaDeb = naturezaDebElement.text.strip()
                        formaConstitui = formaConstituiElement.text.strip()
                        
                        #print(f'{dataApuracao} | {dataVencimen} | {valorDeb} | {valorRemanes} | {multaMora} | {naturezaDeb} | {formaConstitui}')

                        dados.append({
                            'Nº_INSCRICAO': numInscricao,
                            'CNPJ': cnpjEmpresa,
                            'NOME': nomeEmpresa,
                            'DT_INSCRICAO': dataInscricao,
                            'RECEITA': receitaDivida,
                            'Nº_PROCESSO': numProceAdmin,
                            'SITUACAO': SitInscricao,
                            'NATUREZA_DIVIDA': natDivida,

                            'DT_APURACAO': dataApuracao,
                            'DT_VENCIMENTO': dataVencimen,
                            'VL_DEBITO': valorDeb,
                            'VL_REMANESCENTE': valorRemanes,
                            'VL_MULTA_MORA': multaMora,
                            'NATUREZA_DEBITO': naturezaDeb,
                            'FORMA_CONSTITUICAO': formaConstitui
                        })


                        y += 2
                    except TimeoutException:
                        break
            
            if pagDetal == True:
                x = 2
                while True:
                    try:
                        dataArrecadRecepPagElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[{x}]/td[1]')))      
                            
                        valRecolhidoPagElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[{x}]/td[2]')))                   
                        bancoAgenciaPagElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[{x}]/td[3]')))                    
                        numCredTipoCredPagElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[{x}]/td[4]')))                    
                        numPagamSendaPagElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-inscricao/main/tabset/div/tab[2]/div/fieldset[4]/table/tr[{x}]/td[5]')))

                        dataArrecadRecepPag = dataArrecadRecepPagElement.text.replace("\n",":").strip()
                        dataArrecadRecepPag = dataArrecadRecepPag.split(sep=":")
                        dataArrecadPag = dataArrecadRecepPag[0].strip()
                        dataRecepPag = dataArrecadRecepPag[1].strip()
                        valRecolhidoPag = valRecolhidoPagElement.text.strip()
                        bancoAgenciaPag = bancoAgenciaPagElement.text.strip()
                        NumCredTipo = numCredTipoCredPagElement.text.replace("\n",":").strip()
                        NumCredTipo = NumCredTipo.split(sep=":")
                        NumCredPag = NumCredTipo[0].strip()
                        tipoCredPag = NumCredTipo[1].strip()
                        numPagamSenda = numPagamSendaPagElement.text.replace("\n",":").strip()
                        numPagamSenda = numPagamSenda.split(sep=":")
                        numPagamentoPag = numPagamSenda[0].strip()
                        numSendaPag = numPagamSenda[1].strip()

                        #print(f'{dataArrecadPag} | {dataRecepPag} | {valRecolhidoPag} | {bancoAgenciaPag} | {NumCredPag} | {tipoCredPag} | {numPagamentoPag} | {numSendaPag}')

                        dados.append({
                            'Nº_INSCRICAO': numInscricao,
                            'CNPJ': cnpjEmpresa,
                            'NOME': nomeEmpresa,
                            'DT_INSCRICAO': dataInscricao,
                            'RECEITA': receitaDivida,
                            'Nº_PROCESSO': numProceAdmin,
                            'SITUACAO': SitInscricao,
                            'NATUREZA_DIVIDA': natDivida,

                            'DT_APURACAO': dataApuracao,
                            'DT_VENCIMENTO': dataVencimen,
                            'VL_DEBITO': valorDeb,
                            'VL_REMANESCENTE': valorRemanes,
                            'VL_MULTA_MORA': multaMora,
                            'NATUREZA_DEBITO': naturezaDeb,
                            'FORMA_CONSTITUICAO': formaConstitui,
                            'DT_ARRECADACAO_PAGAMENTO':dataArrecadPag,
                            'DT_RECEPCAO_PAGAMENTO':dataRecepPag,
                            'VL_RECOLHIDO_PAGAMENTO':valRecolhidoPag,
                            'BANCO/AGENCIA_PAGAMENTO':bancoAgenciaPag,
                            'NUM_CREDITO_PAGAMENTO':NumCredPag,
                            'TIPO_CREDITO_PAGAMENTO':tipoCredPag,
                            'NUM_PAGAMENTO':numPagamentoPag,
                            'NUM_SENDA_PAGAMENTO':numSendaPag

                        })
                                                
                        x += 1
                    except TimeoutException:
                        break

            adicionar_dados(dados)
            listaExtraidos.append(numInscricao)
            voltarElement.click()
            return numInscricao

        case 'Debcad':
            relatDetalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/ul/li[2]/a')))
            relatDetalElement.click()

            filtTodosElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset/form/div/div[1]/div[1]/input')))                    
            filtTodosElement.click()

            gerarRelatElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset/div/button[1]')))                    
            gerarRelatElement.click()

            time.sleep(0.5)
            
            nomeEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[1]')))
            cnpjEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[2]')))
            numSitInscriElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/div/div/div[2]')))
            dataInscricaoElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[11]')))
            natDividaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[8]')))
            receitaDividaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[2]/div/div[13]')))
            
            valPrincipalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[1]')))
            valMultaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[2]')))
            valMultaOficioElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[3]')))
            valMultaIsoladaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[4]')))
            valJurosElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[5]')))
            valEncargoElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[6]')))
            valHonorarioElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[7]')))
            valTotalElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[3]/div/table/tr[2]/td[8]')))

            debitoDetal = False
            try:
                estabelecimentoElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[2]/td[1]')))
                competenciaElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[2]/td[2]')))
                valorItemElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[2]/td[3]')))
                saldoDevedorElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[2]/td[4]')))
                situacaoElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[2]/td[5]')))
                
                debitoDetal = True
            except TimeoutException:
               pass

            pagDetal = False
            try:
                dataArrecadApropElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[2]/td[1]')))         
                valApropElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[2]/td[2]')))                   
                refenciaElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[2]/td[3]')))                    
                numGuiaElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[2]/td[4]')))                    
                tipoCredElement = WebDriverWait(navegador, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[2]/td[5]')))     

                pagDetal = True               
            except TimeoutException:
                pass
            
            voltarElement = WebDriverWait(navegador, 300).until(
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
            valMultaOficio = valMultaOficioElement.text.strip()
            valMultaIsolada = valMultaIsoladaElement.text.strip()
            valJuros = valJurosElement.text.strip()
            valEncargo = valEncargoElement.text.strip()
            valHonorario = valHonorarioElement.text.strip()
            valTotal = valTotalElement.text.strip()

            dados.append({
                    'Nº_INSCRICAO': numInscricao,
                    'CNPJ': cnpjEmpresa,
                    'NOME': nomeEmpresa,
                    'DT_INSCRICAO': dataInscricao,
                    'RECEITA': receitaDivida,
                    'SITUACAO': SitInscricao,
                    'NATUREZA_DIVIDA': natDivida,

                    'VL_PRINCIPAL': valPrincipal,
                    'VL_MULTA': valMulta,
                    'VL_JUROS': valJuros,
                    'VL_ENCARGO': valEncargo,
                    'VL_TOTAL_CONSOLIDADO': valTotal,
                    'MULTA_DE_OFICIO': valMultaOficio,
                    'MULTA_ISOLADA':valMultaIsolada,
                    'HONORARIO':valHonorario
            })

            if debitoDetal == True:
                y = 2
                while True:
                    try:
                        estabelecimentoElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[{y}]/td[1]')))
                        competenciaElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[{y}]/td[2]')))
                        valorItemElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[{y}]/td[3]')))
                        saldoDevedorElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[{y}]/td[4]')))
                        situacaoElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[4]/div/table/tr[{y}]/td[5]')))
                
                        estabelecimento = estabelecimentoElement.text.strip()
                        competencia = competenciaElement.text.strip()
                        valorItem = valorItemElement.text.strip()
                        saldoDevedor = saldoDevedorElement.text.strip()
                        situacao = situacaoElement.text.strip()
                            
                        #print(f'{estabelecimento} | {competencia} | {valorItem} | {saldoDevedor} | {situacao}')
                    
                        dados.append({
                            'Nº_INSCRICAO': numInscricao,
                            'CNPJ': cnpjEmpresa,
                            'NOME': nomeEmpresa,
                            'DT_INSCRICAO': dataInscricao,
                            'RECEITA': receitaDivida,
                            'SITUACAO': SitInscricao,
                            'NATUREZA_DIVIDA': natDivida,

                            'ESTABELECIMENTO': estabelecimento,
                            'COMPETENCIA': competencia,
                            'VL_ITEM': valorItem,
                            'VL_DEVEDOR': saldoDevedor,
                            'SITUACAO':situacao

                        })

                        y += 2
                    except TimeoutException:
                        break
            
            if pagDetal == True:
                x = 2
                while True:
                    try:
                        dataArrecadApropElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[{x}]/td[1]')))        
                        valApropElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[{x}]/td[2]')))                   
                        refenciaElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[{x}]/td[3]')))                    
                        numGuiaElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[{x}]/td[4]')))                    
                        tipoCredElement = WebDriverWait(navegador, 4).until(
                            EC.presence_of_element_located((By.XPATH, f'/html/body/app-root/div/div[2]/app-debcad/main/tabset/div/tab[2]/fieldset[5]/div/table/tr[{x}]/td[5]'))) 
                            
                        dataArrecadAprop = dataArrecadApropElement.text.replace("\n",":").strip()
                        dataArrecadAprop = dataArrecadAprop.split(sep=":")
                        dataArrecad = dataArrecadAprop[0].strip()
                        dataApropri = dataArrecadAprop[1].strip()
                        valAprop = valApropElement.text.strip()
                        refencia = refenciaElement.text.strip()
                        numGuia = numGuiaElement.text.strip()
                        tipoCred = tipoCredElement.text.strip()
                    
                        dados.append({
                            'Nº_INSCRICAO': numInscricao,
                            'CNPJ': cnpjEmpresa,
                            'NOME': nomeEmpresa,
                            'DT_INSCRICAO': dataInscricao,
                            'RECEITA': receitaDivida,
                            'SITUACAO': SitInscricao,
                            'NATUREZA_DIVIDA': natDivida,

                            'DT_ARRECADACAO': dataArrecad,
                            'DT_APROPRIACAO': dataApropri,
                            'VL_APROPRIADO': valAprop,
                            'REFERENCIA': refencia,
                            'NUM_GUIA': numGuia,
                            'TIPO_CRED': tipoCred

                        })
                        
                        x += 1
                    except TimeoutException:
                        break
    
            adicionar_dados(dados)
            listaExtraidos.append(numInscricao)
            voltarElement.click()

            return numInscricao
        
        case 'Fgts':
            
            nomeEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[1]')))
            cnpjEmpresaElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[2]')))
            numSitInscriElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/div[4]/div[2]')))
            dataInscricaoElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[1]/div[6]')))
            valTotalConsElement = WebDriverWait(navegador, 300).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/div[2]/app-fgts/main/fieldset/div/div[2]/div[2]')))
            voltarElement = WebDriverWait(navegador, 300).until(
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
             
            dados.append({
                    'Nº_INSCRICAO': numInscricao,
                    'CNPJ': cnpjEmpresa,
                    'NOME': nomeEmpresa,
                    'DT_INSCRICAO': dataInscricao,
                    'SITUACAO': SitInscricao,

                    'VL_TOTAL_CONSOLIDADO': valTotalCons,
            })
            
            adicionar_dados(dados)
            listaExtraidos.append(numInscricao)
            voltarElement.click()

            return numInscricao

        case _:
            print("Nenhum valor correspondente foi encontrado")

def renomear_plan():
    verificar_url_PGFN(navegador)
    WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

    abaPerfil = navegador.find_element(By.XPATH, '//*[@id="navbar"]/div/div[1]/div/button')
    navegador.execute_script("arguments[0].click();", abaPerfil)

    campoNome = navegador.find_element(By.XPATH, '//*[@id="userInfoMenu"]/div/div/b')
    nomeEmpresa = campoNome.text

    caminhoPasta = Path(r'C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\PGFN')
    arquivo = r'Relatório - PGFN.xlsx'
    caminhoArquivo = caminhoPasta / arquivo
    try:
        if caminhoArquivo.is_file():
            extensao = caminhoArquivo.suffix
            novo_caminho = caminhoPasta / (f"Relatório - PGFN - {nomeEmpresa}{extensao}")
            caminhoArquivo.rename(novo_caminho)

            print(f"Arquivo {arquivo} renomeado para Relatório PGFN - {nomeEmpresa}{extensao}")
        else:
            print("Arquivo não encontrado")
    except Exception as e:
        print(f"Erro ao renomear arquivo:\n{e}")        

def verificarInscricoes():

    WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

    html_element = navegador.find_element(By.XPATH, '//*[@id="painelConteudo_ATIVA_EM_COBRANCA_TRIBUTARIA_DEMAIS_DEBITOS"]/div[3]/table')
    html_info = html_element.get_attribute('innerHTML')

    soup = BeautifulSoup(html_info, 'html.parser')

    inscricoes = soup.find_all('a', attrs={'title': 'Detalhar'})
    
    listaInscricao = []
    listaInscDuplicadas = []
    totalInscricoes = len(inscricoes)

    count = 1
    for inscricao in inscricoes:
        #print(f'{count}- {inscricao.text}')   

        if inscricao in listaInscricao:
            listaInscDuplicadas.append(inscricao.text)
            pass
        else:
            listaInscricao.append(inscricao.text.strip())

        if count == totalInscricoes:
            break

        count+=1         

    if count == len(listaInscricao):
        print("Não há duplicadas")
    else:
        duplicadas = len(listaInscDuplicadas) 
        print(f"Há {duplicadas} duplicadas em {totalInscricoes} inscrições! \n {listaInscDuplicadas}")

    return []


# ----------------------------------------------------------------------------------------------
print("iniciando...")
time.sleep(2)
verificar_url_PGFN(navegador)
percorrer_inscricoes()
calcular_tempoExecucao(percorrer_inscricoes)
renomear_plan()

