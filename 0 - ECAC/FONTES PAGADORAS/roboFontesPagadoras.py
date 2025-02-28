from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import traceback
import time
import openpyxl
import os
import re

def temp():
    file_path = r'0 - ECAC\FONTES PAGADORAS\temp.xlsx'

    if os.path.exists(file_path):
        os.remove(file_path)

    planTemp = openpyxl.Workbook()
    sheetPlanTemp = planTemp.active

    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    planTemp.save(file_path)
    print(f"Novo arquivo temp criado com sucesso!")

    return planTemp, sheetPlanTemp

planTemp, sheetPlanTemp = temp()


def iniciar_navegador(com_debugging_remoto=True):
    chrome_driver_path = ChromeDriverManager().install()
    service = Service(executable_path=chrome_driver_path)
    
    chrome_options = Options()
    if com_debugging_remoto:
        remote_debugging_port = 9222
        chrome_options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    return navegador

navegador = iniciar_navegador(com_debugging_remoto=True)

navegador.switch_to.default_content()
iframe = navegador.find_element(By.XPATH, '//*[@id="frmApp"]')
navegador.switch_to.frame(iframe)

def renomear_planilha(beneficiario):
    novo_nome = fr'0 - ECAC\FONTES PAGADORAS\Fontes Pagadoras - {beneficiario}.xlsx'
    file_path = r'0 - ECAC\FONTES PAGADORAS\temp.xlsx'
    os.rename(file_path, novo_nome)
    print(f"Planilha renomeada para: {novo_nome}")


def obter_nome_beneficiario(nomeBeneficiario):
    match = re.search(r' - (.+)', nomeBeneficiario)
    if match:
        return match.group(1)
    return nomeBeneficiario

# -----------------------------------------------------------------------------------------------------

def consultar_Ano():
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        anoElement = navegador.find_element(By.XPATH, '//*[@id="opcao_ano"]')
        opcao_ano = Select(anoElement)

        for i in range(10):
            opcao_ano.select_by_index(i)
            ano = opcao_ano.options[i].text
            print(f"Ano selecionado: {ano}")

            time.sleep(0.5)
            btConsultarElement = navegador.find_element(By.XPATH, '//*[@id="btnExecutar"]')
            btConsultarElement.click()

            time.sleep(1)
            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
            avisoElement = navegador.find_element(By.XPATH, '//*[@id="tbmsg_vazio"]/tbody/tr[1]/td[2]/b')

            aviso = avisoElement.text

            if 'não há informações' in aviso:
                print(f'sem info de {ano}')

                
                btVoltarElement = navegador.find_element(By.XPATH, '//*[@id="btnVoltar_msg"]')
                btVoltarElement.click()
                continue
            else:
                expandir_campos()
                beneficiario = percorrer_tabela(ano)

                btVoltarElement = navegador.find_element(By.XPATH, '//*[@id="btnVoltar"]')
                btVoltarElement.click()

        renomear_planilha(beneficiario)
    except TimeoutException:
        print("ERRO ao carregar a pagina!")
        traceback.print_exc()

def expandir_campos():
    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        navegador.switch_to.default_content()
        iframe = navegador.find_element(By.XPATH, '//*[@id="frmApp"]')
        navegador.switch_to.frame(iframe)

        i = 1
        elementos_nao_encontrados = 0
        time.sleep(1)
        while True and i < 100:
            try:
                expandirElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{i}]/td[5]/a')
                expandirElement.click()
                WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
                i += 1
                time.sleep(0.5)
                elementos_nao_encontrados = 0  
            except NoSuchElementException:
                i += 1
                elementos_nao_encontrados += 1
                time.sleep(0.5)
                if elementos_nao_encontrados >= 3:
                    break  
    except TimeoutException:
        print("ERRO ao percorrer a tabela!")
        traceback.print_exc()

def percorrer_tabela(ano):
    navegador.switch_to.default_content()
    iframe = navegador.find_element(By.XPATH, '//*[@id="frmApp"]')
    navegador.switch_to.frame(iframe)

    time.sleep(2)
    beneficiarioElement = navegador.find_element(By.XPATH, '//*[@id="tcab"]/tbody/tr[2]/td[2]/font/b')
    beneficiario = beneficiarioElement.text
    nomeBeneficiario = obter_nome_beneficiario(beneficiario)

    try:
        WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')

        l = 1
        while True:
            try:
                cnpjElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{l}]/td[2]')
                nomeEmpresarialElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{l}]/td[3]/a/span')
                dataEntregaDirfElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{l}]/td[4]')

                if cnpjElement and nomeEmpresarialElement and dataEntregaDirfElement:
                    cnpj = cnpjElement.text
                    nomeEmpresarial = nomeEmpresarialElement.text
                    dataEntregaDirf = dataEntregaDirfElement.text

                c = l + 1
                lc = 3
                while True:
                    try:
                        codigoElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{c}]/td/div/div/table/tbody/tr[{lc}]/td[4]/a')
                        rendimentoTributavelElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{c}]/td/div/div/table/tbody/tr[{lc}]/td[5]')
                        impostoRetidoElement = navegador.find_element(By.XPATH, f'/html/body/div[2]/div[6]/div[2]/table/tbody/tr[{c}]/td/div/div/table/tbody/tr[{lc}]/td[7]')

                        if codigoElement:
                            codigo = codigoElement.text
                            rendimentoTributavel = rendimentoTributavelElement.text
                            impostoRetido = impostoRetidoElement.text
                            print(f'{cnpj} | | {dataEntregaDirf} | {codigo} | {rendimentoTributavel} | {impostoRetido}')
                            adicionar_info_plan(ano, cnpj, nomeEmpresarial, dataEntregaDirf, rendimentoTributavel, impostoRetido, codigo)
                            lc += 1
                        else:
                            break
                    except NoSuchElementException:
                        break
                
                l += 2      
            except NoSuchElementException:
                break

    except TimeoutException:
        print("ERRO ao carregar a pagina!")
        traceback.print_exc()

    return nomeBeneficiario

def adicionar_info_plan(ano, cnpj, nomeEmpresarial, dataEntregaDirf, rendimentoTributavel, impostoRetido, codigo):
    proxima_linha = sheetPlanTemp.max_row + 1


    if proxima_linha == 2:
        sheetPlanTemp['A1'] = "Fonte Pagadora CNPJ/CPF"
        sheetPlanTemp['B1'] = "Nome Empresarial/Nome"
        sheetPlanTemp['C1'] = "Dirf entregue em"
        sheetPlanTemp['D1'] = "Rendimento Tributável"
        sheetPlanTemp['E1'] = "Imposto Retido"
        sheetPlanTemp['F1'] = "Código"
        sheetPlanTemp['G1'] = "Ano Calendário"


    sheetPlanTemp[f'A{proxima_linha}'] = cnpj
    sheetPlanTemp[f'B{proxima_linha}'] = nomeEmpresarial
    sheetPlanTemp[f'C{proxima_linha}'] = dataEntregaDirf
    sheetPlanTemp[f'D{proxima_linha}'] = rendimentoTributavel
    sheetPlanTemp[f'E{proxima_linha}'] = impostoRetido
    sheetPlanTemp[f'F{proxima_linha}'] = codigo
    sheetPlanTemp[f'G{proxima_linha}'] = ano

    planTemp.save(r'0 - ECAC\FONTES PAGADORAS\temp.xlsx')




consultar_Ano()
#expandir_campos()