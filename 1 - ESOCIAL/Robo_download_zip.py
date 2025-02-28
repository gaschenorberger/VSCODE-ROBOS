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
import datetime
from time import sleep
import re

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


from datetime import datetime, time

def tempoSessao():
    contagemTempo_element = navegador.find_element(By.XPATH, '//*[@id="header"]/div[5]')
    contagemTempo_text = contagemTempo_element.text.strip() 
    
    try:
        contagemTempo = datetime.strptime(contagemTempo_text, "%H:%M").time()
        if contagemTempo < time(9, 0):
            navegador.refresh()
            return True  
    except ValueError:
        print("Formato de tempo inválido. Certifique-se de que está no formato HH:MM.")
    return False

def percorrerSolicitacoes():
    try:
        encerrarDownload = False
        for p in range(1, 4):
            if encerrarDownload:
                break

            WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
            botaoSegundPag_element = navegador.find_element(By.XPATH, f'//*[@id="DataTables_Table_0_paginate"]/span/a[{p}]')
            botaoSegundPag_element.click()
            sleep(0.6)

            l = 1
            countFinalizado = 1
            while True:
                if tempoSessao():
                    botaoSegundPag_element = navegador.find_element(By.XPATH, f'//*[@id="DataTables_Table_0_paginate"]/span/a[{p}]')
                    botaoSegundPag_element.click()
                    sleep(0.6)

                WebDriverWait(navegador, 90).until(lambda navegador: navegador.execute_script('return document.readyState') == 'complete')
                numSolicitacao_elements = navegador.find_elements(By.XPATH, f'//*[@id="DataTables_Table_0"]/tbody/tr[{l}]/td[1]')
                    
                if not numSolicitacao_elements:
                    break
                
                #numSolicitacao_elements = navegador.find_elements(By.XPATH, f'//*[@id="DataTables_Table_0"]/tbody/tr[{l}]/td[1]')
                dataIniFim_elements = navegador.find_elements(By.XPATH, f'//*[@id="DataTables_Table_0"]/tbody/tr[{l}]/td[4]')
                situacao_elements = navegador.find_elements(By.XPATH, f'//*[@id="DataTables_Table_0"]/tbody/tr[{l}]/td[5]')

                dataIniFim = dataIniFim_elements[0].text.split()
                numSolicitacao = numSolicitacao_elements[0].text
                dataInicio = dataIniFim[2]
                dataInicio = dataInicio.replace("/","-")
                dataFim = dataIniFim[5]
                dataFim = dataFim.replace("/","-")
                situacao = situacao_elements[0].text

                if "Finalizado" in situacao:
                    countFinalizado += 1

                if "Disponível" in situacao:
                    baixarZip_element = navegador.find_element(By.XPATH, f'//*[@id="DataTables_Table_0"]/tbody/tr[{l}]/td[7]/a')
                    baixarZip_element.click()
                    
                print(f'{l} - {numSolicitacao} | {dataInicio} | {dataFim} || {situacao}')
                l += 1
                if countFinalizado == 12:
                    encerrarDownload = True
                    break

    except TimeoutException:
        print()


percorrerSolicitacoes()
