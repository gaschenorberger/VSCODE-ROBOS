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
import calendar
from datetime import datetime

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

navegador.switch_to.default_content()
iframe = navegador.find_element(By.XPATH, '/html/body/div[3]/div/iframe')

navegador.switch_to.frame(iframe)

option = navegador.find_element(By.XPATH, '/html/body/font/center/form/table/tbody/tr[2]/td[2]/h2/select')

option.click()
ano = navegador.find_element(By.XPATH, '/html/body/font/center/form/table/tbody/tr[2]/td[2]/h2/select/option[1]')
navegador.execute_script("arguments[0].value = arguments[1];", ano, '2019')



    
meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
for ano in range(2012,2025):
    
    for oMes in meses:
        mes = navegador.find_element(By.XPATH, f'/html/body/font/center/form/table/tbody/tr[3]/td[2]/select/option[contains(text(), "{oMes}")]')
        mes.click()
        time.sleep(3)
        print(ano, oMes)

"""


opcao2019 = navegador.find_element(By.XPATH, '//option[contains(text(), "2019")]')


opcaoDez = navegador.find_element(By.XPATH, '//*[@id="dtFin_inline"]/div/div/div/select[1]/option[contains(text(), "Dez")]')

"""