from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import pyautogui
import pyperclip
from tkinter import *
import time
import os
import random
import mousekey
import numpy as np
import cv2
from datetime import datetime
import pyscreeze
from PIL import Image

# start chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Selenium\ChromeTestProfile"

mkey = mousekey.MouseKey()

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

def procurar_imagem(nome_arquivo, confidence=0.8, region=None, maxTentativas=60, horizontal=0, vertical=0, dx=0, dy=0, acao='clicar', clicks=1, ocorrencia=1, delay_tentativa=1):
    mkey = mousekey.MouseKey()

    def click(x, y):
        pyautogui.click(x, y)

    def doubleClick(x, y):
        pyautogui.doubleClick(x, y)

    def coordenada(x, y):
        print(f'Coordenadas da imagem: ({x}, {y})')

    def moveMouse(x, y, variationx=(-5, 5), variationy=(-5, 5), up_down=(0.2, 0.3), min_variation=-10, max_variation=10, use_every=4, sleeptime=(0.009, 0.019), linear=90):
        mkey.left_click_xy_natural(
            int(x) - random.randint(*variationx),
            int(y) - random.randint(*variationy),
            delay=random.uniform(*up_down),
            min_variation=min_variation,
            max_variation=max_variation,
            use_every=use_every,
            sleeptime=sleeptime,
            print_coords=True,
            percent=linear,
        )

    def clickDrag(x, y, dx, dy):
        pyautogui.moveTo(x, y)
        pyautogui.mouseDown()
        pyautogui.moveTo(x + dx, y + dy, duration=0.5)
        pyautogui.mouseUp()


    acoesValidas = ['clicar', 'mover clicar', 'clicar arrastar']

    if acao not in acoesValidas:
        raise ValueError(f"Ação inválida: '{acao}'. Escolha entre {acoesValidas}.")

    tentativas = 0
    while tentativas < maxTentativas:
        tentativas += 1
        try:
            imag = list(pyautogui.locateAllOnScreen(nome_arquivo, confidence=confidence, region=region))
            
            if imag:
                if len(imag) >= ocorrencia:
                    img = imag[ocorrencia - 1]  
                    x, y = pyautogui.center(img) 
                    x += horizontal
                    y += vertical

                    match acao:
                        case 'clicar':
                            match clicks:
                                case 0:
                                    coordenada(x, y)
                                case 1:
                                    click(x, y)
                                case 2:
                                    doubleClick(x, y)

                        case 'mover clicar':
                            moveMouse(x, y)
                        
                        case 'clicar arrastar':
                            clickDrag(x, y, dx, dy)

                    return True
                else:
                    print(f'A ocorrência {ocorrencia} não foi encontrada.')
                    return False

        except pyscreeze.ImageNotFoundException:
            pass
        time.sleep(delay_tentativa)

    print(f'Imagem não encontrada após {maxTentativas} tentativas.')
    return False

def ajustar_data(data):
    meses = {
        "Janeiro": "01",
        "Fevereiro": "02",
        "Marco": "03",
        "Março": "03",
        "Abril": "04",
        "Maio": "05",
        "Junho": "06",
        "Julho": "07",
        "Agosto": "08",
        "Setembro": "09",
        "Outubro": "10",
        "Novembro": "11",
        "Dezembro": "12"
    }
    
    for mes_nome, mes_numero in meses.items():
        data = data.replace(mes_nome, mes_numero)
    
    return data

def periodo(anoInicio, anoFim):
    navegador.switch_to.default_content()
    iframe = navegador.find_element(By.ID, 'frmApp')
    navegador.switch_to.frame(iframe)

    linha = 2
    countIn = 0
    while True:
        try:
            periodo_element = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{linha}]/td[2]')
            periodoTxtIn = periodo_element.text
            #print(f"anoIn: {periodoTxtIn}")
            barraIn = periodoTxtIn.find("/")
            anoIn = periodoTxtIn[barraIn+1:barraIn+5]

            #print(anoIn)
            if anoInicio == anoIn and countIn == 0:
                linhaAnoInicio = linha
                #print(linhaAnoInicio)
                countIn += 1


            periodoFim_element = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{linha+1}]/td[2]')
            periodoTxtFim = periodoFim_element.text
            #print(f"anoFi: {periodoTxtFim}")
            barraFim = periodoTxtFim.find("/")
            anoFi = periodoTxtFim[barraFim+1:barraFim+5]
            
            if anoIn != anoFi:
                if anoIn == anoFim:
                    linhaAnoFim = linha
                    #print(linhaAnoFim)
                    break

            linha += 1
        except NoSuchElementException:
            if anoFim == "2024":
                linhaAnoFim = linha
            break

    return int(linhaAnoInicio), int(linhaAnoFim)


def percorrer_dctf(anoInicio, anoFim, Ativa="S"):
    #procurar_imagem(r'C:\VS_CODE_MAIN\0 - ECAC\DCTF - PDF\prints\chrome.png', maxTentativas=5, ocorrencia=2)

    linhaAnoInicio, linhaAnoFim = periodo(anoInicio, anoFim)
    print(f"{linhaAnoInicio} - {linhaAnoFim}" )


    i = linhaAnoInicio 
    while i <= linhaAnoFim:
        try:
            navegador.switch_to.default_content()

            iframe = navegador.find_element(By.ID, 'frmApp')
            navegador.switch_to.frame(iframe)

            
            data = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[2]')
            dataRecepcao = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[3]')
            status = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[7]')
            btimprimir = navegador.find_element(By.XPATH, f'//*[@id="tbDeclaracoes"]/tbody/tr[{i}]/td[8]/input[2]')
            TXTdata = data.text
            TXTdata = TXTdata.replace("/","_")

            TXTdataRecep = dataRecepcao.text
            TXTdataRecep = TXTdataRecep.replace("/","_")

            TXTstatus = status.text
            TXTstatus = TXTstatus.replace("/ ","_")

            dataRecepcao


            ano = TXTdata[-4:]
            

            mes = ajustar_data(TXTdata)

            #print(mes, " | ", TXTstatus, " | ", ano, " | ", TXTdataRecep)

            time.sleep(0.5)

            if "Ativa" in TXTstatus:
                nome_arquivo = f"{mes} - 0 - {TXTdataRecep} - {TXTstatus} - {i}"
            elif "Cancelada" in TXTstatus:
                nome_arquivo = f"{mes} - 1 - {TXTdataRecep} - {TXTstatus} - {i}"


            print (nome_arquivo)
                    
            match Ativa:
                case ("S" | "s"):
                    if 'Ativa' in TXTstatus:
                        btimprimir.send_keys(Keys.CONTROL + Keys.RETURN)
                        popup = navegador.switch_to.alert
                        time.sleep(1.5)
                        popup.accept()
                        robo(ano, nome_arquivo, i)

                case ("N" | "n"):
                        btimprimir.send_keys(Keys.CONTROL + Keys.RETURN)
                        popup = navegador.switch_to.alert
                        time.sleep(1.5)
                        popup.accept()
                        robo(ano, nome_arquivo, i)

            i += 1
        except NoSuchElementException:
            break

def robo(ano, nome_arquivo, i):
    abas = navegador.window_handles
    navegador.switch_to.window(abas[1])
    procurar_imagem(r'0 - ECAC\DCTF - PDF\prints\ministeriofaze.png')

    novo_caminho = fr"C:\Users\gabriel.alvise\Desktop\VSCODE-ROBOS\0 - ECAC\DCTF - PDF\PDFs\{ano}"

    if not os.path.exists(novo_caminho):
        os.makedirs(novo_caminho)
    time.sleep(0.5)
    
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(1.5)
    procurar_imagem(r'0 - ECAC\DCTF - PDF\prints\btnSalvar.png')

    caminho_completo = f"{novo_caminho}\\{nome_arquivo}.pdf"
    pyperclip.copy(caminho_completo)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'w')
    navegador.switch_to.window(abas[0])
    navegador.switch_to.default_content()

def calcular_tempo_execucao(func, *args, **kwargs):
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

print()
calcular_tempo_execucao(percorrer_dctf, anoInicio="2019", anoFim="2021")

